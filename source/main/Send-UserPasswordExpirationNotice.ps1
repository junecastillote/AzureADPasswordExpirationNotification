Function Send-UserPasswordExpirationNotice {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $InputObject,

        [Parameter()]
        [switch]
        $NotifyUsers,

        [Parameter()]
        [int[]]
        $PasswordNotificationWindowInDays = @(15, 10, 7, 5, 3, 1) ,

        [Parameter()]
        [string]
        $From,

        [Parameter()]
        [string[]]
        $Attachment,

        [Parameter()]
        [string]
        $Subject = "ATTENTION: Your Office 365 Password Will Expire $expireInDays",

        [Parameter()]
        [string[]]
        $RedirectNotificationTo,

        [Parameter()]
        [string[]]
        $SendReportToAdmins,

        [Parameter()]
        [string[]]
        $CopyHtmlToFolder,

        [Parameter()]
        [string]
        $EmailTemplate
    )

    begin {

        if (@($InputObject | Where-Object { $_.DaysRemaining -in $PasswordNotificationWindowInDays }).Count -lt 1) {
            Say "The input data do not contain users whose passwords expire in $($PasswordNotificationWindowInDays -join ',') days."
            Say "No actions taken. Terminating script."
            Continue
        }

        if (($NotifyUsers -or $SendReportToAdmins -or $RedirectNotificationTo) -and !$From) {
            SayError 'The "From" email address cannot be empty when "NotifyUsers", "SendReportToAdmins", or "RedirectNotificationTo" are enabled.'
            Continue
        }

        if (!(Get-MgContext)) {
            SayError "A connection to Microsoft Graph is not found. Run the Connect-MgGraph command first and try again."
            Continue
        }

        if ($RedirectNotificationTo -and !$NotifyUsers) {
            $NotifyUsers = $true
        }

        if ($SendReportToAdmins) {
            SayInfo "Summary report will be sent to [$($SendReportToAdmins -join ',')]."
        }

        if ($RedirectNotificationTo) {
            SayInfo "RedirectNotificationTo is enabled. All email notifications will be sent to [$($RedirectNotificationTo -join ',')]."
        }

        $ThisFunction = ($MyInvocation.MyCommand)
        $ThisModule = Get-Module ($ThisFunction.Source)
        $ResourceFolder = [System.IO.Path]::Combine((Split-Path ($ThisModule.Path) -Parent), 'resource')

        if (!$EmailTemplate) {
            $EmailTemplate = "$($ResourceFolder)\user_notification_template.html"
        }

        $notificationTemplate = Get-Content $EmailTemplate -Raw

        $organization = Get-MgOrganization

        $todayString = (Get-Date -Format 'yyyyMMddTHHmm')

        ## Functions
        ## JSON email address conversion
        Function ConvertRecipientsToJSON {
            param(
                [Parameter(Mandatory)]
                [string[]]
                $Recipients
            )
            $jsonRecipients = @()
            $Recipients | ForEach-Object {
                $jsonRecipients += @{EmailAddress = @{Address = $_ } }
            }
            return $jsonRecipients
        }

        ## Get attachments and convert into Base64 string
        Function GetAttachments {
            param (
                [Parameter(Mandatory)]
                [string[]]
                $Path
            )
            [System.Collections.ArrayList]$fileAttachment = @()
            foreach ($file in $Path) {
                try {
                    $filename = (Resolve-Path $file -ErrorAction STOP).Path

                    if ($PSVersionTable.PSEdition -eq 'Core') {
                        $fileByte = $([convert]::ToBase64String((Get-Content $filename -AsByteStream)))
                    }
                    else {
                        $fileByte = $([convert]::ToBase64String((Get-Content $filename -Raw -Encoding byte)))
                    }

                    $null = $fileAttachment.Add(
                        $(
                            [PSCustomObject]@{
                                Filename = $filename
                                FileByte = $fileByte
                            }
                        )
                    )
                }
                catch {
                    SayError "Attachment: $($_.Exception.Message)"
                }
            }
            return $fileAttachment
        }

        ## End Functions

        ## Create the temporary CSV file for the report.
        $tempCSV = "$($env:temp)\$($organization.DisplayName)_$($todayString)_User_Password_Expiration.csv"
        $null = New-Item -ItemType File -Path $tempCSV -Force -Confirm:$false

        ## Create the HTML email copy path (if CopyHtmlToFolder is enabled)
        if ($CopyHtmlToFolder) {
            if (!(Test-Path ($CopyHtmlToFolder))) {
                try {
                    $null = New-Item -ItemType Directory -Path ($CopyHtmlToFolder) -Force -Confirm:$false -ErrorAction Stop
                    $HtmlFolder = Resolve-Path ($CopyHtmlToFolder)
                }
                catch {
                    SayError $_.Exception.Message
                    $CopyHtmlToFolder = "$($env:temp)\report"
                    $null = New-Item -ItemType Directory -Path $CopyHtmlToFolder -Force -Confirm:$false
                    # $HtmlFolder = Resolve-Path $CopyHtmlToFolder
                }
            }
            $HtmlFolder = Resolve-Path $CopyHtmlToFolder
            SayInfo "Copies of the HTML messages will be saved in $HtmlFolder"
        }

        # Get attachments
        if ($NotifyUsers -and $Attachment) {
            $fileAttachment = GetAttachments $Attachment
        }
    }
    process {
        foreach ($item in ($InputObject | Where-Object { $_.DaysRemaining -in $PasswordNotificationWindowInDays })) {
            ## If the Days Remaining to Expire is not within the specified days in the config, skip the user.
            ## For example, if the NotifyExpireInDays config value is 15,10,5,3,1 days, the script will only notify users
            ## whose passwords will expire 15,10,5,3,1 days.
            # if ($item.daysRemaining -notin $PasswordNotificationWindowInDays) {
            #     # Skip to the next user
            #     continue
            # }

            ## The HTML template can have these three (3) variables and will be substituted with real values.
            ## * $name = will be replaced with the 'DisplayName' value.
            ## * $expireInDays = will be replaced with the 'daysRemaining' value.
            ## * $upn = will be replaced with the 'userPrincipalName' value.

            $expireMessageString = $(
                if (($item.daysRemaining) -eq 0) {
                    "TODAY"
                }
                else {
                    "in $($item.daysRemaining) days"
                }
            )

            $HTMLMessage = $notificationTemplate.Replace(
                '$name', $($item.displayName)
            ).Replace(
                '$expireInDays', $expireMessageString
            ).Replace(
                '$expirationDate', "$(Get-Date $item.PasswordExpiresOn -Format 'MMMM dd, yyyy')"
            ).Replace(
                '$upn', $item.userPrincipalName
            )

            if ($CopyHtmlToFolder) {
                try {
                    $HTMLMessage | Out-File "$($HtmlFolder)\pwdexp_$($todayString)_$($item.mail).html" -ErrorAction Stop
                }
                catch {
                    SayError "Error saving the HTML email copy to file. $($_.Exception.Message)"
                }
            }

            if ($NotifyUsers) {
                if ($item.Mail) {
                    $mailObject = @{
                        Message                = @{
                            Subject     = $(($Subject).Replace('$expireInDays', $expireMessageString))
                            Body        = @{
                                ContentType = "HTML"
                                Content     = $HTMLMessage
                            }
                            attachments = @()
                        }
                        internetMessageHeaders = @(
                            @{
                                name  = "X-Mailer"
                                value = "AADPwdExpNotif by june.castillote@gmail.com"
                            }
                        )
                        SaveToSentItems        = "false"
                    }

                    if ($fileAttachment.count -gt 0) {
                        foreach ($file in $fileAttachment) {
                            $mailObject.message.attachments += @{
                                "@odata.type"  = "#microsoft.graph.fileAttachment"
                                "name"         = $(Split-Path $file.filename -Leaf)
                                "contentBytes" = $file.fileByte
                            }
                        }
                    }

                    # If the RedirectNotificationTo is specified, all user notifications
                    # are redirected. This is useful when testing notifications. Instead
                    # of sending test notifications to users, you can use this parameter
                    # to send the notification to a test mailbox first.
                    if ($RedirectNotificationTo) {
                        $mailObject.message += @{
                            toRecipients = @(
                                $(ConvertRecipientsToJSON $RedirectNotificationTo)
                            )
                        }
                    }
                    else {
                        # If the RedirectNotificationTo is not specified, the notifications
                        # will be sent to the users whose passwords are about to expire.
                        $mailObject.message += @{
                            ToRecipients = @(
                                @{
                                    EmailAddress = @{
                                        Address = $($item.mail)
                                    }
                                }
                            )
                        }
                    }
                    try {
                        if ($RedirectNotificationTo) {
                            SayInfo "Redirecting password expiration notice for [$($item.displayName)] [Expires in: $($item.daysRemaining) days] [Expires on: $($item.PasswordExpiresOn)]"
                            Send-MgUserMail -UserId $From -BodyParameter $mailObject -ErrorAction Stop
                            $item.Notified = 'No (Redirected)'
                        }
                        else {
                            SayInfo "Sending password expiration notice to [$($item.displayName)] [Expires in: $($item.daysRemaining) days] [Expires on: $($item.PasswordExpiresOn)]"
                            Send-MgUserMail -UserId $From -BodyParameter $mailObject -ErrorAction Stop
                            $item.Notified = 'Yes'
                        }
                    }
                    catch {
                        SayError "There was an error sending the notification to $($item.displayName)"
                        SayError $_.Exception.Message
                        $item.Notified = "No (error)"
                    }
                }
                else {
                    $item.Notified = 'No (no email address)'
                }
            }
            $item | Export-Csv -Path $tempCsv -Append -Force -Confirm:$false -NoTypeInformation
        }
    }
    end {
        if ($SendReportToAdmins) {

            # Get attachments
            $fileAttachment = GetAttachments $tempCSV

            # Summary
            $data = Import-Csv $tempCSV
            $summary = [PSCustomObject]([ordered]@{
                    Total       = @($data).count
                    Notified    = @($data | Where-Object { $_.Notified -eq 'Yes' }).Count
                    NotNotified = @($data | Where-Object { $_.Notified -ne 'Yes' }).Count
                })

            # Get summary template
            $reportTemplate = Get-Content "$($ResourceFolder)\admin_report_template.html" -Raw

            # Compose mail body
            $summaryMessage = $reportTemplate.Replace(
                '$totalCount', $($summary.Total)
            ).Replace(
                '$notifiedCount', $($summary.Notified)
            ).Replace(
                '$notNotifiedCount', $($summary.NotNotified)
            ).Replace(
                '$organizationName', $($organization.DisplayName)
            )

            $mailObject = @{
                Message                = @{
                    Subject     = "Azure AD Password Expiration Report"
                    Body        = @{
                        ContentType = "HTML"
                        Content     = $summaryMessage
                    }
                    attachments = @(

                    )
                }
                internetMessageHeaders = @(
                    @{
                        name  = "X-Mailer"
                        value = "AADPwdExpNotif by june.castillote@gmail.com"
                    }
                )
                SaveToSentItems        = "false"
            }

            foreach ($file in $fileAttachment) {
                $mailObject.message.attachments += @{
                    "@odata.type"  = "#microsoft.graph.fileAttachment"
                    "name"         = $(Split-Path $file.filename -Leaf)
                    "contentBytes" = $file.fileByte
                }
            }

            $mailObject.message += @{
                toRecipients = @(
                    $(ConvertRecipientsToJSON $SendReportToAdmins)
                )
            }

            try {
                SayInfo "Sending password expiration report to [$($SendReportToAdmins -join ",")]"
                Send-MgUserMail -UserId $From -BodyParameter $mailObject
            }
            catch {
                SayError "There was an error sending the report."
                SayError $_.Exception.Message
            }
        }
    }
}