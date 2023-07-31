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
        $SendSummaryToAdmins,

        [Parameter()]
        [string[]]
        $CopyHtmlToFolder,

        [Parameter()]
        [string]
        $EmailTemplate
    )

    begin {

        if (($NotifyUsers -or $SendSummaryToAdmins -or $RedirectNotificationTo) -and !$From) {
            SayError 'The "From" email address cannot be empty when "NotifyUsers", "SendSummaryToAdmins", or "RedirectNotificationTo" are enabled.'
            Continue;
        }

        if ($RedirectNotificationTo -and $NotifyUsers) {
            SayError 'The "RedirectNotificationTo" parameter cannot be used simultaneously with "NotifyUsers".'
            Continue;
        }

        if (!(Get-MgContext)) {
            SayError "A connection to Microsoft Graph is not found. Run the Connect-MgGraph command first and try again."
            Continue;
        }

        $ThisFunction = ($MyInvocation.MyCommand)
        $ThisModule = Get-Module ($ThisFunction.Source)
        $ResourceFolder = [System.IO.Path]::Combine((Split-Path ($ThisModule.Path) -Parent), 'resource')

        if (!$EmailTemplate) {
            $EmailTemplate = "$($ResourceFolder)\user_notification_template.html"
        }
        $template = Get-Content $EmailTemplate -Raw

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
        ## End Functions

        ## Create the temporary CSV file for the report.
        $tempCSV = "$($env:temp)\$($organization.DisplayName)_User_Password_Expiration.csv"
        $null = New-Item -ItemType FIle -Path $tempCSV -Force -Confirm:$false

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
        [System.Collections.ArrayList]$fileAttachment = @()
        if ($NotifyUsers -and $Attachment) {
            foreach ($file in $Attachment) {
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
        }

    }
    process {
        foreach ($item in $InputObject) {
            ## If the Days Remaining to Expire is not within the specified days in the config, skip the user.
            ## For example, if the NotifyExpireInDays config value is 15,10,5,3,1 days, the script will only notify users
            ## whose passwords will expire 15,10,5,3,1 days.
            if ($item.daysRemaining -notin $PasswordNotificationWindowInDays) {
                # Skip to the next user
                continue
            }

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

            $HTMLMessage = $template.Replace(
                '$name', $($item.displayName)
            ).Replace(
                '$expireInDays', $expireMessageString
            ).Replace(
                '$expirationDate', "$(Get-Date $item.expiresOn -Format 'MMMM dd, yyyy')"
            ).Replace(
                '$upn', $item.userPrincipalName
            )

            if ($CopyHtmlToFolder) {
                try {
                    $HTMLMessage | Out-File "$($HtmlFolder)\pwdexp_$($item.mail).html" -ErrorAction Stop
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

                    if ($RedirectNotificationTo) {
                        $mailObject.message += @{
                            toRecipients = @(
                                $(ConvertRecipientsToJSON $RedirectNotificationTo)
                            )
                        }
                    }
                    else {
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
                        SayInfo "Sending password expiration notice to [$($item.displayName)] [Expires in: $($item.daysRemaining) days] [Expires on: $($item.expiresOn)]"
                        # $null = Invoke-RestMethod -Method Post -Uri $mailApiUri -Body ($mailObject | ConvertTo-Json -Depth 5) -Headers @{Authorization = "Bearer $AccessToken" } -ContentType application/json -ErrorAction STOP
                        Send-MgUserMail -UserId $From -BodyParameter $mailObject
                        $item.Notified = 'Yes'
                    }
                    catch {
                        SayError "There was an error sending the notification to $($item.displayName)"
                        SayError $_.Exception.Message
                        $item.Notified = 'No (error)'
                    }
                }
                else {
                    $item.Notified = 'No (no email address)'
                }
            }
            $item | Export-Csv -Path $tempCsv -Append -Force -Confirm:$false
        }
    }
    end {

    }
}