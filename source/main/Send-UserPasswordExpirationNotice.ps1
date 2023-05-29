Function Send-UserPasswordExpirationNotice {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $InputObject
        ,
        [Parameter(Mandatory)]
        [string]
        $ConfigFile
    )

    begin {
        ## Functions
        Function ReplaceSmartCharacter {
            #https://4sysops.com/archives/dealing-with-smart-quotes-in-powershell/
            param(
                [parameter(Mandatory)]
                [string]$String
            )

            # Unicode Quote Characters
            $unicodePattern = @{
                '[\u2019\u2018]'                                                                                                                       = "'" # Single quote
                '[\u201C\u201D]'                                                                                                                       = '"' # Double quote
                '\u00A0|\u1680|\u180E|\u2000|\u2001|\u2002|\u2003|\u2004|\u2005|\u2006|\u2007|\u2008|\u2009|\u200A|\u200B|\u202F|\u205F|\u3000|\uFEFF' = " " # Space
            }

            $unicodePattern.Keys | ForEach-Object {
                $stringToReplace = $_
                $String = $String -replace $stringToReplace, $unicodePattern[$stringToReplace]
            }

            return $String
        }

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

        ## Import the configuration file
        if (!(Test-Path $ConfigFile)) {
            SayError "The configuration file [$($ConfigFile)] does not exist."
            return $null
        }

        try {
            $settings = (Get-Content $ConfigFile -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
        }
        catch {
            SayError "Failed to import the configuration file."
            SayError $_.Exception.Message
            return $null
        }

        ## Validate NotifyUsers to ensure the From address is not empty.
        $errorList = @()
        if ($settings.NotifyUsers -eq 'True') {
            SayInfo 'User notification is enabled.'
            if (!$settings.SenderAddress) {
                $errorList += 'The Sender email address cannot be empty.'
            }

            if ($errorList.Count -gt 0) {
                $errorList | ForEach-Object {
                    SayError $_
                }
                return $null
            }
            SayInfo "Sender email address is - $($settings.SenderAddress)"
        }


        # Validate NotifyAdmin and AdminRecipients
        $errorList = @()
        if ($settings.SendSummaryToAdmins -eq 'True') {
            SayInfo 'Admin notification is enabled.'
            if (!$settings.SenderAddress) {
                $errorList += 'The Sender email address cannot be empty.'
            }
            if (!$settings.AdminRecipients) {
                $errorList += 'The Admin recipients list cannot be empty.'
            }

            if ($errorList.Count -gt 0) {
                $errorList | ForEach-Object {
                    SayError $_
                }
                return $null
            }
            SayInfo "Admin recipients - $($settings.AdminRecipients -join ',')"
        }

        # Validate RedirectEnabled and RedirectTo
        $errorList = @()
        if ($settings.RedirectAllNotifications -eq 'True' ) {
            SayInfo 'RedirectAllNotifications is enabled for testing / debugging.'
            if (!$settings.SenderAddress) {
                $errorList += 'The Sender email address cannot be empty.'
            }
            if (!$settings.RedirectTo) {
                $errorList += 'The RedirectTo list cannot be empty.'
            }

            if ($errorList.Count -gt 0) {
                $errorList | ForEach-Object {
                    SayError $_
                }
                return $null
            }
            SayInfo "All email notifications will be redirected to - $($settings.RedirectTo -join ',')"
        }

        #

        ## Acquire access token
        try {
            New-AccessToken `
                -ClientID $settings.ApplicationId `
                -ClientCertificate $(Get-Item Cert:\CurrentUser\My\$($settings.CertificateThumbprint)) `
                -TenantID $settings.TenantDomain `
                -ErrorAction Stop
        }
        catch {
            SayError 'There was an error acquiring the graph api access token.'
            SayError $_.Exception.Message
            return $null
        }

        ## Get access token from the memory cache.
        $AccessToken = (Get-AccessToken).access_token

        ## Create the temporary CSV file for the report.
        $tempCSV = "$($env:temp)\$($settings.TenantDomain)_User_Password_Expiration.csv"
        $null = New-Item -ItemType FIle -Path $tempCSV -Force -Confirm:$false

        ## Create the HTML email copy path (if enabled)
        if ($settings.SaveHTMLEmailToFolder) {
            if (!(Test-Path ($settings.SaveHTMLEmailToFolder))) {
                try {
                    $null = New-Item -ItemType Directory -Path ($settings.SaveHTMLEmailToFolder) -Force -Confirm:$false -ErrorAction Stop
                    $tempHtmlFolder = Resolve-Path ($settings.SaveHTMLEmailToFolder)
                }
                catch {
                    SayError $_.Exception.Message
                    $SaveHTMLEmailToFolder = "$($env:temp)\report"
                    $null = New-Item -ItemType Directory -Path $SaveHTMLEmailToFolder -Force -Confirm:$false
                    $tempHtmlFolder = Resolve-Path $SaveHTMLEmailToFolder
                }
            }
            SayInfo "Copies of the HTML messages will be saved in $tempHtmlFolder"
        }

        # $ThisFunction = ($MyInvocation.MyCommand)
        # $ThisModule = Get-Module ($ThisFunction.Source)
        $template = Get-Content $settings.EmailTemplate -Raw

        # Get attachments
        [System.Collections.ArrayList]$attachments = @()
        if ($settings.NotifyUsers -eq 'True' -and $settings.attachments.count -gt 0) {
            foreach ($file in $settings.attachments) {
                try {
                    $filename = (Resolve-Path $file -ErrorAction STOP).Path

                    if ($PSVersionTable.PSEdition -eq 'Core') {
                        $fileByte = $([convert]::ToBase64String((Get-Content $filename -AsByteStream)))
                    }
                    else {
                        $fileByte = $([convert]::ToBase64String((Get-Content $filename -Raw -Encoding byte)))
                    }

                    $null = $attachments.Add(
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
            ## If the Days Remaining to Expire is not within specified days in the config, skip the user.
            ## For example, if the NotifyExpireInDays config value is 15,10,5,3,1 days, the script will only notify users
            ## whose passwords will expire 15,10,5,3,1 days.
            if ($item.daysRemaining -notin $settings.NotifyExpireInDays) {
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

            if ($settings.SaveHTMLEmailToFolder) {
                try {
                    $HTMLMessage | Out-File "$($tempHtmlFolder)\pwdexp_$($item.mail).html" -ErrorAction Stop
                }
                catch {
                    SayError "Error saving the HTML email copy to file. $($_.Exception.Message)"
                }
            }

            if ($settings.NotifyUsers -eq 'True') {
                if ($item.Mail) {
                    $mailObject = @{
                        Message                = @{
                            Subject     = $(($settings.Subject).Replace('$expireInDays', $expireMessageString))
                            Body        = @{
                                ContentType = "HTML"
                                Content     = $(ReplaceSmartCharacter $HTMLMessage)
                            }
                            attachments = @()
                        }
                        internetMessageHeaders = @(
                            @{
                                name  = "X-Mailer"
                                value = "MailboxQuotaNotification by june.castillote@gmail.com"
                            }
                        )
                        SaveToSentItems        = "false"
                    }


                    if ($attachments.count -gt 0) {
                        foreach ($attachment in $attachments) {
                            $mailObject.message.attachments += @{
                                "@odata.type"  = "#microsoft.graph.fileAttachment"
                                "name"         = $(Split-Path $filename -Leaf)
                                "contentBytes" = $fileByte
                            }
                        }
                    }

                    if ($settings.RedirectAllNotifications -eq 'True') {
                        $mailObject.message += @{
                            toRecipients = @(
                                $(ConvertRecipientsToJSON $settings.RedirectTo)
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
                    $mailApiUri = "https://graph.microsoft.com/beta/users/$($settings.SenderAddress)/sendmail"
                    try {
                        SayInfo "Sending password expiration notice to [$($item.displayName)] [Expires in: $($item.daysRemaining) days] [Expires on: $($item.expiresOn)]"
                        $null = Invoke-RestMethod -Method Post -Uri $mailApiUri -Body ($mailObject | ConvertTo-Json -Depth 5) -Headers @{Authorization = "Bearer $AccessToken" } -ContentType application/json -ErrorAction STOP
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
        # if ($settings.SendSummaryToAdmins -eq 'True') {
        #     Sayinfo "Sending summary email report to the administrator recipients."
        #     $ResourceFolder = [System.IO.Path]::Combine((Split-Path ($ThisModule.Path) -Parent), 'resource')
        #     $mailboxQuotaList = Import-Csv -Path $tempCSV
        #     $reportRefreshDate = ([datetime]($mailboxQuotaList[0].'Report Refresh Date')).ToString('MMMM dd, yyyy')
        #     $HTMLMessageToAdmins = Get-Content "$($ResourceFolder)\admin_summary_template.html" -Raw

        #     $totalMailbox = "{0:N0}" -f $mailboxQuotaList.Count
        #     $green = "{0:N0}" -f (($mailboxQuotaList).Where({ $_.'Quota status' -eq 'Good (Under Limits)' }).Count)
        #     $yellow = "{0:N0}" -f (($mailboxQuotaList).Where({ $_.'Quota status' -eq 'Warning Issued' }).Count)
        #     $orange = "{0:N0}" -f (($mailboxQuotaList).Where({ $_.'Quota status' -eq 'Send Disabled' }).Count)
        #     $red = "{0:N0}" -f (($mailboxQuotaList).Where({ $_.'Quota status' -eq 'Send/Receive Disabled' }).Count)

        #     $HTMLMessageToAdmins = $HTMLMessageToAdmins.Replace(
        #         '$totalMailbox', $totalMailbox
        #     ).Replace(
        #         '$date', $reportRefreshDate
        #     ).Replace(
        #         '$green', $green
        #     ).Replace(
        #         '$yellow', $yellow
        #     ).Replace(
        #         '$orange', $orange
        #     ).Replace(
        #         '$red', $red
        #     )

        #     $mailObject = @{
        #         Message                = @{
        #             Subject     = 'Mailbox Quota Summary Report'
        #             Body        = @{
        #                 ContentType = "HTML"
        #                 Content     = $(ReplaceSmartCharacter $HTMLMessageToAdmins)
        #             }
        #             attachments = @()
        #         }
        #         internetMessageHeaders = @(
        #             @{
        #                 name  = "X-Mailer"
        #                 value = "AzAdPasswordNotif by june.castillote@gmail.com"
        #             }
        #         )

        #         SaveToSentItems        = "false"
        #     }

        #     if ($settings.RedirectAllNotifications -eq 'True') {
        #         $mailObject.message += @{
        #             toRecipients = @(
        #                 $(ConvertRecipientsToJSON $settings.RedirectTo)
        #             )
        #         }
        #     }
        #     else {
        #         $mailObject.message += @{
        #             toRecipients = @(
        #                 $(ConvertRecipientsToJSON $settings.AdminRecipients)
        #             )
        #         }
        #     }

        #     if ($PSVersionTable.PSEdition -eq 'Core') {
        #         $fileByte = $([convert]::ToBase64String((Get-Content $tempCSV -AsByteStream)))
        #     }
        #     else {
        #         $fileByte = $([convert]::ToBase64String((Get-Content $tempCSV -Raw -Encoding byte)))
        #     }

        #     $mailObject.message.attachments += @{
        #         "@odata.type"  = "#microsoft.graph.fileAttachment"
        #         "name"         = $(Split-Path $tempCSV -Leaf)
        #         "contentBytes" = $fileByte
        #     }

        #     $mailApiUri = "https://graph.microsoft.com/beta/users/$($settings.SenderAddress)/sendmail"
        #     try {
        #         $null = Invoke-RestMethod -Method Post -Uri $mailApiUri -Body ($mailObject | ConvertTo-Json -Depth 5) -Headers @{Authorization = "Bearer $AccessToken" } -ContentType application/json -ErrorAction STOP
        #     }
        #     catch {
        #         SayError "There was an error sending the summary report."
        #         SayError $_.Exception.Message
        #     }

        # }
    }
}