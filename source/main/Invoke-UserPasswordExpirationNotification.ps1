Function Invoke-UserPasswordExpirationNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $ConfigFile
    )

    begin {

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

        ## Request a new access token from the Graph API endpoint.
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

        ## Retrieve the access token from the session
        $AccessToken = (Get-AccessToken).access_token
    }
    process {

    }
    end {

    }
}