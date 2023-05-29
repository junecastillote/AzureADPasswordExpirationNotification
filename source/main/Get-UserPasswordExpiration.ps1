Function Get-UserPasswordExpiration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $ConfigFile
    )

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

    ## Get domains list
    try {
        SayInfo "Getting all domains password policies."
        $uri = "https://graph.microsoft.com/v1.0/domains"
        $domains = [System.Collections.ArrayList]@($(Invoke-RestMethod -Method Get -Uri $uri -Headers @{Authorization = "Bearer $AccessToken" } -ContentType 'application/json' -ErrorAction Stop).Value) | Sort-Object Id
        foreach ($domain in $domains) {
            if ($domain) {
                if (!($domain.passwordValidityPeriodInDays)) {
                    # if the passwordValidityPeriodInDays is not set, the default value is 90 days
                    $domain.passwordValidityPeriodInDays = 90
                }

                if (!($domain.passwordNotificationWindowInDays)) {
                    # if the passwordNotificationWindowInDays is not set, the default value is 14 days
                    $domain.passwordNotificationWindowInDays = 14
                }
            }
        }
    }
    catch {
        SayError 'There was an error getting the list domains. The script will terminate.'
        SayError $_.Exception.Message
        return $null
    }

    ## Create a domain lookup table
    $domainTable = [ordered]@{}
    foreach ($item in $domains.SyncRoot) {
        $domainTable.Add($item.id, $item.passwordValidityPeriodInDays)
    }

    ## Get enabled accounts excluding guests
    try {
        SayInfo "Getting all enabled accounts excluding guests and with DisablePasswordExpiration policy assigned."
        $uri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'member' and accountEnabled eq true&`$select=UserPrincipalName,mail,displayName,PasswordPolicies,LastPasswordChangeDateTime&`$top=999&`$count=true"
        $result = (Invoke-RestMethod -Method Get -Uri $uri -Headers @{Authorization = "Bearer $AccessToken"; ConsistencyLevel = 'Eventual' } -ContentType 'application/json' -ErrorAction Stop)
        $users = [System.Collections.ArrayList]@()
        $users.AddRange($($result.value | Where-Object { $_.PasswordPolicies -ne 'DisablePasswordExpiration' } ))
        $totalUserCount = $result.'@odata.count'
        SayInfo "$($users.Count) of $($totalUserCount) users"
        while ($result.'@odata.nextLink') {
            $result = (Invoke-RestMethod -Method Get -Uri $result.'@odata.nextLink' -Headers @{Authorization = "Bearer $AccessToken" } -ContentType 'application/json' -ErrorAction Stop)
            $users.AddRange($($result.value | Where-Object { $_.PasswordPolicies -ne 'DisablePasswordExpiration' } ))
            SayInfo "$($users.Count) of $($totalUserCount) users"
        }
    }
    catch {
        SayError 'There was an error getting the list of users in Azure AD. The script will terminate.'
        SayError $_.Exception.Message
        return $null
    }

    ## Add properties to the $users objects
    $users | Add-Member -MemberType NoteProperty -Name Domain -Value $null
    $users | Add-Member -MemberType NoteProperty -Name maxPasswordAge -Value 0
    $users | Add-Member -MemberType NoteProperty -Name passwordAge -Value 0
    $users | Add-Member -MemberType NoteProperty -Name expiresOn -Value (Get-Date '1970-01-01')
    $users | Add-Member -MemberType NoteProperty -Name daysRemaining -Value 0
    $users | Add-Member -MemberType NoteProperty -Name Notified -Value $null

    ## If the organizational password expiration policy is set to no expiration.
    $maxDaysValue = (([DateTime]::MaxValue) - (Get-Date)).TotalDays

    ## Get the current datetime
    $timeNow = Get-Date

    foreach ($user in $users) {

        $userDomain = ($user.userPrincipalName).Split('@')[1]
        $maxPasswordAge = $domainTable["$($userDomain)"]
        $passwordAge = (New-TimeSpan -Start $user.LastPasswordChangeDateTime -End $timeNow).Days
        # $expiresOn = (Get-Date $user.LastPasswordChangeDateTime).AddDays($maxPasswordAge)
        $expiresOn = $(
            if ($maxPasswordAge -gt $maxDaysValue) {
                ([DateTime]::MaxValue)
            }
            else {
                (Get-Date $user.LastPasswordChangeDateTime).AddDays($maxPasswordAge)
            }
        )
        $daysRemaining = (New-TimeSpan -Start $timeNow -End $expiresOn).Days

        $user.Domain = $userDomain
        $user.maxPasswordAge = $maxPasswordAge
        $user.passwordAge = $passwordAge
        $user.expiresOn = $expiresOn
        $user.daysRemaining = $daysRemaining
    }
    return $users
}