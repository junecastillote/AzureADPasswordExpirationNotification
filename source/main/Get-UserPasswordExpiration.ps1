Function Get-UserPasswordExpiration {
    [CmdletBinding()]
    param (

    )

    if (!(Get-MgContext)) {
        SayError "A connection to Microsoft Graph is not found. Run the Connect-MgGraph command first and try again."
        return $null
    }

    ## Get domains list
    try {
        SayInfo "Getting all domains password policies."
        $domains = Get-MgDomain | Sort-Object Id
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
        $domainsList = $domains.id
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
        $properties = "UserPrincipalName", "mail", "displayName", "PasswordPolicies", "LastPasswordChangeDateTime", "CreatedDateTime"
        # $users = Get-MgUser -Filter "userType eq 'member' and accountEnabled eq true" -Property $properties -CountVariable userCount -ConsistencyLevel Eventual -All -PageSize 999 -Verbose | Select-Object $properties | Where-Object {$_.PasswordPolicies -ne 'DisablePasswordExpiration'}
        $users = Get-MgUser -Filter "userType eq 'member' and accountEnabled eq true" `
            -Property $properties -CountVariable userCount `
            -ConsistencyLevel Eventual -All -PageSize 999 -Verbose | `
            Select-Object $properties | Where-Object {
            $_.PasswordPolicies -ne 'DisablePasswordExpiration' -and "$(($_.userPrincipalName).Split('@')[1])" -in $domainsList
        }
    }
    catch {
        SayError 'There was an error getting the list of users in Azure AD. The script will terminate.'
        SayError $_.Exception.Message
        return $null
    }

    ## Add properties to the $users objects
    $users | Add-Member -MemberType NoteProperty -Name Domain -Value $null
    $users | Add-Member -MemberType NoteProperty -Name MaxPasswordAge -Value 0
    $users | Add-Member -MemberType NoteProperty -Name PasswordAge -Value 0
    $users | Add-Member -MemberType NoteProperty -Name ExpiresOn -Value (Get-Date '1970-01-01')
    $users | Add-Member -MemberType NoteProperty -Name DaysRemaining -Value 0
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
        $user.Domain = $userDomain
        $user.maxPasswordAge = $maxPasswordAge
        $user.passwordAge = $passwordAge
        $user.expiresOn = $expiresOn
        $user.daysRemaining = $(
            if (($daysRemaining = (New-TimeSpan -Start $timeNow -End $expiresOn).Days) -lt 1) { 0 }
            else { $daysRemaining }
        )
    }
    return $users
}