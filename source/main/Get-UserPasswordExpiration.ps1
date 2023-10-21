Function Get-UserPasswordExpiration {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [Parameter(
            Mandatory,
            ParameterSetName = 'UserId'
        )]
        [String[]]
        $UserId,

        [Parameter(ParameterSetName = 'All')]
        [Switch]
        $IncludeDisabled,

        [Parameter(ParameterSetName = 'All')]
        [Switch]
        $IncludeNoExpiration
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

    try {
        $properties = "Id", "UserPrincipalName", "Mail", "DisplayName", "AccountEnabled", "PasswordPolicies", "LastPasswordChangeDateTime"
        ## Get user
        if ($UserId) {
            $users = $(
                foreach ($id in $UserId) {
                    try {
                        Get-MgUser -UserId $id -Property $properties -ErrorAction Stop | Select-Object $properties
                    }
                    catch {

                    }
                }
            )
        }
        else {

            if ($IncludeDisabled) {
                ## Exclude guests
                $filter = "userType eq 'member'"
            }
            else {
                ## Exclude disabled accounts and guests
                $filter = "userType eq 'member' and accountEnabled eq true"
            }

            ## Get users excluding guests, and including those with non-expiring passwords.
            if ($IncludeNoExpiration) {
                SayInfo "Getting users including those with non-expiring passwords."
                $users = Get-MgUser -Filter $filter `
                    -Property $properties -CountVariable userCount `
                    -ConsistencyLevel Eventual -All -PageSize 999 -Verbose |
                Select-Object $properties | Where-Object {
                    "$(($_.userPrincipalName).Split('@')[1])" -in $domainsList `
                        -and $_.UserPrincipalName -notlike "*#EXT#*"
                }
            }
            ## Get users excluding guests and those with non-expiring passwords.
            else {
                SayInfo "Getting users excluding those with non-expiring passwords and guests."
                $users = Get-MgUser -Filter $filter `
                    -Property $properties -CountVariable userCount `
                    -ConsistencyLevel Eventual -All -PageSize 999 -Verbose |
                Select-Object $properties | Where-Object {
                    $_.PasswordPolicies -notcontains 'DisablePasswordExpiration' `
                        -and $domainTable["$(($_.userPrincipalName).Split('@')[1])"] -ne 2147483647 `
                        -and "$(($_.userPrincipalName).Split('@')[1])" -In $domainsList `
                        -and $_.UserPrincipalName -NotLike "*#EXT#*"
                }
            }
            SayInfo "Retrieved $($users.count) users."
        }
    }
    catch {
        SayError 'There was an error getting the list of users in Azure AD. The script will terminate.'
        SayError $_.Exception
        return $null
    }

    ## Add properties to the $users objects
    $users | Add-Member -MemberType NoteProperty -Name Domain -Value $null
    $users | Add-Member -MemberType NoteProperty -Name MaximumPasswordAge -Value 0
    $users | Add-Member -MemberType NoteProperty -Name CurrentPasswordAge -Value 0
    $users | Add-Member -MemberType NoteProperty -Name PasswordExpiresOn -Value (Get-Date '1970-01-01')
    $users | Add-Member -MemberType NoteProperty -Name DaysRemaining -Value 0
    $users | Add-Member -MemberType NoteProperty -Name PasswordNeverExpires -Value $null
    $users | Add-Member -MemberType NoteProperty -Name PasswordState -Value $null
    $users | Add-Member -MemberType NoteProperty -Name AccountState -Value $null
    $users | Add-Member -MemberType NoteProperty -Name Notified -Value $null


    ## If the organizational password expiration policy is set to no expiration.
    $maxDaysValue = (([DateTime]::MaxValue) - (Get-Date)).Days

    ## Get the current datetime
    $timeNow = Get-Date

    SayInfo "Calculating password expiration."
    foreach ($user in ($users | Sort-Object UserPrincipalName)) {
        $userDomain = ($user.userPrincipalName).Split('@')[1]
        $maximumPasswordAge = $(
            if ($user.PasswordPolicies -contains 'DisablePasswordExpiration') {
                $maxDaysValue
            }
            elseif ($domainTable["$($userDomain)"] -eq 2147483647) {
                $maxDaysValue
            }
            else {
                $domainTable["$($userDomain)"]
            }
        )
        $currentPasswordAge = (New-TimeSpan -Start $user.LastPasswordChangeDateTime -End $timeNow).Days
        $passwordExpiresOn = $(
            if ($user.PasswordPolicies -contains 'DisablePasswordExpiration') {
                ([DateTime]::MaxValue)
            }
            elseif ($domainTable["$($userDomain)"] -eq 2147483647) {
                ([DateTime]::MaxValue)
            }
            else {
                (Get-Date $user.LastPasswordChangeDateTime).AddDays($maximumPasswordAge)
            }
        )
        $user.Domain = $userDomain
        $user.MaximumPasswordAge = $maximumPasswordAge
        $user.CurrentPasswordAge = $currentPasswordAge
        $user.PasswordExpiresOn = $passwordExpiresOn
        $daysRemaining = ((New-TimeSpan -Start $timeNow -End $passwordExpiresOn).Days)
        $user.DaysRemaining = $(
            if ($daysRemaining -lt 1) { 0 }
            else { $daysRemaining }
        )

        $user.PasswordNeverExpires = $(
            if ($user.PasswordPolicies -contains 'DisablePasswordExpiration') {
                "Yes (password policy)"
            }
            elseif ($domainTable["$($userDomain)"] -eq 2147483647) {
                "Yes (domain setting)"
            }
            else {
                "No"
            }
        )

        $user.AccountState = $(
            if ($user.AccountEnabled) {
                'Unlocked'
            }
            else {
                'Locked'
            }
        )

        $user.PasswordState = $(
            if ($user.daysRemaining -lt 1) {
                'Expired'
            }
            else {
                'Current'
            }
        )
    }
    SayInfo "Done."
    $users | Select-Object * -ExcludeProperty PasswordPolicies, AccountEnabled, Domain
}