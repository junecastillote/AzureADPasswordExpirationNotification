Function Invoke-UserPasswordExpirationNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $ConfigFile
    )

    try {
        $users = Get-UserPasswordExpiration -ConfigFile $ConfigFile -ErrorAction Stop
        Send-UserPasswordExpirationNotice -ConfigFile $ConfigFile -InputObject $users -ErrorAction Stop
    }
    catch {
        SayError $_.Exception.Message
        return $null
    }

}