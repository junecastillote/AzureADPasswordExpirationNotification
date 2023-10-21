# [enum]::GetValues([System.ConsoleColor])
Function Say {
    param(
        [Parameter(Mandatory)]
        $Text,
        [Parameter()]
        $Color = 'Cyan'
    )

    # $originalForegroundColor = $Host.UI.RawUI.ForegroundColor
    if ($Color) {
        $Host.UI.RawUI.ForegroundColor = $Color
    }
    $Text | Out-Default
    [Console]::ResetColor()
}

Function SayError {
    param(
        [Parameter(Mandatory)]
        $Text,
        [Parameter()]
        $Color = 'Red'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [ERROR] - $Text" | Out-Default
    [Console]::ResetColor()
}

Function SayInfo {
    param(
        [Parameter(Mandatory)]
        $Text,
        [Parameter()]
        $Color = 'Green'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [INFO] - $Text" | Out-Default
    [Console]::ResetColor()
}

Function SayWarning {
    param(
        [Parameter(Mandatory)]
        $Text,
        [Parameter()]
        $Color = 'DarkYellow'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [WARNING] - $Text" | Out-Default
    [Console]::ResetColor()
}

