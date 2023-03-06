using namespace System.Management.Automation

class validConsoleColors : IValidateSetValuesGenerator {
    [string[]] GetValidValues() {
        $Values = ([enum]::GetValues([System.ConsoleColor]))
        return $Values
    }
}

Function Say {
    param(
        [Parameter(Mandatory)]
        $Text,
        [Parameter()]
        [ValidateSetAttribute([validConsoleColors])]
        [System.ConsoleColor]
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
        [ValidateSetAttribute([validConsoleColors])]
        [System.ConsoleColor]
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
        [ValidateSetAttribute([validConsoleColors])]
        [System.ConsoleColor]
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
        [ValidateSetAttribute([validConsoleColors])]
        [System.ConsoleColor]
        $Color = 'DarkYellow'
    )
    $Host.UI.RawUI.ForegroundColor = $Color
    "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [WARNING] - $Text" | Out-Default
    [Console]::ResetColor()
}

