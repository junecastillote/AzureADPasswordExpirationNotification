#Dot-Source all functions
Get-ChildItem -Path $PSScriptRoot\source\*.ps1 -Recurse |
ForEach-Object {
    . $_.FullName
}

$Script:AzAdPwdExpAlertGraphApiToken = $null
$Script:AzAdPwdExpAlertGraphStartDate = '1970-01-01'
$Script:AzAdPwdExpAlertGraphEndDate = '1970-01-01'

#Enable TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12