# Get this script's current parent folder.
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

#
Import-Module "$($script_root)\ExchangeServerHealthReport.psd1" -Force


<#
In this example, the configuration file used is demo-config.psd1,
which is on the same folder as this script.
#>
Get-ExchangeServerHealth -ConfigFile $script_root\demo-config.psd1