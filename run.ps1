# Start Exchange PowerShell implicit remoting session
. "$($env:ExchangeInstallPath)\Bin\RemoteExchange.ps1"; Connect-ExchangeServer -auto

# Get this script's current parent folder.
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

<#
In this example, the configuration file used is config.psd1,
which is on the same folder as this script.
#>
& $script_root\Get-ExchangeServerHealth.ps1 -ConfigFile $script_root\config.psd1