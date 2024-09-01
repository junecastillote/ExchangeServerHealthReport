# Start Exchange PowerShell implicit remoting session
. "$($env:ExchangeInstallPath)\Bin\RemoteExchange.ps1"; Connect-ExchangeServer -auto

# Run the server health report
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
& $script_root\Get-ExchangeServerHealth.ps1 -ConfigFile $script_root\config.psd1