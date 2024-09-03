Function Disconnect-ExchangeRemoteShell {
    [CmdletBinding()]
    param ()

    Remove-PSSession 'Microsoft.Exchange.Management.PowerShell.E2010' -ErrorAction SilentlyContinue

    $ps_session = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }
    if ($ps_session) {
        $ps_session | ForEach-Object { Remove-Module $_.ComputerName -Force -ErrorAction SilentlyContinue }
        $ps_session | Remove-PSSession -Confirm:$false
    }
}