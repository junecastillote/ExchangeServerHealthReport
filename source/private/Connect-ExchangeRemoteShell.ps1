Function Connect-ExchangeRemoteShell {
    [CmdletBinding()]
    param ()

    if (!($env:ExchangeInstallPath)) {
        "The Exchange Server Management Tools is not installed on this computer." | Say -Color Red
        return $null
    }

    Remove-PSSnapin 'Microsoft.Exchange.Management.PowerShell.E2010' -Confirm:$false -ErrorAction SilentlyContinue

    $orgConfigCmd = Get-Command Get-OrganizationConfig -ErrorAction SilentlyContinue

    if (!$orgConfigCmd) {
        "Attempting to connect to the Remote Exchange Management Shell." | Say
        . "$($env:ExchangeInstallPath)\Bin\RemoteExchange.ps1"; Connect-ExchangeServer -auto
    }
}