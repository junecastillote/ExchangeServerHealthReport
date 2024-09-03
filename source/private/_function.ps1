#Region Functions
Function LogEnd {
    $txnLog = ""
    Do {
        try {
            Stop-Transcript | Out-Null
        }
        catch [System.InvalidOperationException] {
            $txnLog = "stopped"
        }
    } While ($txnLog -ne "stopped")
}

Function LogStart {
    param (
        [Parameter(Mandatory = $true)]
        [string]$logPath
    )
    LogEnd
    Start-Transcript $logPath -Force | Out-Null
}

Function GetExchangeServerVerion {
    param(
        [string]$ComputerName
    )

    $paramCollection = @{ }

    if ($ComputerName) {
        $paramCollection.Add('ComputerName', $ComputerName)
    }
    $paramCollection.Add(
        'ScriptBlock', {
            $(
                $product_table = @{
                    '14.0' = 'Exchange Server 2010'
                    '14.1' = 'Exchange Server 2010 SP1'
                    '14.2' = 'Exchange Server 2010 SP2'
                    '14.3' = 'Exchange Server 2010 SP3'
                    '15.0' = 'Exchange Server 2013'
                    '15.1' = 'Exchange Server 2016'
                    '15.2' = 'Exchange Server 2019'
                }

                try {
                    $exSetup = Get-Command Exsetup.exe -ErrorAction Stop
                    $exSetup | Add-Member -MemberType NoteProperty -Name ProductName -Value $($product_table["$($exSetup.Version.Major).$($exSetup.Version.Minor)"])
                    $exSetup
                }
                catch {
                    $null
                }
            )
        }

    )

    Invoke-Command @paramCollection
}

Function Say {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,

        [Parameter()]
        [ValidateSet(
            'Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White'
        )]
        [string]$Color
    )
    process {

        if ($Text -eq '') {
            '' | Out-Default
        }
        else {
            if ($Color) {
                $Host.UI.RawUI.ForegroundColor = $Color
            }
            "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : $Text" | Out-Host
        }
        [Console]::ResetColor()
    }
}

#EndRegion Functions