# Exchange Server Health Report

PowerShell script to extract and report Exchange server health statistics.

- [What the script does?](#what-the-script-does)
- [Sample Output](#sample-output)
- [Get-ExchangeServerHealth.ps1](#get-exchangeserverhealthps1)
- [Configuration File Template](#configuration-file-template)
- [Configuration File Settings Explained](#configuration-file-settings-explained)
  - [ReportOptions](#reportoptions)
  - [Thresholds](#thresholds)
  - [MailAndReportParameters](#mailandreportparameters)
  - [Exclusions](#exclusions)
- [Usage Examples](#usage-examples)
  - [Example 1: Running Manually in Exchange Management Shell](#example-1-running-manually-in-exchange-management-shell)

## What the script does?

The script performs several checks on your Exchange Servers like the ones below:

- Server Health (Up Time, Server Roles Services, Mail flow,...)
- Mailbox Database Status (Mounted, Backup, Size, and Space, Mailbox Count, Paths,...)
- Public Folder Database Status (Mount, Backup, Size, and Space,...)
- Database Copy Status
- Database Replication Status
- Mail Queue
- Disk Space
- Server Components

Then an HTML report will be generated and can be sent via email if enabled in the configuration file.

## Sample Output

## Get-ExchangeServerHealth.ps1

The `Get-ExchangeServerHealth.ps1` script accepts the following parameter.

`-ConfigFile` : To specify the PowerShell data file (*.psd1) that contains the configuration for the script.

`-EnableDebug` : Optional switch to start a transcript output to debugLog.txt

## Configuration File Template

The configuration file is an PSD1 file containing the options, thresholds, mail settings, and exclusions that will be used by the script. The snapshot of the configuration file template is shown below:

```powershell
@{
    ReportOptions           = @{
        RunCPUandMemoryReport   = $true
        RunServerHealthReport   = $true
        RunMdbReport            = $true
        RunComponentReport      = $true
        RunPdbReport            = $true
        RunDAGReplicationReport = $true
        RunQueueReport          = $true
        RunDiskReport           = $true
        RunDBCopyReport         = $true
        SendReportViaEmail      = $false
        ReportFile              = "MG_PoshLab_Exchange.html"
    }
    Thresholds              = @{
        LastFullBackup        = 0
        LastIncrementalBackup = 0
        DiskSpaceFree         = 12
        MailQueueCount        = 20
        CopyQueueLenght       = 10
        ReplayQueueLenght     = 10
        CpuUsage              = 60
        RamUsage              = 80
    }
    MailAndReportParameters = @{
        CompanyName = "MG PoshLab"
        MailSubject = "Exchange Service Health Report"
        MailServer  = "mail.mg.poshlab.xyz"
        MailSender  = "MG PostMaster <exchange-Admin@mg.poshlab.xyz>"
        MailTo      = @('june@poshlab.xyz', 'tito.castillote-jr@dxc.com', 'june.castillote@gmail.com')
        MailCc      = @()
        MailBcc     = @()
        SSLEnabled  = $false
        Port        = 25
    }
    Exclusions              = @{
        IgnoreServer     = @()
        IgnoreDatabase   = @()
        IgnorePFDatabase = @()
        IgnoreComponent  = @('ForwardSyncDaemon', 'ProvisioningRps')
    }
}
```

## Configuration File Settings Explained

### ReportOptions

This section can be toggled by changing values with `$true` or `$false`.

- `RunServerHealthReport` - Run test and report the Server Health status
- `RunMdbReport` - Mailbox Database test and report
- `RunComponentReport` - Server Components check
- `RunPdbReport` - For checking the Public Folder database(s)
- `RunDAGReplicationReport` - Check and test replication status
- `RunQueueReport` - Inspect mail queue count
- `RunDiskReport` - Disk space report for each server
- `RunDBCopyReport` - Checking the status of the Database Copies
- `SendReportViaEmail` - Option to send the HTML report via email
- `ReportFile` - File path and name of the HTML Report

### Thresholds

This section defines at which levels the script will report a problem for each check item.

- `LastFullBackup` - age of full backup in days. Setting this to zero (0) will cause the script to ignore this threshold
- `LastIncrementalBackup` - age of incremental backup in days. Setting this to zero (0) will cause the script to ignore this threshold.
- `DiskSpaceFree` - percent (%) of free disk space left
- `MailQueueCount` - Mail transport queue threshold
- `CopyQueueLenght` - CopyQueueLenght threshold for the DAG replication
- `ReplayQueueLenght` - ReplayQueueLenght threshold
- `CpuUsage` - CPU usage threshold %
- `RamUsage` - Memory usage threshold %

### MailAndReportParameters

This section specifies the mail parameters.

- `CompanyName` - The name of the organization or company that you want to appear in the banner of the report.
- `MailSubject` - Subject of the email report.
- `MailServer` - The SMTP Relay server.
- `MailSender` - Mail sender address.
- `MailTo` - Recipient TO address(es).
- `MailCc` - Recipient CC address(es).
- `MailBcc` - Recipient BCC address(es).
- `SSLEnabled` - Turn on or off the SSL connection.
- `Port` - The SMTP server's listening port number.

### Exclusions

This section is where the exclusions can be defined.

- `IgnoreServer` - List of servers to be ignored by the script.
- `IgnoreDatabase` - List of Mailbox Database to be ignored by the script.
- `IgnorePFDatabase` - List of Public Folder Database to be ignored by the script.
- `IgnoreComponent` - List of Server Components to be ignored by the script.

## Usage Examples

> NOTE: Use this script only in Exchange Management Shell session.

### Example 1: Running Manually in Exchange Management Shell

```PowerShell
.\Get-ExchangeServerHealth.ps1 -ConfigFile .\config.psd1 -EnableDebug
```

The above example runs the script using the `config.psd1` configuration file on the same directory. The transcript will be saved to the `debugLog.txt`