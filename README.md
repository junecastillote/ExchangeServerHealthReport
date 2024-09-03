# Exchange Server Health Report

PowerShell module to extract and report Exchange server health statistics.

> NOTE: This module is based on the previous script called [Get-ExchangeHealth.ps1](https://github.com/junecastillote/Get-ExchangeHealth). I decided to create a new repository for a new script because of many breaking changes made on this one, especially to the configuration file contents and parameters.

- [What the module does?](#what-the-module-does)
- [System Requirements](#system-requirements)
- [Permission Requirements](#permission-requirements)
- [Downloading the Module](#downloading-the-module)
- [Get-ExchangeServerHealth](#get-exchangeserverhealth)
- [Configuration File Template](#configuration-file-template)
- [Configuration File Settings Explained](#configuration-file-settings-explained)
  - [Branding](#branding)
  - [TestItem](#testitem)
  - [Output](#output)
  - [Threshold](#threshold)
  - [Mail](#mail)
  - [Exclusion](#exclusion)
- [Usage Examples](#usage-examples)
  - [Example 1: Manual Run](#example-1-manual-run)
  - [Example 2: Running as a Scheduled Task](#example-2-running-as-a-scheduled-task)
- [Sample HTML Report](#sample-html-report)
- [Current Limitations](#current-limitations)
- [Q and A](#q-and-a)

## What the module does?

The module performs several checks on your Exchange Servers like the ones below:

- Server Health (Up Time, Server Roles Services, Mail flow,...)
- Mailbox Database Status (Mounted, Backup, Size, and Space, Mailbox Count, Paths,...)
- Public Folder Database Status (Mount, Backup, Size, and Space,...) - *Planned for removal in future versions*
- Database Copy Status
- Database Replication Status
- Mail Queue
- Disk Space
- Server Components

Then an HTML report will be generated and can be sent via email if enabled in the configuration file.

## System Requirements

- This module is compatible with Exchange Server versions:
  - 2010 (will be removed in future script releases)
  - 2013
  - 2016
  - 2019
- This module must be run on an Exchange Server with at least one mailbox database active.
- Windows PowerShell 5.1

## Permission Requirements

The account that runs this module must have the following:

- Minimum: Exchange role group membership `View-Only Organization Management`.
- Minimum: Local `administrator` rights on Exchange Server computers.

## Downloading the Module

- Download the latest version from this GitHub repository - [Exchange Server Health Report (main)](https://github.com/junecastillote/ExchangeServerHealthReport).
  - Or you can click this [direct link](https://github.com/junecastillote/Exchange-Server-Health-Report/archive/refs/heads/main.zip) to download the ZIP package.
- Once downloaded, extract the ZIP to your desired folder.For example, to `C:\Scripts\ExchangeServerHealth`
- Finally, import the module by running this command in PowerShell.

  ```PowerShell
  Import-Module C:\Scripts\ExchangeServerHealth\ExchangeServerHealthReport.psd1
  ```

## Get-ExchangeServerHealth

The `Get-ExchangeServerHealth` function accepts the following parameter.

`-ConfigFile` : To specify the PowerShell data file (*.psd1) that contains the configuration for the script. (see [Configuration File Template](#configuration-file-template)).

## Configuration File Template

The configuration file is an `PSD1` file containing the options, thresholds, mail settings, and exclusions that will be used by the script. The sample of the configuration file template is shown below:

```powershell
@{
    Branding  = @{
        Company_Name = 'Organization Name Here'
    }
    TestItem  = @{
        CPU_and_RAM            = $true
        Server_Health          = $true
        Mailbox_Database       = $true
        Server_Component       = $true
        Public_Folder_Database = $false
        DAG_Replication        = $true
        Mail_Queue             = $true
        Disk_Space             = $true
        Database_Copy          = $true
    }
    Output    = @{
        Report_File_Path             = "C:\Scripts\ExchangeServiceHealth\report.html"
        Transcript_File_Path         = "C:\Scripts\ExchangeServiceHealth\transcript.log"
        Enable_Transcript_Logging    = $true
        Append_Timestamp_To_Filename = $false
    }
    Threshold = @{
        Last_Full_Backup_Age_Day        = 7
        Last_Incremental_Backup_Age_Day = 1
        Disk_Space_Free_Percent         = 12
        Mail_Queue_Count                = 20
        Copy_Queue_Length               = 10
        Replay_Queue_Length             = 10
        CPU_Usage_Percent               = 60
        RAM_Usage_Percent               = 80
    }
    Mail      = @{
        Send_Email_Report = $true
        Email_Subject     = "Exchange Service Health Report"
        SMTP_Server       = "mail.server.address.here"
        Sender_Address    = "Exchange Admin <exchange-Admin@domain.tld>"
        To_Address        = @('admin1@domain.tld')
        Cc_Address        = @()
        Bcc_Address       = @()
        SSL_Enabled       = $false
        Port              = 25
    }
    Exclusion = @{
        Ignore_Server_Name      = @('')
        Ignore_MB_Database      = @('DUMMY001')
        Ignore_PF_Database      = @()
        Ignore_Server_Component = @('ForwardSyncDaemon', 'ProvisioningRps')
    }
}
```

## Configuration File Settings Explained

### Branding

- `Company_Name` : The name of the organization or company that you want to appear in the banner of the report

### TestItem

This section list the tests that can be toggled by changing values with `$true` or `$false`.

- `CPU_and_RAM` : Get CPU and RAM usage.
- `Mailbox_Database` : Run mailbox database health checks.
- `Database_Copy` : Databas copy health check for DAG member servers.
- `Public_Folder_Database` : Run public folder database health checks.
- `Server_Health` : Run server health status checks.
- `DAG_Replication` : Run DAG database replication checks.
- `Server_Component` : Get server components status (Exchange 2013+).
- `Mail_Queue` : Get mail queue count.
- `Disk_Space` : Get server disk space statistics.

### Output

- `Report_File_Path` : File path and name of the HTML Report.
- `Transcript_File_Path` : File path and name of the transcript.
- `Enable_Transcript_Logging` : Specify whether to enable transcript logging.
- `Append_Timestamp_To_Filename` : When enabled, the output filename of the report and transcript will have the timestamp appended.
  - eg. If the `Report_File_Path` = `C:\report.html`, the resulting filename on each report is `C:\report_20240804T001023.html` - August 4, 2024, 12:10:23 AM (local time)

### Threshold

This section defines at which levels the script will report a problem for each check item.

- `Last_Full_Backup_Age_Day` : age of full backup in days. Setting this to zero (0) will cause the script to ignore this threshold.
- `Last_Incremental_Backup_Age_Day` : age of incremental backup in days. Setting this to zero (0) will cause the script to ignore this threshold.
- `Disk_Space_Free_Percent` : percent (%) of free disk space left.
- `Mail_Queue_Count` : Mail transport queue threshold.
- `Copy_Queue_Length` : Copy_Queue_Length threshold for the DAG replication.
- `Replay_Queue_Length` : Replay_Queue_Length threshold.
- `CPU_Usage_Percent` : CPU usage threshold %.
- `RAM_Usage_Percent` : Memory usage threshold %.

### Mail

This section specifies the mail parameters.

- `Email_Subject` : Subject of the email report.
- `SMTP_Server` : The SMTP Relay server.
- `Sender_Address` : Mail sender address.
- `To_Address` : Recipient TO address(es).
- `Cc_Address` : Recipient CC address(es).
- `Bcc_Address` : Recipient BCC address(es).
- `SSL_Enabled` : Turn on or off the SSL connection.
- `Port` : The SMTP server's listening port number.

### Exclusion

This section is where the exclusions can be defined.

- `Ignore_Server_Name` : List of servers to be ignored by the script.
- `Ignore_MB_Database` : List of Mailbox Database to be ignored by the script.
- `Ignore_PF_Database` : List of Public Folder Database to be ignored by the script.
- `Ignore_Server_Component` : List of Server Components to be ignored by the script.

## Usage Examples

### Example 1: Manual Run

```PowerShell
# Import the module
Import-Module .\ExchangeServerHealthReport.psd1
# Run the report
Get-ExchangeServerHealth -ConfigFile .\demo-config.psd1
```

This example runs the script using the `demo-config.psd1` configuration file on the same directory.

 ![Example 1](resource/images/Example_01.png)

### Example 2: Running as a Scheduled Task

1. First, create a run file. Below is an example called `run.ps1`.

    > **Note**: Replace the full path of the `$moduleFile` and `$configFile` files as needed.

    ```PowerShell
    # run.ps1

    # Specify the module file path
    $moduleFile = "C:\Scripts\ExchangeServiceHealth\ExchangeServerHealthReport.psd1"

    # Specify the configuration file path
    $configFile = "C:\Scripts\ExchangeServiceHealth\config.psd1"

    # Import the module
    Import-Module $moduleFile -Force

    # Run the report
    Get-ExchangeServerHealth -ConfigFile $configFile

    ```

2. Open the Task Scheduler and create a new task.
3. Specify the name of the new task.
4. Select "Run whether user is logged on or not".

    ![General tab](resource/images/ts_general.png)

5. Go to the **Actions** tab.
6. Add a new action with the following details.
   - Action: Start a program
   - Program/script: `powershell.exe`
   - Add arguments: `-file "Path to run file"` (eg. `-file "C:\Scripts\ExchangeServerhealt\run.ps1"`)

    ![New action](resource/images/ts_action.png)

7. Go to the **Triggers** tab and add your preferred trigger/interval. The below example is a daily trigger at 5AM.

   ![new trigger](resource/images/ts_trigger.png)

8. Save the new task and test it.
9. Review the transcript (if transcript loggin is enabled in the configuration file).
    ![Transcript](resource/images/transcript.png)

## Sample HTML Report

![Sample HTML report](resource/images/sample_html_report.png)

## Current Limitations

- **No SMTP authentication capability**.
  - The built-in logic to send email reports has no feature to use authentication. It assumes that the SMTP relay is anonymous, which is usually the case when using on-premises SMTP relay services that uses IP-based security to allow relay.
  - If you need to use an authenticated SMTP relay, you must implement a separate script that takes the HTML report output of this module and send it separately.
  - There is currently no plan to include this feature.

## Q and A

- **I run this in Task Scheduler but there's no output. How can I validate it?**
  - Ensure that the `Enable_Transcript_Logging` is set to `$true` in the configuration file.
  - Check the transcript log file.
  - Check the error log at `$env:USERPROFILE\ExchangeServerHealth\error.txt`.
    ![error.txt](resource/images/errorLog.png)
- **Do I have to run this in Exchange Management Shell?**
  - No. You can run this in a normal Windows PowerShell window. The `Get-ExchangeServerHealth` command will automatically call the Exchange Management Shell implicit remoting, provided that the Exchange Management Tools is installed on the local machine.
  - But you can also run this inside the Exchange Management Shell and it will bypass the call for implicit remoting.
- **Will this work on PowerShell 7?**
  - No. Stick with Windows PowerShell.
    ![PowerShell Core](resource/images/pscore.png)
- **The HTML report design hurts by eyes. Is there any way to change it?**
  - First, sorry, I'm not a designer. I such as UX.
  - But feel free to take a stab at it by modifying the `resources/style.css` file.
- **Where can I report issues or provide feedback?**
  - You can [open an issue](https://github.com/junecastillote/ExchangeServerHealthReport/issues).
