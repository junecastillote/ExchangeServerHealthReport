Function Get-ExchangeServerHealth {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ConfigFile
    )

    $now = [System.DateTime]::Now

    # Start Script

    #Import Configuration File
    if ((Test-Path $configFile) -eq $false) {
        "ERROR: File $($configFile) does not exist. Script cannot continue" | Say
        "ERROR: File $($configFile) does not exist. Script cannot continue" | Out-File error.txt -Append
        return $null
    }

    $config = Import-PowerShellDataFile $configFile

    #Define Variables
    $module_Info = Get-Module $($MyInvocation.MyCommand.ModuleName)
    # $script_info = Test-ScriptFileInfo $MyInvocation.MyCommand.Definition
    # $script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    $availableTestCount = $config.TestItem.Count
    $enabledTestCount = ($config.TestItem.GetEnumerator() | Where-Object { $_.Value }).Count
    $testFailed = 0
    $testPassed = 0
    $percentPassed = 0
    $overAllResult = "PASSED"
    $errSummary = ""
    $today = '{0:dd-MMM-yyyy hh:mm tt}' -f (Get-Date)
    $css_string = Get-Content (($module_Info.ModuleBase.ToString()) + '\resource\style.css') -Raw

    # Thresholds from config
    [int]$t_Last_Full_Backup_Age_Day = $config.Threshold.Last_Full_Backup_Age_Day
    [int]$t_Last_Incremental_Backup_Age_Day = $config.Threshold.Last_Incremental_Backup_Age_Day
    [double]$t_DiskBadPercent = $config.Threshold.Disk_Space_Free_Percent
    [int]$t_mQueue = $config.Threshold.Mail_Queue_Count
    [int]$t_copyQueue = $config.Threshold.Copy_Queue_Length
    [int]$t_replayQueue = $config.Threshold.Replay_Queue_Length
    [double]$t_CPU_Usage_Percent = $config.Threshold.CPU_Usage_Percent
    [double]$t_RAM_Usage_Percent = $config.Threshold.RAM_Usage_Percent

    # Options from config
    [bool]$CPU_and_RAM = $config.TestItem.CPU_and_RAM
    [bool]$Server_Health = $config.TestItem.Server_Health
    [bool]$Mailbox_Database = $config.TestItem.Mailbox_Database
    [bool]$Server_Component = $config.TestItem.Server_Component
    [bool]$Public_Folder_Database = $config.TestItem.Public_Folder_Database
    [bool]$Database_Copy = $config.TestItem.Database_Copy
    [bool]$DAG_Replication = $config.TestItem.DAG_Replication
    [bool]$Mail_Queue = $config.TestItem.Mail_Queue
    [bool]$Disk_Space = $config.TestItem.Disk_Space

    # Mail settings
    [bool]$Send_Email_Report = $config.Mail.Send_Email_Report
    [string]$Company_Name = $config.Branding.Company_Name
    [string]$Email_Subject = $config.Mail.Email_Subject
    [string]$SMTP_Server = $config.Mail.SMTP_Server
    [int]$Port = $config.Mail.Port
    [bool]$SSL_Enabled = $config.Mail.SSL_Enabled
    [string]$Sender_Address = $config.Mail.Sender_Address
    [string[]]$To_Address = @($config.Mail.To_Address)
    [string[]]$Cc_Address = @($config.Mail.Cc_Address)
    [string[]]$Bcc_Address = @($config.Mail.Bcc_Address)

    # Exclusions
    [string[]]$Ignore_Server_Name = @($config.Exclusion.Ignore_Server_Name)
    [string[]]$Ignore_MB_Database = @($config.Exclusion.Ignore_MB_Database)
    [string[]]$Ignore_PF_Database = @($config.Exclusion.Ignore_PF_Database)
    [string[]]$Ignore_Server_Component = @($config.Exclusion.Ignore_Server_Component)

    # Output
    [bool]$Append_Timestamp_To_Filename = $config.Output.Append_Timestamp_To_Filename

    if ($Append_Timestamp_To_Filename) {
        $timeStamp = $now.ToString('yyyyMMddTHHmmss')

        $reportFilename = [string]($config.Output.Report_File_Path).Replace('.html', "_$($timeStamp).html")
        $Report_File_Path = (New-Item -ItemType File -Path $reportFilename).FullName

        $transcriptFilename = [string]($config.Output.Transcript_File_Path).Replace('.log', "_$($timeStamp).log")
        $Transcript_File_Path = $transcriptFilename
    }
    else {
        $Report_File_Path = (New-Item -ItemType File -Path $config.Output.Report_File_Path -Force).FullName
        $Transcript_File_Path = $config.Output.Transcript_File_Path
    }

    if ($config.Output.Enable_Transcript_Logging) {
        LogStart $Transcript_File_Path
    }

    $hr = "=" * ($module_Info.ProjectUri.OriginalString.Length)
    $hr | Say
    "$($module_Info.Name) v$($module_Info.Version.ToString())" | Say
    "$($module_Info.ProjectUri.OriginalString)" | Say
    $hr | Say

    if ($config.Output.Enable_Transcript_Logging) {
        "Transcript @ $Transcript_File_Path" | Say
    }

    # Check if connected to Exchange Management Shell (Implicit Remoting)

    try {
        Connect-ExchangeRemoteShell
    }
    catch {
        $_
        LogEnd
        return $null
    }

    Function Get-CPUAndMemoryLoad ($exchangeServers) {
        "CPU and Memory Load check..." | Say
        $stats_collection = @()
        $TopProcessCPU = ""
        $tCounter = 0
        foreach ($exchangeServer in $exchangeServers) {
            # Get CPU Usage
            $x = Get-Counter -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -computer $exchangeServer.Name | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue

            Say "     --> Getting CPU Load for $($exchangeServer.Name)"
            $cpuMemObject = "" | Select-Object Server, CPU_Usage, Top_CPU_Consumers, Total_Memory_KB, Memory_Free_KB, Memory_Used_KB, Memory_Free_Percent, Memory_Used_Percent, Top_Memory_Consumers
            $cpuMemObject.Server = $exchangeServer.Name
            $cpuMemObject.CPU_Usage = "{0:N0}" -f ($x.cookedvalue)

            # Get Top 3
            $TopProcessCPU = ""
            $y = Get-Counter '\Process(*)\% Processor Time' -computer $exchangeServer.Name | Select-Object -ExpandProperty countersamples | Where-Object { $_.instancename -ne 'idle' -and $_.instancename -ne '_total' } | Select-Object -Property instancename, cookedvalue | Sort-Object -Property cookedvalue -Descending | Select-Object -First 5
            foreach ($tproc in $y) {
                $z = "$($tproc.instancename) `n"
                #$TopProcessCPU += "$z"
                if ($tCounter -ne ($y.count - 1)) {
                    $TopProcessCPU += "$z"
                    $tCounter = $tCounter + 1
                }
                else {
                    $TopProcessCPU += "$z"
                }
            }
            $cpuMemObject.Top_CPU_Consumers = $TopProcessCPU

            Say "     --> Getting Memory Load for $($exchangeServer.Name)"
            $memObj = Get-CimInstance -ComputerName $exchangeServer.Name -ClassName Win32_operatingsystem -Property CSName, TotalVisibleMemorySize, FreePhysicalMemory
            $cpuMemObject.Total_Memory_KB = $memObj.TotalVisibleMemorySize
            $cpuMemObject.Memory_Free_KB = $memObj.FreePhysicalMemory
            $cpuMemObject.Memory_Used_KB = ($cpuMemObject.Total_Memory_KB - $cpuMemObject.Memory_Free_KB)
            $cpuMemObject.Memory_Used_Percent = "{0:N0}" -f (($cpuMemObject.Memory_Used_KB / $cpuMemObject.Total_Memory_KB) * 100)
            $cpuMemObject.Memory_Free_Percent = "{0:N0}" -f (($cpuMemObject.Memory_Free_KB / $cpuMemObject.Total_Memory_KB) * 100)

            # Get the Top Memory Consumers
            $processes = Get-Process -ComputerName $exchangeServer.Name | Group-Object -Property ProcessName
            $proc_collection = @()
            foreach ($process in $processes) {
                $tempproc = "" | Select-Object Server, ProcessName, MemoryUsed
                $tempproc.ProcessName = $process.Name
                $tempproc.MemoryUsed = (($process.Group | Measure-Object WorkingSet -Sum).sum / 1kb)
                $proc_collection += $tempproc
            }

            $proclist = $proc_collection | Sort-Object MemoryUsed -Descending | Select-Object -First 5

            $TopProcessMemory = ""
            foreach ($proc in $proclist) {
                $topProc = "$($proc.ProcessName) | $($proc.MemoryUsed.ToString('N0')) KB `n"
                $TopProcessMemory += $topProc
            }

            $cpuMemObject.Top_Memory_Consumers = $TopProcessMemory


            $stats_collection += $cpuMemObject

        }
        Return $stats_collection
    }

    # Ping function
    Function Ping-Server ($server) {
        $ping = Test-Connection $server -Quiet -Count 1
        return $ping
    }

    Function Get-MdbStatistic ($mailboxdblist) {
        'Mailbox Database Check... ' | Say
        $stats_collection = @()
        foreach ($mailboxdb in $mailboxdblist) {
            Say "     --> Ping test on $($mailboxdb.Server.Name)"
            if (Ping-Server($mailboxdb.Server.Name) -eq $true) {
                $mdbobj = "" | Select-Object Name, Mounted, MountedOnServer, ActivationPreference, DatabaseSize, AvailableNewMailboxSpace, ActiveMailboxCount, DisconnectedMailboxCount, TotalItemSize, TotalDeletedItemSize, EdbFilePath, LogFolderPath, LogFilePrefix, LastFullBackup, LastIncrementalBackup, BackupInProgress, MapiConnectivity, EDBFreeSpace, LogFreeSpace
                if ($mailboxdb.Mounted -eq $true) {
                    Say "     --> Getting mailbox database statistics on $($mailboxdb.Server.Name)"
                    $mdbStat = Get-MailboxStatistics -Database $mailboxdb
                    $mbxItemSize = $mdbStat | ForEach-Object { $_.TotalItemSize.Value } | Measure-Object -Sum
                    $mbxDelSize = $mdbStat | ForEach-Object { $_.TotalDeletedItemSize.Value } | Measure-Object -Sum
                    $mdbobj.ActiveMailboxCount = ($mdbStat | Where-Object { !$_.DisconnectDate }).count
                    $mdbobj.DisconnectedMailboxCount = ($mdbStat | Where-Object { $_.DisconnectDate }).count
                    $mdbobj.TotalItemSize = "{0:N2}" -f ($mbxItemSize.sum / 1GB)
                    $mdbobj.TotalDeletedItemSize = "{0:N2}" -f ($mbxDelSize.sum / 1GB)
                    $mdbobj.MountedOnServer = $mailboxdb.Server.Name
                    $mdbobj.ActivationPreference = $mailboxdb.ActivationPreference | Where-Object { $_.Key -eq $mailboxdb.Server.Name }
                    $mdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastFullBackup
                    $mdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $mailboxdb.LastIncrementalBackup
                    $mdbobj.BackupInProgress = $mailboxdb.BackupInProgress
                    $mdbobj.DatabaseSize = "{0:N2}" -f ($mailboxdb.DatabaseSize.tobytes() / 1GB)
                    $mdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($mailboxdb.AvailableNewMailboxSpace.tobytes() / 1GB)
                    $mdbobj.MapiConnectivity = Test-MapiConnectivity -Database $mailboxdb.Identity -PerConnectionTimeout 10
                    # Get Disk Details
                    $dbDrive = (Get-CimInstance Win32_LogicalDisk -Computer $mailboxdb.Server.Name | Where-Object { $_.DeviceID -eq $mailboxdb.EdbFilePath.DriveName })
                    $logDrive = (Get-CimInstance Win32_LogicalDisk -Computer $mailboxdb.Server.Name | Where-Object { $_.DeviceID -eq $mailboxdb.LogFolderPath.DriveName })
                    $mdbobj.EDBFreeSpace = "{0:N2}" -f ($dbDrive.Size / 1GB) + ' [' + "{0:N2}" -f ($dbDrive.FreeSpace / 1GB) + ']'
                    $mdbobj.LogFreeSpace = "{0:N2}" -f ($logDrive.Size / 1GB) + ' [' + "{0:N2}" -f ($logDrive.FreeSpace / 1GB) + ']'
                }
                else {
                    $mdbobj.ActiveMailboxCount = "DISMOUNTED"
                    $mdbobj.DisconnectedMailboxCount = "DISMOUNTED"
                    $mdbobj.TotalItemSize = "DISMOUNTED"
                    $mdbobj.TotalDeletedItemSize = "DISMOUNTED"
                    $mdbobj.MountedOnServer = "DISMOUNTED"
                    $mdbobj.ActivationPreference = "DISMOUNTED"
                    $mdbobj.LastFullBackup = "DISMOUNTED"
                    $mdbobj.LastIncrementalBackup = "DISMOUNTED"
                    $mdbobj.BackupInProgress = "DISMOUNTED"
                    $mdbobj.DatabaseSize = "DISMOUNTED"
                    $mdbobj.AvailableNewMailboxSpace = "DISMOUNTED"
                    $mdbobj.MapiConnectivity = "Failed"
                    # Get Disk Details
                    $dbDrive = "DISMOUNTED"
                    $logDrive = "DISMOUNTED"
                    $mdbobj.EDBFreeSpace = "DISMOUNTED"
                    $mdbobj.LogFreeSpace = "DISMOUNTED"
                }
                $mdbobj.Name = $mailboxdb.name
                $mdbobj.EdbFilePath = $mailboxdb.EdbFilePath
                $mdbobj.LogFolderPath = $mailboxdb.LogFolderPath
                $mdbobj.Mounted = $mailboxdb.Mounted
            }
            else {
                $mdbobj = "" | Select-Object Name, Mounted, MountedOnServer, ActivationPreference, DatabaseSize, AvailableNewMailboxSpace, ActiveMailboxCount, DisconnectedMailboxCount, TotalItemSize, TotalDeletedItemSize, EdbFilePath, LogFolderPath, LogFilePrefix, LastFullBackup, LastIncrementalBackup, BackupInProgress, MapiConnectivity, EDBFreeSpace, LogFreeSpace
                $mdbobj.Name = $mailboxdb.name
                $mdbobj.EdbFilePath = $mailboxdb.EdbFilePath
                $mdbobj.LogFolderPath = $mailboxdb.LogFolderPath
                $mdbobj.Mounted = "$($mailboxdb.Server.Name): Connection/Ping Failed"
                $mdbobj.MountedOnServer = "-"
                $mdbobj.ActivationPreference = "-"
                $mdbobj.LastFullBackup = "-"
                $mdbobj.LastIncrementalBackup = "-"
                $mdbobj.BackupInProgress = "-"
                $mdbobj.DatabaseSize = "-"
                $mbxItemSize = "-"
                $mbxDelSize = "-"
                $mdbobj.TotalItemSize = "-"
                $mdbobj.TotalDeletedItemSize = "-"
                $mdbobj.ActiveMailboxCount = "-"
                $mdbobj.DisconnectedMailboxCount = "-"
                $mdbobj.AvailableNewMailboxSpace = "-"
                $mdbobj.MapiConnectivity = "-"
                $mdbobj.EDBFreeSpace = "-"
                $mdbobj.LogFreeSpace = "-"
            }
            $stats_collection += $mdbobj
        }
        return $stats_collection
    }

    Function Get-PdbStatistic ($pfdblist) {
        'Public Folder Database Check... ' | Say
        $stats_collection = @()
        foreach ($pfdb in $pfdblist) {
            $pfdbobj = "" | Select-Object Name, Mounted, MountedOnServer, DatabaseSize, AvailableNewMailboxSpace, FolderCount, TotalItemSize, LastFullBackup, LastIncrementalBackup, BackupInProgress, MapiConnectivity
            $pfdbobj.Name = $pfdb.Name
            $pfdbobj.Mounted = $pfdb.Mounted
            if ($pfdb.Mounted -eq $true) {
                $pfdbobj.MountedOnServer = $pfdb.Server.Name
                $pfdbobj.DatabaseSize = "{0:N2}" -f ($pfdb.DatabaseSize.tobytes() / 1GB)
                $pfdbobj.AvailableNewMailboxSpace = "{0:N2}" -f ($pfdb.AvailableNewMailboxSpace.tobytes() / 1GB)
                $pfdbobj.LastFullBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $pfdb.LastFullBackup
                $pfdbobj.LastIncrementalBackup = '{0:dd-MMM-yyyy hh:mm tt}' -f $pfdb.LastIncrementalBackup
                $pfdbobj.BackupInProgress = $pfdb.BackupInProgress
                $pfdbobj.MapiConnectivity = Test-MapiConnectivity -Database $pfdb.Identity -PerConnectionTimeout 10
            }
            else {
                $pfdbobj.MountedOnServer = "DISMOUNTED"
                $pfdbobj.DatabaseSize = "DISMOUNTED"
                $pfdbobj.AvailableNewMailboxSpace = "DISMOUNTED"
                $pfdbobj.LastFullBackup = "DISMOUNTED"
                $pfdbobj.LastIncrementalBackup = "DISMOUNTED"
                $pfdbobj.BackupInProgress = "DISMOUNTED"
                $pfdbobj.MapiConnectivity = "DISMOUNTED"
            }

            $stats_collection += $pfdbobj
        }
        return $stats_collection
    }

    Function Get-DiskSpaceStatistic ($serverlist) {
        'Disk Space Check... ' | Say
        $stats_collection = @()
        foreach ($server in $serverlist) {
            try {
                Say "     --> Getting fixed disk information on $($server)"
                $diskObj = Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -computer $server | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace
                foreach ($disk in $diskObj) {
                    $serverobj = "" | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace, PercentFree
                    $serverobj.SystemName = $disk.SystemName
                    $serverobj.DeviceID = $disk.DeviceID
                    $serverobj.VolumeName = $disk.VolumeName
                    $serverobj.Size = "{0:N2}" -f ($disk.Size / 1GB)
                    $serverobj.FreeSpace = "{0:N2}" -f ($disk.FreeSpace / 1GB)
                    [int]$serverobj.PercentFree = "{0:N0}" -f (($disk.freespace / $disk.size) * 100)
                    $stats_collection += $serverobj
                }
            }
            catch {
                $serverobj = "" | Select-Object SystemName, DeviceID, VolumeName, Size, FreeSpace, PercentFree
                $serverobj.SystemName = $server
                $serverobj.DeviceID = $disk.DeviceID
                $serverobj.VolumeName = $disk.VolumeName
                $serverobj.Size = 0
                $serverobj.FreeSpace = 0
                [int]$serverobj.PercentFree = 20000
                $stats_collection += $serverobj
            }
        }
        return $stats_collection
    }

    Function Get-ReplicationHealth {
        'Replication Health Check... ' | Say
        $stats_collection = @(Get-MailboxServer | Where-Object { $_.DatabaseAvailabilityGroup -and $_.Name -notin $Ignore_Server_Name } | Sort-Object Name | ForEach-Object {
                Say "     --> Testing replication health on $($_.Name)"
                Test-ReplicationHealth -Identity $_
            })

        return $stats_collection
    }

    Function Get-MailQueueCount ($transportServerList) {
        'Mail Queue Check... ' | Say
        $stats_collection = $transportServerList | Where-Object { $_.Name -notin $Ignore_Server_Name } | Sort-Object Name | ForEach-Object {
            Say "     --> Checking mail queue on $($_.Name)"
            Get-Queue -Server $_ | Where-Object { $_.Identity -notmatch 'Shadow' }
        }

        return $stats_collection
    }

    Function Get-ServerHealth ($serverlist) {
        'Server Status Check... ' | Say
        $stats_collection = @()
        foreach ($server in $serverlist) {
            if (Ping-Server($server.name) -eq $true) {
                # $exchange_product = (AdminDisplayVersionToName -AdminDisplayVersion $server.AdminDisplayVersion)
                Say "     --> Getting Exchange Server version on $($server.name)"

                if ($server.name -ne $($env:computername)) {
                    $exchange_version = GetExchangeServerVerion -ComputerName $server.name
                }
                else {
                    $exchange_version = GetExchangeServerVerion
                }

                Say "     --> Getting Operating System information on $($server.name)"
                $serverOS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $server

                $serverobj = "" | Select-Object Server, ProductName, BuildNumber, KB, Version, Edition, Connectivity, ADSite, UpTime, HubTransportRole, ClientAccessRole, MailboxRole, MailFlow, MessageLatency
                $timespan = ($serverOS.LocalDateTime) - ($serverOS.LastBootUpTime)
                [int]$uptime = "{0:00}" -f $timespan.TotalHours

                $serverobj.Server = $server.Name
                $serverobj.ProductName = $exchange_version.ProductName
                $serverobj.BuildNumber = $exchange_version.version.ToString()
                $serverobj.Edition = $server.Edition
                $serverobj.UpTime = $uptime
                $serverobj.Connectivity = "Passed"
                $serviceStatus = Test-ServiceHealth -Server $server
                $serverobj.HubTransportRole = ""
                $serverobj.ClientAccessRole = ""
                $serverobj.MailboxRole = ""
                $site = ($server.site.ToString()).Split("/")
                $serverObj.ADSite = $site[-1]
                foreach ($service in $serviceStatus) {

                    if ($service.Role -eq 'Hub Transport Server Role') {
                        Say "     --> Testing 'Hub Transport Server Role' on $($server.name)"
                        if ($service.RequiredServicesRunning -eq $true) {
                            $serverobj.HubTransportRole = "Passed"
                        }
                        else {
                            $serverobj.HubTransportRole = "Failed"
                        }
                    }


                    if ($service.Role -eq 'Client Access Server Role') {
                        Say "     --> Testing 'Client Access Server Role' on $($server.name)"
                        if ($service.RequiredServicesRunning -eq $true) {
                            $serverobj.ClientAccessRole = "Passed"
                        }
                        else {
                            $serverobj.ClientAccessRole = "Failed"
                        }
                    }

                    if ($service.Role -eq 'Mailbox Server Role') {
                        Say "     --> Testing 'Mailbox Server Role' on $($server.name)"
                        if ($service.RequiredServicesRunning -eq $true) {
                            $serverobj.MailboxRole = "Passed"
                        }
                        else {
                            $serverobj.MailboxRole = "Failed"
                        }
                    }
                }
                #Mail Flow
                if ($server.serverrole -match 'Mailbox') {
                    if ($server.name -in $activeServers) {
                        Say "     --> Testing mail flow on $($server.name)"
                        $mailflowresult = $null
                        $result = Test-MailFlow -TargetMailboxServer $server.Name
                        $mailflowresult = $result.TestMailflowResult
                        $serverObj.MailFlow = $mailflowresult
                    }
                    else {
                        Say "     --> Skipping mail flow test on $($server.name) because it has no active mailbox databases."
                        $mailflowresult = $null
                        # $result = Test-MailFlow -TargetMailboxServer $server.Name
                        $mailflowresult = 'NotApplicable'
                        $serverObj.MailFlow = $mailflowresult
                    }
                }
            }
            else {
                $serverobj = "" | Select-Object Server, Connectivity, ADSite, UpTime, HubTransportRole, ClientAccessRole, MailboxRole

                $site = ($server.site.ToString()).Split("/")
                $serverObj.ADSite = $site[-1]
                $serverobj.Server = $server.Name
                $serverobj.Connectivity = "Failed"
                $serverobj.UpTime = "Cannot retrieve up time"
                $serverobj.HubTransportRole = "Failed"
                $serverobj.ClientAccessRole = "Failed"
                $serverobj.MailboxRole = "Failed"
                $serverObj.MailFlow = "Failed"
                $serverObj.MessageLatency = "Failed"
            }
            $stats_collection += $serverobj
        }
        return $stats_collection
    }

    Function Get-ServerHealthReport ($serverhealthinfo) {
        'Server Health Report... ' | Say
        $mResult = "<tr><td>Server Health Status</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $mbody += '<table id="SectionLabels"><tr><th class="data">Server Health Status</th></tr></table>'
        $mbody += '<table id="data">'
        $mbody += '<tr><th>Server Name</th><th>Product</th><th>Site</th><th>Connectivity</th><th>Up Time (Hours)</th><th>Hub Transport Role</th><th>Client Access Role</th><th>Mailbox Role</th><th>Mail Flow</th></tr>'
        foreach ($server in $serverhealthinfo) {
            $mbody += "<tr><td>$($server.server)</td><td>Name: $($server.ProductName)<br/>Build: $($server.BuildNumber)<br/>Edition: $($server.Edition)</td><td>$($server.ADSite)</td>"
            # Uptime
            if ($server.UpTime -lt 24) {
                $mbody += "<td class = ""good"">$($server.Connectivity)</td><td class = ""bad"">$($server.UpTime)</td>"
            }
            elseif ($server.Uptime -eq 'Cannot retrieve up time') {
                $errString += "<tr><td>Server Connectivity [$($server.server)]</td></td><td>$($server.server) - connection test failed. SERVER MIGHT BE DOWN!!!</td></tr>"
                $mbody += "<td class = ""bad"">$($server.Connectivity)</td><td class = ""bad"">$($server.UpTime)</td>"
            }
            else {
                $mbody += "<td class = ""good"">$($server.Connectivity)</td><td class = ""good"">$($server.UpTime)</td>"
            }
            # Transport Role
            if ($server.HubTransportRole -eq 'Passed') {
                $mbody += '<td class = "good">Passed</td>'
            }
            elseif ($server.HubTransportRole -eq 'Failed') {
                $errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Hub Transport Role services are running</td></tr>"
                $mbody += '<td class = "bad">Failed</td>'
            }
            else {
                $mbody += '<td class = "good"></td>'
            }
            # CAS Role
            if ($server.ClientAccessRole -eq 'Passed') {
                $mbody += '<td class = "good">Passed</td>'
            }
            elseif ($server.ClientAccessRole -eq 'Failed') {
                $errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Client Access Role services are running</td></tr>"
                $mbody += '<td class = "bad">Failed</td>'
            }
            else {
                $mbody += '<td class = "good"></td>'
            }
            # Mailbox Role
            if ($server.MailboxRole -eq 'Passed') {
                $mbody += '<td class = "good">Passed</td>'
            }
            elseif ($server.MailboxRole -eq 'Failed') {
                $errString += "<tr><td>Role Services</td></td><td>$($server.server) - not all required Mailbox Role services are running</td></tr>"
                $mbody += '<td class = "bad">Failed</td>'
            }
            else {
                $mbody += '<td class = "good"></td>'
            }

            # Mail Flow
            if ($server.MailFlow -eq "Failed") {
                $errString += "<tr><td>Mail Flow</td></td><td>$($db.Name) - Mail Flow Result FAILED</td></tr>"
                $mbody += '<td class = "bad">Failed</td>'
            }
            elseif ($server.MailFlow -eq 'Success') {
                $mbody += '<td class = "good">Success</td>'
            }
            else {
                $mbody += '<td class = "good">' + $server.MailFlow + '</td>'
            }
            $mbody += '</tr>'
        }
        if ($errString) { $mResult = "<tr><td>Server Health Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-DatabaseCopyStatus ($mailboxdblist) {
        'Mailbox Database Copy Status Check... ' | Say
        $stats_collection = @()

        foreach ($db in $mailboxdblist) {
            if ($db.DatabaseCopies.Count -lt 2) {
                continue
            }
            foreach ($dbCopy in ($db.DatabaseCopies | Where-Object { $_.HostServerName -notin $Ignore_Server_Name })) {
                Say "     --> Getting database copy status of $($dbCopy.Identity.ToString())"
                $temp = "" | Select-Object Name, Status, CopyQueueLength, LogCopyQueueIncreasing, ReplayQueueLength, LogReplayQueueIncreasing, ContentIndexState, ContentIndexErrorMessage
                $dbStatus = Get-MailboxDatabaseCopyStatus -Identity $dbCopy
                $temp.Name = $dbStatus.Name
                $temp.Status = $dbStatus.Status
                $temp.CopyQueueLength = $dbStatus.CopyQueueLength
                $temp.LogCopyQueueIncreasing = $dbStatus.LogCopyQueueIncreasing
                $temp.ReplayQueueLength = $dbStatus.ReplayQueueLength
                $temp.LogReplayQueueIncreasing = $dbStatus.LogReplayQueueIncreasing
                if ($db.IndexEnabled -eq $false) {
                    $temp.ContentIndexState = "Disabled"
                    $temp.ContentIndexErrorMessage = $dbStatus.ContentIndexErrorMessage
                }
                else {
                    $temp.ContentIndexState = $dbStatus.ContentIndexState
                    $temp.ContentIndexErrorMessage = $dbStatus.ContentIndexErrorMessage
                }
                $stats_collection += $temp
            }

        }

        return $stats_collection | Sort-Object Name
    }

    Function Get-DAGCopyStatusReport ($mdbCopyStatus) {
        'Mailbox Database Copy Status... ' | Say
        $mResult = "<tr><td>Mailbox Database Copy Status</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $mbody += '<table id="SectionLabels"><tr><th class="data">Mailbox Database Copy Status</th></tr></table>'
        $mbody += '<table id="data">'
        $mbody += '<tr><th>Name</th><th>Status</th><th>CopyQueueLength</th><th>LogCopyQueueIncreasing</th><th>ReplayQueueLength</th><th>LogReplayQueueIncreasing</th><th>ContentIndexState</th><th>ContentIndexErrorMessage</th></tr>'

        foreach ($mdbCopy in $mdbCopyStatus) {

            $mbody += "<tr><td>$($mdbCopy.Name)</td>"

            # Status
            if ($mdbCopy.Status -eq 'Mounted' -or $mdbCopy.Status -eq 'Healthy') {
                $mbody += "<td class = ""good"">$($mdbCopy.Status)</td>"
            }
            else {
                $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - Status is [$($mdbCopy.Status)]</td></tr>"
                $mbody += "<td class = ""bad"">$($mdbCopy.Status)</td>"
            }
            # CopyQueueLength
            if ($mdbCopy.CopyQueueLength -ge $t_copyQueue) {
                $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - CopyQueueLength [$($mdbCopy.CopyQueueLength)] is >= $($t_copyQueue)</td></tr>"
                $mbody += "<td class = ""bad"">$($mdbCopy.CopyQueueLength)</td>"
            }
            else {
                $mbody += "<td class = ""good"">$($mdbCopy.CopyQueueLength)</td>"
            }
            # LogCopyQueueIncreasing
            if ($mdbCopy.LogCopyQueueIncreasing -eq $true) {
                $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - LogCopyQueueIncreasing</tr>"
                $mbody += "<td class = ""bad"">$($mdbCopy.LogCopyQueueIncreasing)</td>"
            }
            else {
                $mbody += "<td class = ""good"">$($mdbCopy.LogCopyQueueIncreasing)</td>"
            }
            # ReplayQueueLength
            if ($mdbCopy.ReplayQueueLength -ge $t_replayQueue) {
                $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - ReplayQueueLength [$($mdbCopy.CopyQueueLength)] is >= $($t_replayQueue)</td></tr>"
                $mbody += "<td class = ""bad"">$($mdbCopy.ReplayQueueLength)</td>"
            }
            else {
                $mbody += "<td class = ""good"">$($mdbCopy.ReplayQueueLength)</td>"
            }
            # LogReplayQueueIncreasing
            if ($mdbCopy.LogReplayQueueIncreasing -eq $true) {
                $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - LogReplayQueueIncreasing</tr>"
                $mbody += "<td class = ""bad"">$($mdbCopy.LogReplayQueueIncreasing)</td>"
            }
            else {
                $mbody += "<td class = ""good"">$($mdbCopy.LogReplayQueueIncreasing)</td>"
            }
            # ContentIndexState
            if ($mdbCopy.ContentIndexState -eq "Healthy") {
                $mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"
            }
            elseif ($mdbCopy.ContentIndexState -eq "Disabled") {
                $mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"
            }
            elseif ($mdbCopy.ContentIndexState -eq "NotApplicable") {
                $mbody += "<td class = ""good"">$($mdbCopy.ContentIndexState)</td>"
            }
            else {
                $errString += "<tr><td>Database Copy</td></td><td>$($mdbCopy.Name) - ContentIndexState is $($mdbCopy.ContentIndexState)</tr>"
                $mbody += "<td class = ""bad"">$($mdbCopy.ContentIndexState)</td>"
            }
            # ContentIndexErrorMessage
            $mbody += "<td class = ""bad"">$($mdbCopy.ContentIndexErrorMessage)</td>"
        }
        $mbody += '</tr>'
        if ($errString) { $mResult = "<tr><td>Mailbox Database Copy Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-ExServerComponent ($exServerList) {
        'Server Component State... ' | Say
        foreach ($exServer in $exServerList) {
            Say "     --> Getting server component state on $($exServer)"
            $stats_collection += (Get-ServerComponentState $exServer | Where-Object { $_.Component -notin $Ignore_Server_Component } | Select-Object Identity, Component, State)
        }
        Say "     --> Ignored server component: $($Ignore_Server_Component -join ';')"

        return $stats_collection
    }

    Function Get-QueueReport ($queueInfo) {
        'Mail Queue Report... ' | Say
        $mResult = "<tr><td>Mail Queue</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $currentServer = ""
        $mbody += '<table id="SectionLabels"><tr><th class="data">Mail Queue</th></tr></table>'
        $mbody += '<table id="data">'

        foreach ($queue in $queueInfo) {
            $xq = $queue.Identity.ToString()
            $transportServer = $xq.split("\")
            if ($currentServer -ne $transportServer[0]) {
                $currentServer = $transportServer[0]
                $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Delivery Type</th><th>Status</th><th>Message Count</th><th>Next Hop Domain</th><th>Last Error</th></tr>'
            }

            if ($queue.MessageCount -ge $t_mQueue) {
                $errString += "<tr><td>Mail Queue</td></td><td>$($transportServer[0]) - $($queue.Identity) - Message Count is >= $($t_mQueue)</td></tr>"
                $mbody += "<tr><td>$($queue.Identity)</td><td>$($queue.DeliveryType)</td><td>$($queue.Status)</td><td class = ""bad"">$($queue.MessageCount)</td><td>$($queue.NextHopDomain)</td><td>$($queue.LastError)</td></tr>"
            }
            else {
                $mbody += "<tr><td>$($queue.Identity)</td><td>$($queue.DeliveryType)</td><td>$($queue.Status)</td><td>$($queue.MessageCount)</td><td>$($queue.NextHopDomain)</td><td>$($queue.LastError)</td></tr>"
            }

        }
        if ($errString) { $mResult = "<tr><td>Mail Queue</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-ReplicationReport ($replInfo) {
        'Replication Health Report... ' | Say
        $mResult = "<tr><td>DAG Members Replication</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $currentServer = ""
        $mbody += '<table id="SectionLabels"><tr><th class="data">DAG Members Replication</th></tr></table>'
        $mbody += '<table id="data">'

        foreach ($repl in $replInfo) {
            if ($currentServer -ne $repl.Server) {
                $currentServer = $repl.Server
                $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Result</th><th>Error</th></tr>'
                # $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Result</th><th>Error</th><th>Notes</th></tr>'
            }

            if ($repl.Error) {
                # Remove leading spaces by splitting an re-joining
                $replError = (($repl.Error).replace('Failures:', $null).split("`n") -replace '^\s+', '') -join "`n"
                # $replError = (($repl.Error).replace('Failures:', $null).split("`n") -replace '^\s+', '' | Select-Object -Skip 1) -join "`n"
                # Split by single blank line
                $replError = [regex]::Split($replError, '(?<=\S)\r?\n\r?\n(?=\S)')


                $replErrorMessage = @()
                foreach ($item in $replError) {
                    # If the error message doesn't match the ignored database name, add it to the $replErrorMessage collection
                    if (!($Ignore_MB_Database | Where-Object { $item -match $_ }) ) {
                        $replErrorMessage += $item
                    }
                }

                if ($replErrorMessage.Count -gt 0) {
                    $mbody += "<tr><td>$($repl.Check)</td><td class = ""bad"">$($repl.Result.ToString())</td><td>$((($replErrorMessage) -join '<br>==============================<br>').Replace("`n","<br>"))</td></tr>"
                    $errString += "<tr><td>Replication</td></td><td> [$($currentServer)] - $($repl.Check) is $($repl.Result.ToString()) - $((($replErrorMessage) -join '<br>==============================<br>').Replace("`n","<br>"))</td></tr>"
                }
                else {
                    $mbody += "<tr><td>$($repl.Check)</td><td class = ""good"">$($repl.Result.ToString())</td><td></td></tr>"
                }
            }
            else {
                $mbody += "<tr><td>$($repl.Check)</td><td class = ""good"">$($repl.Result.ToString())</td><td></td></tr>"
            }
        }
        $mbody += ""

        if ($errString) { $mResult = "<tr><td>DAG Members Replication</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-ServerComponentStateReport ($serverComponentStateInfo) {

        'Server Component State... ' | Say
        $mResult = "<tr><td>Server Component State</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $currentServer = ""
        $mbody += '<table id="SectionLabels"><tr><th class="data">Server Component State</th></tr></table>'
        $mbody += '<table id="data">'

        foreach ($componentInfo in $serverComponentStateInfo) {
            if ($currentServer -ne $componentInfo.Identity) {
                $currentServer = $componentInfo.Identity
                $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Component State</th></tr>'
            }

            [string]$componentName = $componentInfo.Component
            [string]$componentState = $componentInfo.State

            if ($componentState -ne 'Active') {
                $errString += "<tr><td>Component State</td></td><td>$($currentServer) - $($componentName) [$($componentState)]</td></tr>"
                $mbody += "<tr><td>$($componentName)</td><td class = ""bad"">$($componentState)</td></tr>"
            }
            else {
                $mbody += "<tr><td>$($componentName)</td><td class = ""good"">$($componentState)</td></tr>"
            }
        }
        if ($errString) { $mResult = "<tr><td>Server Component State</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-DiskReport ($diskinfo) {
        'Disk Space Report... ' | Say
        $mResult = "<tr><td>Disk Space</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $currentServer = ""
        $mbody += '<table id="SectionLabels"><tr><th class="data">Disk Space</th></tr></table>'
        $mbody += '<table id="data">'
        foreach ($diskdata in $diskinfo) {
            if ($currentServer -ne $diskdata.SystemName) {
                $currentServer = $diskdata.SystemName
                $mbody += '<tr><th><b><u>' + $currentServer + '</b></u></th><th>Size (GB)</th><th>Free (GB)</th><th>Free (%)</th></tr>'
            }

            if ($diskdata.PercentFree -eq 20000) {
                $errString += "<tr><td>Disk</td></td><td>$($currentServer) - Error Fetching Data</td></tr>"
                $mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""bad"">Error Fetching Data</td></tr>"
            }
            elseif ($diskdata.PercentFree -ge $t_DiskBadPercent) {
                $mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""good"">$($diskdata.PercentFree)</td></tr>"
            }
            else {
                $errString += "<tr><td>Disk</td></td><td>$($currentServer) - $($diskdata.DeviceID) [$($diskdata.VolumeName)] [$($diskdata.FreeSpace) GB / $($diskdata.PercentFree)%] is <= $($t_DiskBadPercent)% Free</td></tr>"
                $mbody += "<tr><td>$($diskdata.DeviceID) [$($diskdata.VolumeName)] </td><td>$($diskdata.Size)</td><td>$($diskdata.FreeSpace)</td><td class = ""bad"">$($diskdata.PercentFree)</td></tr>"
            }

        }
        if ($errString) { $mResult = "<tr><td>Disk Space </td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-MdbReport ($dblist) {
        'Mailbox Database Report... ' | Say
        $mResult = "<tr><td>Mailbox Database Status</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody = @()
        $errString = @()
        $mbody += '<table id="SectionLabels"><tr><th class="data">Mailbox Database Status</th></tr></table>'
        $mbody += '<table id="data"><tr><th>[Name][EDB Path][Log Path]</th><th>Mounted</th><th>On Server [Preference]</th><th>EDB Disk Size [Free] <br /> Log Disk Size [Free]</th><th>Size (GB)</th><th>White Space (GB)</th><th>Active Mailbox</th><th>Disconnected Mailbox</th><th>Item Size (GB)</th><th>Deleted Items Size (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>Mapi Connectivity</th></tr>'
        ForEach ($db in $dblist) {
            if ($db.mounted -eq $true) {
                # Calculate backup age----------------------------------------------------------
                if ($db.LastFullBackup) {
                    $LastFullBackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastFullBackup
                    $LastFullBackupElapsed = New-TimeSpan -Start $db.LastFullBackup
                }
                Else {
                    $LastFullBackupElapsed = ''
                    $LastFullBackup = '[NO DATA]'
                }

                if ($db.LastIncrementalBackup) {
                    $LastIncrementalBackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastIncrementalBackup
                    $LastIncrementalBackupElapsed = New-TimeSpan -Start $db.LastIncrementalBackup
                }
                Else {
                    $LastIncrementalBackupElapsed = ''
                    $LastIncrementalBackup = '[NO DATA]'
                }

                if ($t_Last_Full_Backup_Age_Day -eq 0) {
                    [int]$full_backup_age = -1
                }
                else {
                    [int]$full_backup_age = $LastFullBackupElapsed.totaldays
                }

                if ($t_Last_Incremental_Backup_Age_Day -eq 0) {
                    [int]$incremental_backup_age = -1
                }
                else {
                    [int]$incremental_backup_age = $LastIncrementalBackupElapsed.totaldays
                }
                #-------------------------------------------------------------------------------
                $mbody += '<tr>'
                $mbody += '<td>[' + $db.Name + ']<br />[' + $db.EdbFilePath + ']<br />[' + $db.LogFolderPath + ']</td>'
                if ($db.Mounted -eq $true) {
                    $mbody += '<td class = "good">' + $db.Mounted + '</td>'
                }
                Else {
                    $errString += "<tr><td>Database Mount</td></td><td>$($db.Name) - is NOT MOUNTED</td></tr>"
                    $mbody += '<td class = "bad">' + $db.Mounted + '</td>'
                }

                if ($db.ActivationPreference.Value -eq 1) {
                    $mbody += '<td class = "good">' + $db.MountedOnServer + ' [' + $db.ActivationPreference.value + ']' + '</td>'
                }
                Else {
                    $errString += "<tr><td>Database Activation</td></td><td>$($db.Name) - is mounted on $($db.MountedOnServer) which is NOT the preferred active server</td></tr>"
                    $mbody += '<td class = "bad">' + $db.MountedOnServer + ' [' + $db.ActivationPreference.value + ']' + '</td>'
                }

                $mbody += '<td>' + $db.EDBFreeSpace + '<br />' + $db.LogFreeSpace + '</td>'
                $mbody += '<td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td><td>' + $db.ActiveMailboxCount + '</td><td>' + $db.DisconnectedMailboxCount + '</td><td>' + $db.TotalItemSize + '</td><td>' + $db.TotalDeletedItemSize + '</td>'

                if ($full_backup_age -gt $t_Last_Full_Backup_Age_Day) {
                    $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$LastFullBackup] is OLDER than $($t_Last_Full_Backup_Age_Day) Day(s)</td></tr>"
                    $mbody += '<td class = "bad">' + $LastFullBackup + '</td>'
                }
                elseif ($LastFullBackup -eq '[NO DATA]' -and $t_Last_Full_Backup_Age_Day -ne 0) {
                    $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$LastFullBackup] is OLDER than $($t_Last_Full_Backup_Age_Day) Day(s)</td></tr>"
                    $mbody += '<td class = "bad">' + $LastFullBackup + '</td>'
                }
                Else {
                    $mbody += '<td class = "good">' + $LastFullBackup + '</td>'
                }

                if ($incremental_backup_age -gt $t_Last_Incremental_Backup_Age_Day) {
                    $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date [$LastIncrementalBackup] is OLDER than $($t_Last_Incremental_Backup_Age_Day) Day(s)</td></tr>"
                    $mbody += '<td class = "bad">' + $LastIncrementalBackup + '</td>'
                }
                Else {
                    $mbody += '<td class = "good"> ' + $LastIncrementalBackup + '</td>'
                }

                $mbody += '</td><td>' + $db.BackupInProgress + '</td>'

                if ($db.MapiConnectivity.Result.Value -eq 'Success') {
                    $mbody += '<td class = "good"> ' + $db.MapiConnectivity.Result.Value + '</td>'
                }
                else {
                    $errString += "<tr><td>MAPI Connectivity</td></td><td>$($db.Name) - MAPI Connectivity Result is $($db.MapiConnectivity.Result.Value)</td></tr>"
                    $mbody += '<td class = "bad"> ' + $db.MapiConnectivity.Result.Value + '</td>'
                }
            }
            else {
                $errString += "<tr><td>Mailbox Datababase</td></td><td>$($db.Name) is DISMOUNTED</td></tr>"
                $mbody += "<tr><td>$($db.Name)</td><td class = ""bad"">$($db.Mounted)</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td></tr>"
            }
            $mbody += '</tr>'
        }
        if ($errString) { $mResult = "<tr><td>Mailbox Database Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-PdbReport ($dblist) {
        'Public Folder Database Report... ' | Say
        $mResult = "<tr><td>Public Folder Database Status</td><td class = ""good"">Passed</td></tr>"
        $testFailed = 0
        $mbody += '<table id="SectionLabels"><tr><th class="data">Public Folder Database</th></tr></table>'
        $mbody += '<table id="data"><tr><th>Name</th><th>Mounted</th><th>On Server</th><th>Size (GB)</th><th>White Space (GB)</th><th>Full Backup</th><th>Incremental Backup</th><th>Backup In Progress</th><th>MAPI Connectivity</th></tr>'
        ForEach ($db in $dblist) {
            if ($db.Mounted -eq $true) {
                #Calculate backup age----------------------------------------------------------
                if ($db.LastFullBackup) {
                    $LastFullBackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastFullBackup
                    $LastFullBackupElapsed = New-TimeSpan -Start $db.LastFullBackup
                }
                Else {
                    $LastFullBackupElapsed = ''
                    $LastFullBackup = '[NO DATA]'
                }

                if ($db.LastIncrementalBackup) {
                    $LastIncrementalBackup = '{0:dd/MM/yyyy hh:mm tt}' -f $db.LastIncrementalBackup
                    $LastIncrementalBackupElapsed = New-TimeSpan -Start $db.LastIncrementalBackup
                }
                Else {
                    $LastIncrementalBackupElapsed = ''
                    $LastIncrementalBackup = '[NO DATA]'
                }
                [int]$full_backup_age = $LastFullBackupElapsed.totaldays
                [int]$incremental_backup_age = $LastIncrementalBackupElapsed.totaldays
                #-------------------------------------------------------------------------------
                $mbody += '<tr>'
                $mbody += '<td>' + $db.Name + '</td>'
                if ($db.Mounted -eq $true) {
                    $mbody += '<td class = "good">' + $db.Mounted + '</td>'
                }
                Else {
                    $errString += "<tr><td>Database Mount</td></td><td>$($db.Name) - is NOT MOUNTED</td></tr>"
                    $mbody += '<td class = "bad">' + $db.Mounted + '</td>'
                }

                $mbody += '<td>' + $db.MountedOnServer + '</td><td>' + $db.DatabaseSize + '</td><td>' + $db.AvailableNewMailboxSpace + '</td>'

                if ($full_backup_age -gt $t_Last_Full_Backup_Age_Day) {
                    $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last full backup date [$LastFullBackup] is OLDER than $($t_Last_Full_Backup_Age_Day) days</td></tr>"
                    $mbody += '<td class = "bad">' + $LastFullBackup + '</td>'
                }
                Else {
                    $mbody += '<td class = "good">' + $LastFullBackup + '</td>'
                }

                if ($incremental_backup_age -gt $t_Last_Incremental_Backup_Age_Day) {
                    $errString += "<tr><td>Database Backup</td></td><td>$($db.Name) - last incremental backup date [$LastIncrementalBackup] is OLDER than $($t_Last_Incremental_Backup_Age_Day) days</td></tr>"
                    $mbody += '<td class = "bad">' + $LastIncrementalBackup + '</td>'
                }
                Else {
                    $mbody += '<td class = "good"> ' + $LastIncrementalBackup + '</td>'
                }
                $mbody += '</td><td>' + $db.BackupInProgress + '</td>'

                if ($db.MapiConnectivity.Result.Value -eq 'Success') {
                    $mbody += '<td class = "good"> ' + $db.MapiConnectivity.Result.Value + '</td>'
                }
                else {
                    $mbody += '<td class = "bad"> ' + $db.MapiConnectivity.Result.Value + '</td>'
                }
            }
            else {
                $errString += "<tr><td>Public Folder Datababase</td></td><td>$($db.Name) is DISMOUNTED</td></tr>"
                $mbody += "<tr><td>$($db.Name)</td><td class = ""bad"">$($db.Mounted)</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td><td>DISMOUNTED</td></tr>"
            }

            $mbody += '</tr>'
        }
        if ($errString) { $mResult = "<tr><td>Public Folder Database Status</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    Function Get-CPUAndMemoryReport ($cpuAndMemDataResult) {
        'CPU and Memory Report... ' | Say
        $mResult = "<tr><td>CPU and Memory Usage</td><td class = ""good"">Passed</td></tr>"
        $mbody = @()
        $errString = @()
        $testFailed = 0
        $currentServer = ""
        $mbody += '<table id="SectionLabels"><tr><th class="data">CPU and Memory Load</th></tr></table>'
        $mbody += '<table id="data">'

        foreach ($cpuAndMemData in $cpuAndMemDataResult) {
            $Top_CPU_Consumers = $cpuAndMemData.Top_CPU_Consumers -replace "`n", "<br />"
            $Top_Memory_Consumers = $cpuAndMemData.Top_Memory_Consumers -replace "`n", "<br />"

            if ($currentServer -ne $cpuAndMemData.Server) {
                $currentServer = $cpuAndMemData.Server
                $mbody += '<tr><th>Server Name</th><th>CPU Load</th><th>CPU Top Processes</th><th>Memory Load</th><th>Memory Top Processes</th></tr>'
            }

            if ([int]$cpuAndMemData.CPU_Usage -lt $t_CPU_Usage_Percent) {
                $mbody += "<tr><td>$($currentServer)</td><td class = ""good"">$($cpuAndMemData.CPU_Usage)%</td><td>$($Top_CPU_Consumers)</td>"
            }
            elseif ([int]$cpuAndMemData.CPU_Usage -ge $t_CPU_Usage_Percent) {
                $mbody += "<tr><td>$($currentServer)</td><td class = ""bad"">$($cpuAndMemData.CPU_Usage)%</td><td>$($Top_CPU_Consumers)</td>"
                $errString += "<tr><td>CPU</td></td><td>$($currentServer) - $($cpuAndMemData.CPU_Usage)% CPU Load IS OVER the $($t_CPU_Usage_Percent)% threshold </td></tr>"
            }

            if ([int]$cpuAndMemData.Memory_Used_Percent -lt $t_RAM_Usage_Percent) {
                $mbody += "<td class = ""good"">$($cpuAndMemData.Memory_Used_Percent)%</td><td>$($Top_Memory_Consumers)</td></tr>"
            }
            elseif ([int]$cpuAndMemData.Memory_Used_Percent -ge $t_RAM_Usage_Percent) {
                $errString += "<td>Memory</td></td><td>$($currentServer) - $($cpuAndMemData.Memory_Used_Percent)% Memory Load IS OVER the $($t_RAM_Usage_Percent)% threshold </td></tr>"
                $mbody += "<td class = ""bad"">$($cpuAndMemData.Memory_Used_Percent)%</td><td>$($Top_Memory_Consumers)</td></tr>"
            }
        }

        if ($errString) { $mResult = "<tr><td>CPU and Memory Usage</td><td class = ""bad"">Failed</td></tr>" ; $testFailed = 1 }

        return $mbody, $errString, $mResult, $testFailed
    }

    # SCRIPT BEGIN---------------------------------------------------------------

    # Get-List of Exchange Servers and assign to array----------------------------
    'Building List of Servers - excluding Edge' | Say
    $temp_ExServerList = Get-ExchangeServer | Where-Object { $_.ServerRole -notmatch 'Edge' } | Sort-Object Name
    $dagMemberCount = Get-MailboxServer | Where-Object { $_.DatabaseAvailabilityGroup }
    if (!$dagMemberCount) { $dagMemberCount = @() }

    # Get rid of excluded Servers
    $ExServerList = @()
    foreach ($ExServer in $temp_ExServerList) {
        if ($Ignore_Server_Name -notcontains $ExServer.Name) {
            $exServerList += $ExServer
        }
    }

    Say "     --> Exchange Servers: $($ExServerList -join ';')"
    Say "     --> Ignored Servers: $($Ignore_Server_Name -join ';')"

    $nonEx2010 = $ExServerList | Where-Object { $_.AdminDisplayVersion -notlike "Version 14*" }
    $nonEx2010transportServers = @()
    $nonEx2010transportServers += $ExServerList | Where-Object { $_.AdminDisplayVersion -notlike "Version 14*" -and $_.ServerRole -match 'Mailbox' }
    $Ex2010TransportServers = @()
    $Ex2010TransportServers += $ExServerList | Where-Object { $_.AdminDisplayVersion -like "Version 14*" -and $_.ServerRole -match 'HubTransport' }
    $transportServers = @()
    $transportServers += $nonEx2010transportServers + $Ex2010TransportServers
    # ----------------------------------------------------------------------------
    # Get-List of Mailbox Database and assign to array----------------------------
    if ($Mailbox_Database -eq $true -OR $Database_Copy -eq $true) {
        'Building List of Mailbox Database' | Say
        $temp_ExMailboxDBList = Get-MailboxDatabase -Status | Where-Object { $_.Recovery -eq $False -and $_.Server -notin $Ignore_Server_Name }
        #Get rid of excluded Mailbox Database
        $ExMailboxDBList = @()
        $activeServers = @()
        foreach ($ExMailboxDB in $temp_ExMailboxDBList) {
            if ($Ignore_MB_Database -notcontains $ExMailboxDB.Name) {
                $ExMailboxDBList += $ExMailboxDB
                $activeServers += ($ExMailboxDB.MountedOnServer).Split(".")[0]
            }
        }
        Say "     --> Mailbox Database: $($ExMailboxDBList -join ';')"
        Say "     --> Ignored Database: $($Ignore_MB_Database -join ';')"
        $activeServers = $activeServers | Select-Object -Unique
    }
    # ----------------------------------------------------------------------------
    # Get-List of Public Folder Database and assign to array----------------------
    if ($Public_Folder_Database -eq $true) {
        'Building List of Public Folder Database' | Say
        $temp_ExPFDBList = Get-PublicFolderDatabase -Status | Where-Object { $_.Recovery -eq $False }
        if (!$temp_ExPFDBList) { $temp_ExPFDBList = @() }
        $ExPFDBList = @()

        # Get rid of excluded PF Database
        foreach ($ExPFDB in $temp_ExPFDBList) {
            if ($Ignore_PF_Database -notcontains $ExPFDB.Name) {
                $ExPFDBList += $ExPFDB
            }
        }
    }
    #----------------------------------------------------------------------------

    # Begin Data Extraction-------------------------------------------------------
    $hr | Say
    'Begin Data Extraction' | Say
    if ($CPU_and_RAM -eq $true) { $cpuHealthData = Get-CPUAndMemoryLoad($ExServerList) ; }
    if ($Server_Health -eq $true) { $serverhealthdata = Get-ServerHealth($ExServerList) ; }
    if ($Server_Component -eq $true -AND $nonEx2010.count -gt 0) { $componentHealthData = Get-ExServerComponent ($nonEx2010) ; }
    if ($Mailbox_Database -eq $true) { $mdbdata = Get-MdbStatistic ($ExMailboxDBList) | Sort-Object Name ; }
    if ($Public_Folder_Database -eq $true -AND $ExPFDBList.Count -gt 0) { $pdbdata = Get-PdbStatistic ($ExPFDBList) ; }
    else {
        # $enabledTestCount--
        Say 'Public Folder Database Check... '
        Say '     --> No public folder databases found.'
    }
    if ($Database_Copy -eq $true) { $dagCopyData = Get-DatabaseCopyStatus ($ExMailboxDBList) ; }
    if ($DAG_Replication -eq $true -and $dagMemberCount.count -gt 0) { $repldata = Get-ReplicationHealth ; }
    if ($Mail_Queue -eq $true) { $queueData = Get-MailQueueCount ($transportServers) ; }
    if ($Disk_Space -eq $true) { $diskdata = Get-DiskSpaceStatistic($ExServerList) ; }

    # ----------------------------------------------------------------------------
    # Build Report --------------------------------------------------------------
    $hr | Say
    'Create Report' | Say
    if ($CPU_and_RAM -eq $true) {
        $cpuAndMemoryCheckResult, $cpuError, $cpuResult, $cpuFailed = Get-CPUAndMemoryReport ($cpuHealthData)
        $errSummary += $cpuError
        $testFailed += $cpuFailed
    }
    if ($Server_Health -eq $true) { $serverhealthreport, $sError, $sResult, $sFailed = Get-ServerHealthReport ($serverhealthdata) ; $errSummary += $sError; $testFailed += $sFailed }
    if ($Server_Component -eq $true -AND $nonEx2010.count -gt 0) { $componentHealthReport, $cError, $cResult, $cFailed = Get-ServerComponentStateReport ($componentHealthData) ; $errSummary += $cError; $testFailed += $cFailed }
    if ($Mailbox_Database -eq $true) { $mdbreport, $mError, $mdbResult, $mdbFailed = Get-MdbReport ($mdbdata) ; $errSummary += $mError; $testFailed += $mdbFailed }
    if ($Database_Copy -eq $true) { $dbcopyreport, $dbCopyError, $dbResult, $dbFailed = Get-DAGCopyStatusReport ($dagCopyData) ; $errSummary += $dbCopyError; $testFailed += $dbFailed }
    if ($DAG_Replication -eq $true -and $dagMemberCount.count -gt 0) { $replicationreport, $rError, $rResult, $rFailed = Get-ReplicationReport ($repldata) ; $errSummary += $rError; $testFailed += $rFailed }
    if ($Public_Folder_Database -eq $true -AND $ExPFDBList.Count -gt 0) { $pdbreport, $pdbError, $pdbResult, $pdbFailed = Get-PdbReport ($pdbdata) ; $errSummary += $pdbError; $testFailed += $pdbFailed }
    if ($Mail_Queue -eq $true) { $queuereport, $qError, $qResult, $qFailed = Get-QueueReport($queueData) ; $errSummary += $qError; $testFailed += $qFailed }
    if ($Disk_Space -eq $true) { $diskreport, $dError, $dResult, $dFailed = Get-DiskReport ($diskdata) ; $errSummary += $dError; $testFailed += $dFailed }

    $mail_body = "<html><head><title>[$($Company_Name)] $($Email_Subject) $($today)</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
    'Formatting Report...' | Say
    $mail_body += '<style type="text/css">'
    $mail_body += $css_string
    $mail_body += '</style></head><body>'
    $mail_body += '<table id="HeadingInfo">'
    $mail_body += '<tr><th>' + $Company_Name + '<br />' + $Email_Subject + '<br />' + $today + '</th></tr>'
    $mail_body += '</table>'

    # Set Individual Test Results
    $testPassed = $enabledTestCount - $testFailed
    $percentPassed = ($testPassed / $enabledTestCount) * 100
    $percentPassed = [math]::Round($percentPassed)
    if ($testPassed -lt $enabledTestCount) { $overAllResult = "FAILED" }

    $mail_body += '<table id="SectionLabels">'
    $mail_body += "<tr><th class=""data"">Overall Health: $($percentPassed)% - $($overAllResult)</th></tr></table>"
    $mail_body += '<table id="data"><tr><th>Test Items</th><th>Result</th></tr>'
    if ($CPU_and_RAM -eq $true) { $mail_body += $cpuResult }
    if ($Server_Health -eq $true) { $mail_body += $sResult }
    if ($Server_Component -eq $true -AND $nonEx2010.count -gt 0) { $mail_body += $cResult }
    if ($Mailbox_Database -eq $true) { $mail_body += $mdbResult }
    if ($Database_Copy -eq $true) { $mail_body += $dbResult }
    if ($DAG_Replication -eq $true -and $dagMemberCount.count -gt 0) { $mail_body += $rResult }
    if ($Public_Folder_Database -eq $true -AND $ExPFDBList.Count -gt 0) { $mail_body += $pdbResult }
    if ($Mail_Queue -eq $true) { $mail_body += $qResult }
    if ($Disk_Space -eq $true) { $mail_body += $dResult }
    $mail_body += '</table>'
    if ($overAllResult -eq 'FAILED') {
        $mail_body += '<table id="SectionLabels">'
        $mail_body += '<tr><th class="data">Issues</th></tr></table>'
        $mail_body += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'
        $mail_body += $errSummary
        $mail_body += '</table>'
    }

    if ($CPU_and_RAM -eq $true) { $mail_body += $cpuAndMemoryCheckResult ; $mail_body += '</table>' }
    if ($Server_Health -eq $true) { $mail_body += $serverhealthreport ; $mail_body += '</table>' }
    if ($Server_Component -eq $true -AND $nonEx2010.count -gt 0) { $mail_body += $componentHealthReport ; $mail_body += '</table>' }
    if ($Mailbox_Database -eq $true) { $mail_body += $mdbreport ; $mail_body += '</table>' }
    if ($DAG_Replication -eq $true) { $mail_body += $replicationreport ; $mail_body += '</table>' }
    if ($Database_Copy -eq $true) { $mail_body += $dbcopyreport ; $mail_body += '</table>' }
    if ($Public_Folder_Database -eq $true) { $mail_body += $pdbreport ; $mail_body += '</table>' }
    if ($Mail_Queue -eq $true) { $mail_body += $queuereport ; $mail_body += '</table>' }
    if ($Disk_Space -eq $true) { $mail_body += $diskreport ; $mail_body += '</table>' }
    $mail_body += '<p><table id="SectionLabels">'
    $mail_body += '<tr><th>----END of REPORT----</th></tr></table></p>'
    # $mail_body += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
    $mail_body += '<p><font size="2" face="Tahoma"><br />'
    $mail_body += '<b>[THRESHOLD]</b><br />'
    $mail_body += 'Last Full Backup: ' + $t_Last_Full_Backup_Age_Day + ' Day(s)<br />'
    $mail_body += 'Last Incremental Backup: ' + $t_Last_Incremental_Backup_Age_Day + ' Day(s)<br />'
    $mail_body += 'Mail Queue: ' + $t_mQueue + '<br />'
    $mail_body += 'Copy Queue: ' + $t_copyQueue + '<br />'
    $mail_body += 'Replay Queue: ' + $t_replayQueue + '<br />'
    $mail_body += 'Disk Space Critical: ' + $t_DiskBadPercent + ' (%) <br />'
    $mail_body += 'CPU: ' + $t_CPU_Usage_Percent + ' (%) <br />'
    $mail_body += 'Memory: ' + $t_RAM_Usage_Percent + ' (%) <br />'

    if ($Send_Email_Report) {
        $mail_body += '<br /><b>[MAIL]</b><br />'
        $mail_body += 'SMTP Server: ' + $SMTP_Server + '<br />'
        $mail_body += 'To: ' + ($To_Address -join ';') + '<br />'
        if ($Cc_Address) {
            $mail_body += 'Cc: ' + ($Cc_Address -join ';') + '<br />'
        }
        if ($bcc_Address) {
            $mail_body += 'Cc: ' + ($bcc_Address -join ';') + '<br />'
        }
    }

    $mail_body += '<br /><b>[REPORT]</b><br />'
    $mail_body += 'Host: ' + ($env:computername) + '<br />'
    $mail_body += 'Config File: ' + (Resolve-Path $configFile).Path + '<br />'
    $mail_body += 'Report File: ' + (Resolve-Path $Report_File_Path).Path + '<br />'

    if ($config.Output.Enable_Transcript_Logging) {
        $mail_body += 'Transcript File: ' + (Resolve-Path $Transcript_File_Path).Path + '<br />'
    }

    $mail_body += '<br /><b>[EXCLUSIONS]</b><br />'
    $mail_body += 'Excluded Servers: ' + (@($config.Exclusion.Ignore_Server_Name) -join ';') + '<br />'
    $mail_body += 'Excluded Components: ' + (@($config.Exclusion.Ignore_Server_Component) -join ';') + '<br />'
    $mail_body += 'Excluded Mailbox Database: ' + (@($config.Exclusion.Ignore_MB_Database) -join ';') + '<br />'
    $mail_body += 'Excluded Public Database: ' + (@($config.Exclusion.Ignore_PF_Database) -join ';') + '<br />'
    $mail_body += '</p><p>'
    $mail_body += '<a href="' + $module_Info.ProjectURI.OriginalString + '" target="_blank">' + $module_Info.Name.ToString() + ' ' + $module_Info.Version.ToString() + '</a></p>'
    $mail_body += '</html>'
    $mail_body | Out-File $Report_File_Path
    'HTML Report @ ' + $Report_File_Path | Say
    # ----------------------------------------------------------------------------
    # Mail Parameters------------------------------------------------------------
    $params = @{
        Body       = $mail_body
        BodyAsHtml = $true
        Subject    = "[$($Company_Name)] $($Email_Subject) $($today)"
        From       = $Sender_Address
        SmtpServer = $SMTP_Server
        UseSsl     = $SSL_Enabled
        Port       = $Port
    }

    if ($To_Address) { $params.Add('To', $To_Address) }
    if ($Cc_Address) { $params.Add('Cc', $Cc_Address) }
    if ($Bcc_Address) { $params.Add('Bcc', $Bcc_Address) }

    # ----------------------------------------------------------------------------
    # Send Report----------------------------------------------------------------
    if ($Send_Email_Report -eq $true) { 'Sending Report...' | Say ; Send-MailMessage @params }
    # ----------------------------------------------------------------------------
    # "" | Say
    $hr | Say
    "Enabled Tests: $($enabledTestCount) [of $($availableTestCount)]" | Say
    "Failed: $($testFailed)" | Say
    "Passed: $($testPassed) [$($percentPassed)%]" | Say
    switch ($overAllResult) {
        'PASSED' { "Overall Result: $overAllResult" | Say -Color Green }
        'FAILED' { "Overall Result: $overAllResult" | Say -Color Yellow }
    }

    # "Overall Result: $($overAllResult)" | Say

    "======================================" | Say
    "" | Say

    # SCRIPT END------------------------------------------------------------------
    LogEnd
}
