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