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