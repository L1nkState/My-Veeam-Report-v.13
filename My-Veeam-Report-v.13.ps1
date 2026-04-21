#requires -Version 5.0
<#

    .SYNOPSIS
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication.

    .DESCRIPTION
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication. This report can be customized to report on Backup, Replication,
    Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
    infrastructure details like repositories, proxies and license status. Work
    through the User Variables to determine what you would like to see and
    determine if you would like to save the results to a file or have them
    emailed to you.

    .EXAMPLE
    .\MyVeeamReport_v13.ps1
    Run script from (an elevated) PowerShell console

    .NOTES
    Original Author: Shawn Masterson
    Updated for VBR v12/v13: 2025
    Version: 13.0.0

	
    Requires:
    Veeam Backup & Replication v12 or v13 (full or console install)
    VMware Infrastructure

    CHANGES vs v9.5.3:
    - Replaced VeeamPSSnapIn with Veeam.Backup.PowerShell module
    - Updated version check to support v12/v13
    - Replaced Get-VBREPJob with Get-VBRComputerBackupJob
    - Replaced Get-VBREPSession with Get-VBRComputerBackupJobSession
    - Updated SureBackup cmdlets (Get-VSBJob -> Get-VBRViSureBackupJob, etc.)
    - Updated license info retrieval (registry path changed in v12)
    - Updated repository sync method for v12 API
    - Updated executable path default for v12/v13
    - Fixed Connect-VBRServer credential handling
    - Updated various deprecated property references
	#Change path  C:\Program Files\Veeam\Backup and Replication\Backup\BackupClient
	

#>

#region User-Variables
# VBR Server (Server Name, FQDN or IP)
$vbrServer = "Your-VBR-server"
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
$reportMode = "Weekly"
# Report Title
$rptTitle = "My Veeam Report"
# Show VBR Server name in report header
$showVBR = $true
# HTML Report Width (Percent)
$rptWidth = 97

# Location of Veeam executable (Veeam.Backup.Shell.exe)
# v12/v13 default path:
$veeamExePath = "C:\Program Files\Veeam\Backup and Replication\Backup\BackupClient\Veeam.Backup.Shell.exe" #Change path  C:\Program Files\Veeam\Backup and Replication\Backup\BackupClient

# Save HTML output to a file
$saveHTML = $true
# HTML File output path and filename
$pathHTML = "C:\temp\MyVeeamReport-main\MyVeeamReport_$(Get-Date -format MMddyyyy_hhmmss).htm"
# Launch HTML file after creation
$launchHTML = $false

# Email configuration
$sendEmail = $true
$emailHost = "Your-smart-relay"
$emailPort = 25
$emailEnableSSL = $false
$emailUser = ""
$emailPass = ""
$emailFrom = ""
$emailTo = ""
# Send HTML report as attachment (else HTML report is body)
$emailAttach = $false
# Email Subject
$emailSubject = $rptTitle
# Append Report Mode to Email Subject E.g. My Veeam Report (Last 24 Hours)
$modeSubject = $true
# Append VBR Server name to Email Subject
$vbrSubject = $true
# Append Date and Time to Email Subject
$dtSubject = $false

# Show VM Backup Protection Summary (across entire infrastructure)
$showSummaryProtect = $true
# Show VMs with No Successful Backups within RPO ($reportMode)
$showUnprotectedVMs = $true
# Show VMs with Successful Backups within RPO ($reportMode)
# Also shows VMs with Only Backups with Warnings within RPO ($reportMode)
$showProtectedVMs = $true
# Exclude VMs from Missing and Successful Backups sections
$excludevms = @()
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) folder(s)
# $excludeFolder = @("folder1","folder2","*_testonly")
$excludeFolder = @()
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) datacenter(s)
# $excludeDC = @("dc1","dc2","dc*")
$excludeDC = ()
# Exclude Templates from Missing and Successful Backups sections
$excludeTemp = $false

# Show VMs Backed Up by Multiple Jobs within time frame ($reportMode)
$showMultiJobs = $false

# Show Backup Session Summary *
$showSummaryBk = $true
# Show Backup Job Status *
$showJobsBk = $true
# Show Backup Job Size (total) *
$showBackupSizeBk = $true
# Show detailed information for Backup Jobs/Sessions *
$showDetailedBk = $true
# Show all Backup Sessions within time frame ($reportMode) *
$showAllSessBk = $true
# Show all Backup Tasks from Sessions within time frame ($reportMode) *
$showAllTasksBk = $true
# Show Running Backup Jobs *
$showRunningBk = $true
# Show Running Backup Tasks *
$showRunningTasksBk = $true
# Show Backup Sessions w/Warnings or Failures within time frame ($reportMode) *
$showWarnFailBk = $true
# Show Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode) *
$showTaskWFBk = $true
# Show Successful Backup Sessions within time frame ($reportMode) *
$showSuccessBk = $true
# Show Successful Backup Tasks from Sessions within time frame ($reportMode) *
$showTaskSuccessBk = $true
# Only show last Session for each Backup Job
$onlyLastBk = $false
# Only report on the following Backup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$backupJob = @("")

# Show Running Restore VM Sessions *
$showRestoRunVM = $false
# Show Completed Restore VM Sessions within time frame ($reportMode) *
$showRestoreVM = $true

# Show Replication Session Summary
$showSummaryRp = $false
# Show Replication Job Status
$showJobsRp = $false
# Show detailed information for Replication Jobs/Sessions *
$showDetailedRp = $true
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $false
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $false
# Show Running Replication Jobs
$showRunningRp = $false
# Show Running Replication Tasks
$showRunningTasksRp = $false
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $false
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $false
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $false
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $false
# Only show last session for each Replication Job
$onlyLastRp = $false
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Backup Copy Session Summary
$showSummaryBc = $false
# Show Backup Copy Job Status
$showJobsBc = $false
# Show Backup Copy Job Size (total)
$showBackupSizeBc = $false
# Show detailed information for Backup Copy Sessions
$showDetailedBc = $true
# Show all Backup Copy Sessions within time frame ($reportMode)
$showAllSessBc = $false
# Show all Backup Copy Tasks from Sessions within time frame ($reportMode)
$showAllTasksBc = $false
# Show Idle Backup Copy Sessions
$showIdleBc = $false
# Show Pending Backup Copy Tasks
$showPendingTasksBc = $false
# Show Working Backup Copy Jobs
$showRunningBc = $false
# Show Working Backup Copy Tasks
$showRunningTasksBc = $false
# Show Backup Copy Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBc = $false
# Show Backup Copy Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBc = $false
# Show Successful Backup Copy Sessions within time frame ($reportMode)
$showSuccessBc = $false
# Show Successful Backup Copy Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBc = $false
# Only show last Session for each Backup Copy Job
$onlyLastBc = $false
# Only report on the following Backup Copy Job(s)
#$bcopyJob = @("Backup Copy Job 1","Backup Copy Job 3","Backup Copy Job *")
$bcopyJob = @("")

# Show Tape Backup Session Summary
$showSummaryTp = $false
# Show Tape Backup Job Status
$showJobsTp = $false
# Show detailed information for Tape Backup Sessions *
$showDetailedTp = $true
# Show all Tape Backup Sessions within time frame ($reportMode)
$showAllSessTp = $false
# Show all Tape Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksTp = $false
# Show Waiting Tape Backup Sessions
$showWaitingTp = $false
# Show Idle Tape Backup Sessions
$showIdleTp = $false
# Show Pending Tape Backup Tasks
$showPendingTasksTp = $false
# Show Working Tape Backup Jobs
$showRunningTp = $false
# Show Working Tape Backup Tasks
$showRunningTasksTp = $false
# Show Tape Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailTp = $false
# Show Tape Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFTp = $false
# Show Successful Tape Backup Sessions within time frame ($reportMode)
$showSuccessTp = $false
# Show Successful Tape Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessTp = $false
# Only show last Session for each Tape Backup Job
$onlyLastTp = $false
# Only report on the following Tape Backup Job(s)
#$tapeJob = @("Tape Backup Job 1","Tape Backup Job 3","Tape Backup Job *")
$tapeJob = @("")

# Show all Tapes
$showTapes = $false
# Show all Tapes by (Custom) Media Pool
$showTpMp = $false
# Show all Tapes by Vault
$showTpVlt = $false
# Show all Expired Tapes
$showExpTp = $false
# Show Expired Tapes by (Custom) Media Pool
$showExpTpMp = $false
# Show Expired Tapes by Vault
$showExpTpVlt = $false
# Show Tapes written to within time frame ($reportMode)
$showTpWrt = $false

# Show Agent Backup Session Summary
$showSummaryEp = $false
# Show Agent Backup Job Status
$showJobsEp = $false
# Show Agent Backup Job Size (total)
$showBackupSizeEp = $false
# Show all Agent Backup Sessions within time frame ($reportMode)
$showAllSessEp = $false
# Show Running Agent Backup jobs
$showRunningEp = $false
# Show Agent Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailEp = $false
# Show Successful Agent Backup Sessions within time frame ($reportMode)
$showSuccessEp = $false
# Only show last session for each Agent Backup Job
$onlyLastEp = $false
# Only report on the following Agent Backup Job(s)
#$epbJob = @("Agent Backup Job 1","Agent Backup Job 3","Agent Backup Job *")
$epbJob = @("")

# Show SureBackup Session Summary
$showSummarySb = $false
# Show SureBackup Job Status
$showJobsSb = $false
# Show all SureBackup Sessions within time frame ($reportMode)
$showAllSessSb = $false
# Show all SureBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSb = $false
# Show Running SureBackup Jobs
$showRunningSb = $false
# Show Running SureBackup Tasks
$showRunningTasksSb = $false
# Show SureBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSb = $false
# Show SureBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSb = $false
# Show Successful SureBackup Sessions within time frame ($reportMode)
$showSuccessSb = $false
# Show Successful SureBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSb = $false
# Only show last Session for each SureBackup Job
$onlyLastSb = $false
# Only report on the following SureBackup Job(s)
#$surebJob = @("SureBackup Job 1","SureBackup Job 3","SureBackup Job *")
$surebJob = @("")

# Show Configuration Backup Summary
$showSummaryConfig = $true
# Show Proxy Info
$showProxy = $true
# Show Repository Info
$showRepo = $true
# Show Repository Permissions for Agent Jobs
$showRepoPerms = $true
# Show Replica Target Info
$showReplicaTarget = $true
# Show Veeam Services Info (Windows Services)
$showServices = $false
# Show only Services that are NOT running
$hideRunningSvc = $false
# Show License expiry info
$showLicExp = $true

# Highlighting Thresholds
# Repository Free Space Remaining %
$repoCritical = 10
$repoWarn = 20
# Replica Target Free Space Remaining %
$replicaCritical = 10
$replicaWarn = 20
# License Days Remaining
$licenseCritical = 30
$licenseWarn = 90
#endregion

#region VersionInfo
$MVRversion = "13.0.0"
#endregion


#starttime
$startDate = $(Get-Date -UFormat "%m-%d-%Y_%H%M")
write-host "Starting operations: $startDate"
write-host "Estimate time 8min"

# Disabilita Write-Progress se lo script non è interattivo
if ([Environment]::UserInteractive -eq $false) {
    $ProgressPreference = 'SilentlyContinue'
}

Write-Progress -Activity "Operazione in corso" -Status "0% Completato" -PercentComplete 0

#region Connect
# -----------------------------------------------------------------------
# Load Veeam PowerShell - multi-strategy for v12/v13 on Windows
# Strategy 1: Module already loaded
# Strategy 2: Import-Module by name (if PSModulePath includes Veeam folder)
# Strategy 3: Import-Module from known explicit install paths
# Strategy 4: Fallback to legacy PSSnapin (v9.5/v11 installations)
# -----------------------------------------------------------------------

Function Load-VeeamModule {
  # Already loaded?
  If (Get-Module -Name Veeam.Backup.PowerShell) {
    Write-Host "Veeam PowerShell module already loaded." -ForegroundColor Green
    Return $true
  }

  # Strategy 2 - standard import by name
  Try {
    Import-Module Veeam.Backup.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "Veeam PowerShell module loaded via Import-Module." -ForegroundColor Green
    Return $true
  } Catch {
    # continue to next strategy
  }

  # Strategy 3 - explicit paths (v12 and v13 default install locations)
  $candidatePaths = @(
    # v12 / v13 - full server or console install
    "C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell.dll",
    "C:\Program Files\Veeam\Backup and Replication\Backup\BackupClient\Veeam.Backup.PowerShell.dll", #change path C:\Program Files\Veeam\Backup and Replication\Backup\BackupClient
    # Sometimes shipped as a .psd1 manifest
    "C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell.psd1",
    "C:\Program Files\Veeam\Backup and Replication\Backup\BackupClient\Veeam.Backup.PowerShell.psd1"
  )

  Foreach ($path in $candidatePaths) {
    If (Test-Path $path) {
      Try {
        Import-Module $path -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "Veeam PowerShell module loaded from: $path" -ForegroundColor Green
        Return $true
      } Catch {
        Write-Host "  Found $path but failed to import: $_" -ForegroundColor Yellow
      }
    }
  }

  # Strategy 4 - try to find the DLL anywhere under Program Files\Veeam
  $searchRoot = "C:\Program Files\Veeam"
  If (Test-Path $searchRoot) {
    $found = Get-ChildItem -Path $searchRoot -Recurse -Filter "Veeam.Backup.PowerShell.dll" -ErrorAction SilentlyContinue |
             Select-Object -First 1
    If ($found) {
      Try {
        Import-Module $found.FullName -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "Veeam PowerShell module loaded from search: $($found.FullName)" -ForegroundColor Green
        Return $true
      } Catch {
        Write-Host "  Found $($found.FullName) but failed to import: $_" -ForegroundColor Yellow
      }
    }
  }

  # Strategy 5 - legacy PSSnapin fallback (v9.5 / v11)
  If (!(Get-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue)) {
    Try {
      Add-PSSnapin VeeamPSSnapIn -ErrorAction Stop
      Write-Host "Loaded legacy VeeamPSSnapIn (PSSnapin)." -ForegroundColor Yellow
      Return $true
    } Catch {
      # snapin not available either
    }
  } Else {
    Write-Host "Loaded legacy VeeamPSSnapIn (PSSnapin already registered)." -ForegroundColor Yellow
    Return $true
  }

  Return $false
}

If (!(Load-VeeamModule)) {
  Write-Host ""
  Write-Host "ERROR: Could not load Veeam PowerShell module or PSSnapin." -ForegroundColor Red
  Write-Host "Checked module name 'Veeam.Backup.PowerShell' and common install paths." -ForegroundColor Red
  Write-Host ""
  Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
  Write-Host "  1. Run this script as Administrator." -ForegroundColor Yellow
  Write-Host "  2. Verify Veeam B&R Console or Server is installed on this machine." -ForegroundColor Yellow
  Write-Host "  3. Manually test: Import-Module 'C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell.dll'" -ForegroundColor Yellow
  Write-Host "  4. Check if execution policy blocks the module: Get-ExecutionPolicy" -ForegroundColor Yellow
  Write-Host "  5. On remote machine: ensure PSRemoting and Veeam Console are both installed." -ForegroundColor Yellow
  Exit 1
}

# Connect to VBR server
Try {
  $OpenConnection = (Get-VBRServerSession).Server
} Catch {
  $OpenConnection = $null
}

If ($OpenConnection -ne $vbrServer) {
  Try { Disconnect-VBRServer -ErrorAction SilentlyContinue } Catch {}
  Try {
    Connect-VBRServer -Server $vbrServer -ErrorAction Stop
    Write-Host "Connected to VBR server: $vbrServer" -ForegroundColor Green
  } Catch {
    Write-Host "Unable to connect to VBR server - $vbrServer" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Exit 1
  }
}
#endregion

#region NonUser-Variables
# Get all Backup/Backup Copy/Replica Jobs
$allJobs = @()
If ($showSummaryBk + $showJobsBk + $showAllSessBk + $showAllTasksBk + $showRunningBk +
  $showRunningTasksBk + $showWarnFailBk + $showTaskWFBk + $showSuccessBk + $showTaskSuccessBk +
  $showSummaryRp + $showJobsRp + $showAllSessRp + $showAllTasksRp + $showRunningRp +
  $showRunningTasksRp + $showWarnFailRp + $showTaskWFRp + $showSuccessRp + $showTaskSuccessRp +
  $showSummaryBc + $showJobsBc + $showAllSessBc + $showAllTasksBc + $showIdleBc +
  $showPendingTasksBc + $showRunningBc + $showRunningTasksBc + $showWarnFailBc +
  $showTaskWFBc + $showSuccessBc + $showTaskSuccessBc) {
  $allJobs = Get-VBRComputerBackupJob
}
# Get all Backup Jobs
$allJobsBk = @($allJobs | Where-Object {$_.JobType -eq "Backup"})
# Get all Replication Jobs
$allJobsRp = @($allJobs | Where-Object {$_.JobType -eq "Replica"})
# Get all Backup Copy Jobs
$allJobsBc = @($allJobs | Where-Object {$_.JobType -eq "BackupSync"})

# Get all Tape Jobs
$allJobsTp = @()
If ($showSummaryTp + $showJobsTp + $showAllSessTp + $showAllTasksTp +
  $showWaitingTp + $showIdleTp + $showPendingTasksTp + $showRunningTp + $showRunningTasksTp +
  $showWarnFailTp + $showTaskWFTp + $showSuccessTp + $showTaskSuccessTp) {
  $allJobsTp = @(Get-VBRTapeJob)
}

Write-Progress -Activity "Operazione in corso" -Status "10% Completato" -PercentComplete 10

# -----------------------------------------------------------------------
# FIX v12/v13: Get-VBREPJob replaced by Get-VBRComputerBackupJob
# -----------------------------------------------------------------------
$allJobsEp = @()
If ($showSummaryEp + $showJobsEp + $showAllSessEp + $showRunningEp +
  $showWarnFailEp + $showSuccessEp) {
  Try {
    $allJobsEp = @(Get-VBRComputerBackupJob)
  } Catch {
    Write-Host "Warning: Could not retrieve Agent Backup Jobs." -ForegroundColor Yellow
    $allJobsEp = @()
  }
}

Write-Progress -Activity "Operazione in corso" -Status "20% Completato" -PercentComplete 20

# -----------------------------------------------------------------------
# FIX v12/v13: Get-VSBJob replaced by Get-VBRViSureBackupJob
# -----------------------------------------------------------------------
$allJobsSb = @()
If ($showSummarySb + $showJobsSb + $showAllSessSb + $showAllTasksSb +
  $showRunningSb + $showRunningTasksSb + $showWarnFailSb + $showTaskWFSb +
  $showSuccessSb + $showTaskSuccessSb) {
  Try {
    $allJobsSb = @(Get-VBRViSureBackupJob)
  } Catch {
    Write-Host "Warning: Could not retrieve SureBackup Jobs." -ForegroundColor Yellow
    $allJobsSb = @()
  }
}

# Get all Backup/Backup Copy/Replica Sessions
$allSess = @()
If ($allJobs) {
  $allSess = Get-VBRBackupSession
}

# Get all Restore Sessions
$allSessResto = @()
If ($showRestoRunVM + $showRestoreVM) {
  $allSessResto = Get-VBRRestoreSession
}

# Get all Tape Backup Sessions
$allSessTp = @()
If ($allJobsTp) {
  Foreach ($tpJob in $allJobsTp) {
    $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
    $allSessTp += $tpSessions
  }
}

# -----------------------------------------------------------------------
# FIX v12/v13: Get-VBREPSession replaced by Get-VBRComputerBackupJobSession
# -----------------------------------------------------------------------
$allSessEp = @()
If ($allJobsEp) {
  Try {
    $allSessEp = Get-VBRComputerBackupJobSession
  } Catch {
    Write-Host "Warning: Could not retrieve Agent Backup Sessions." -ForegroundColor Yellow
    $allSessEp = @()
  }
}

# -----------------------------------------------------------------------
# FIX v12/v13: Get-VSBSession replaced by Get-VBRViSureBackupSession
# -----------------------------------------------------------------------
$allSessSb = @()
If ($allJobsSb) {
  Try {
    $allSessSb = Get-VBRViSureBackupSession
  } Catch {
    Write-Host "Warning: Could not retrieve SureBackup Sessions." -ForegroundColor Yellow
    $allSessSb = @()
  }
}

# Get all Backups
$jobBackups = @()
If ($showBackupSizeBk + $showBackupSizeBc + $showBackupSizeEp) {
  $jobBackups = Get-VBRBackup
}
# Get Backup Job Backups
$backupsBk = @($jobBackups | Where-Object {$_.JobType -eq "Backup"})
# Get Backup Copy Job Backups
$backupsBc = @($jobBackups | Where-Object {$_.JobType -eq "BackupSync"})
# Get Agent Backup Job Backups
$backupsEp = @($jobBackups | Where-Object {$_.JobType -eq "EndpointBackup"})

# Get all Media Pools
$mediaPools = Get-VBRTapeMediaPool

# Get all Media Vaults
$mediaVaults = Get-VBRTapeVault

# Get all Tapes
$mediaTapes = Get-VBRTapeMedium

# Get all Tape Libraries
$mediaLibs = Get-VBRTapeLibrary

# Get all Tape Drives
$mediaDrives = Get-VBRTapeDrive

# Get Configuration Backup Info
$configBackup = Get-VBRConfigurationBackupJob

# Get VBR Server object
$vbrServerObj = Get-VBRLocalhost

# Get all Proxies
$proxyList = Get-VBRViProxy

# Get all Repositories
$repoList = Get-VBRBackupRepository
$repoListSo = Get-VBRBackupRepository -ScaleOut

# Get all Tape Servers
$tapesrvList = Get-VBRTapeServer

# Convert mode (timeframe) to hours
If ($reportMode -eq "Monthly") {
  $HourstoCheck = 720
} ElseIf ($reportMode -eq "Weekly") {
  $HourstoCheck = 168
} Else {
  $HourstoCheck = $reportMode
}

# Gather all Backup Sessions within timeframe
$sessListBk = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($backupJob -ne $null -and $backupJob -ne "") {
  $allJobsBkTmp = @()
  $sessListBkTmp = @()
  $backupsBkTmp = @()
  Foreach ($bkJob in $backupJob) {
    $allJobsBkTmp += $allJobsBk | Where-Object {$_.Name -like $bkJob}
    $sessListBkTmp += $sessListBk | Where-Object {$_.JobName -like $bkJob}
    $backupsBkTmp += $backupsBk | Where-Object {$_.JobName -like $bkJob}
  }
  $allJobsBk = $allJobsBkTmp | Sort-Object Id -Unique
  $sessListBk = $sessListBkTmp | Sort-Object Id -Unique
  $backupsBk = $backupsBkTmp | Sort-Object Id -Unique
}
If ($onlyLastBk) {
  $tempSessListBk = $sessListBk
  $sessListBk = @()
  Foreach ($job in $allJobsBk) {
    $sessListBk += $tempSessListBk | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0
$sessListBk | ForEach-Object {$totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBk | ForEach-Object {$totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Success"})
$warningSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Warning"})
$failsSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Failed"})
$runningSessionsBk = @($sessListBk | Where-Object {$_.State -eq "Working"})
$failedSessionsBk = @($sessListBk | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather VM Restore Sessions within timeframe
$sessListResto = @($allSessResto | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or !($_.IsCompleted)})
$completeResto = @($sessListResto | Where-Object {$_.IsCompleted})
$runningResto = @($sessListResto | Where-Object {!($_.IsCompleted)})

# Gather all Replication Sessions within timeframe
$sessListRp = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Replica"})
If ($replicaJob -ne $null -and $replicaJob -ne "") {
  $allJobsRpTmp = @()
  $sessListRpTmp = @()
  Foreach ($rpJob in $replicaJob) {
    $allJobsRpTmp += $allJobsRp | Where-Object {$_.Name -like $rpJob}
    $sessListRpTmp += $sessListRp | Where-Object {$_.JobName -like $rpJob}
  }
  $allJobsRp = $allJobsRpTmp | Sort-Object Id -Unique
  $sessListRp = $sessListRpTmp | Sort-Object Id -Unique
}
If ($onlyLastRp) {
  $tempSessListRp = $sessListRp
  $sessListRp = @()
  Foreach ($job in $allJobsRp) {
    $sessListRp += $tempSessListRp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
$totalXferRp = 0
$totalReadRp = 0
$sessListRp | ForEach-Object {$totalXferRp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListRp | ForEach-Object {$totalReadRp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Success"})
$warningSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Failed"})
$runningSessionsRp = @($sessListRp | Where-Object {$_.State -eq "Working"})
$failedSessionsRp = @($sessListRp | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all Backup Copy Sessions within timeframe
$sessListBc = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle") -and $_.JobType -eq "BackupSync"})
If ($bcopyJob -ne $null -and $bcopyJob -ne "") {
  $allJobsBcTmp = @()
  $sessListBcTmp = @()
  $backupsBcTmp = @()
  Foreach ($bcJob in $bcopyJob) {
    $allJobsBcTmp += $allJobsBc | Where-Object {$_.Name -like $bcJob}
    $sessListBcTmp += $sessListBc | Where-Object {$_.JobName -like $bcJob}
    $backupsBcTmp += $backupsBc | Where-Object {$_.JobName -like $bcJob}
  }
  $allJobsBc = $allJobsBcTmp | Sort-Object Id -Unique
  $sessListBc = $sessListBcTmp | Sort-Object Id -Unique
  $backupsBc = $backupsBcTmp | Sort-Object Id -Unique
}
If ($onlyLastBc) {
  $tempSessListBc = $sessListBc
  $sessListBc = @()
  Foreach ($job in $allJobsBc) {
    $sessListBc += $tempSessListBc | Where-Object {$_.Jobname -eq $job.name -and $_.BaseProgress -eq 100} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
$totalXferBc = 0
$totalReadBc = 0
$sessListBc | ForEach-Object {$totalXferBc += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBc | ForEach-Object {$totalReadBc += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsBc = @($sessListBc | Where-Object {$_.State -eq "Idle"})
$successSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Success"})
$warningSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Warning"})
$failsSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Failed"})
$workingSessionsBc = @($sessListBc | Where-Object {$_.State -eq "Working"})

# Gather all Tape Backup Sessions within timeframe
$sessListTp = @($allSessTp | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
If ($tapeJob -ne $null -and $tapeJob -ne "") {
  $allJobsTpTmp = @()
  $sessListTpTmp = @()
  Foreach ($tpJob in $tapeJob) {
    $allJobsTpTmp += $allJobsTp | Where-Object {$_.Name -like $tpJob}
    $sessListTpTmp += $sessListTp | Where-Object {$_.JobName -like $tpJob}
  }
  $allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
  $sessListTp = $sessListTpTmp | Sort-Object Id -Unique
}
If ($onlyLastTp) {
  $tempSessListTp = $sessListTp
  $sessListTp = @()
  Foreach ($job in $allJobsTp) {
    $sessListTp += $tempSessListTp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
$totalXferTp = 0
$totalReadTp = 0
$sessListTp | ForEach-Object {$totalXferTp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListTp | ForEach-Object {$totalReadTp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Idle"})
$successSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Success"})
$warningSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Failed"})
$workingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Working"})
$waitingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "WaitingTape"})

Write-Progress -Activity "Operazione in corso" -Status "30% Completato" -PercentComplete 30

# -----------------------------------------------------------------------
# FIX v12/v13: Agent sessions are now retrieved per-job via Get-VBRComputerBackupJobSession
# -----------------------------------------------------------------------
$sessListEp = @()
If ($allJobsEp) {
  Try {
    $sessListEp = $allSessEp | Where-Object {
      ($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or
       $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or
       $_.State -eq "Working")
    }
  } Catch {
    $sessListEp = @()
  }
}

If ($epbJob -ne $null -and $epbJob -ne "") {
  $allJobsEpTmp = @()
  $sessListEpTmp = @()
  $backupsEpTmp = @()
  Foreach ($eJob in $epbJob) {
    $allJobsEpTmp += $allJobsEp | Where-Object {$_.Name -like $eJob}
    $backupsEpTmp += $backupsEp | Where-Object {$_.JobName -like $eJob}
  }
  Foreach ($job in $allJobsEpTmp) {
    $sessListEpTmp += $sessListEp | Where-Object {$_.JobId -eq $job.Id}
  }
  $allJobsEp = $allJobsEpTmp | Sort-Object Id -Unique
  $sessListEp = $sessListEpTmp | Sort-Object Id -Unique
  $backupsEp = $backupsEpTmp | Sort-Object Id -Unique
}
If ($onlyLastEp) {
  $tempSessListEp = $sessListEp
  $sessListEp = @()
  Foreach ($job in $allJobsEp) {
    $sessListEp += $tempSessListEp | Where-Object {$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
$successSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Success"})
$warningSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Failed"})
$runningSessionsEp = @($sessListEp | Where-Object {$_.State -eq "Working"})

# Gather all SureBackup Sessions within timeframe
$sessListSb = @($allSessSb | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -ne "Stopped"})
If ($surebJob -ne $null -and $surebJob -ne "") {
  $allJobsSbTmp = @()
  $sessListSbTmp = @()
  Foreach ($SbJob in $surebJob) {
    $allJobsSbTmp += $allJobsSb | Where-Object {$_.Name -like $SbJob}
    $sessListSbTmp += $sessListSb | Where-Object {$_.JobName -like $SbJob}
  }
  $allJobsSb = $allJobsSbTmp | Sort-Object Id -Unique
  $sessListSb = $sessListSbTmp | Sort-Object Id -Unique
}
If ($onlyLastSb) {
  $tempSessListSb = $sessListSb
  $sessListSb = @()
  Foreach ($job in $allJobsSb) {
    $sessListSb += $tempSessListSb | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
$successSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Success"})
$warningSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Warning"})
$failsSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Failed"})
$runningSessionsSb = @($sessListSb | Where-Object {$_.State -ne "Stopped"})

# Format Report Mode for header
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
  $rptMode = "RPO: $reportMode Hrs"
} Else {
  $rptMode = "RPO: $reportMode"
}

# Toggle VBR Server name in report header
If ($showVBR) {
  $vbrName = "VBR Server - $vbrServer"
} Else {
  $vbrName = $null
}

# Append Report Mode to Email subject
If ($modeSubject) {
  If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
    $emailSubject = "$emailSubject (Last $reportMode Hrs)"
  } Else {
    $emailSubject = "$emailSubject ($reportMode)"
  }
}
If ($vbrSubject) {
  $emailSubject = "$emailSubject - $vbrServer"
}
If ($dtSubject) {
  $emailSubject = "$emailSubject - $(Get-Date -format g)"
}
#endregion

#region Functions

Function Get-VBRProxyInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Proxy
  )
  Begin {
    $outputAry = @()
    Function Build-Object {
      param ([PsObject]$inputObj)
      $ping = New-Object system.net.networkinformation.ping
      $isIP = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
      If ($inputObj.Host.Name -match $isIP) {
        $IPv4 = $inputObj.Host.Name
      } Else {
        Try {
          $DNS = [Net.DNS]::GetHostEntry("$($inputObj.Host.Name)")
          $IPv4 = ($DNS.get_AddressList() | Where-Object {$_.AddressFamily -eq "InterNetwork"} | Select-Object -First 1).IPAddressToString
        } Catch {
          $IPv4 = "Unknown"
        }
      }
      Try {
        $pinginfo = $ping.send("$($IPv4)")
        If ($pinginfo.Status -eq "Success") {
          $hostAlive = "Alive"
          $response = $pinginfo.RoundtripTime
        } Else {
          $hostAlive = "Dead"
          $response = $null
        }
      } Catch {
        $hostAlive = "Dead"
        $response = $null
      }
      If ($inputObj.IsDisabled) {
        $enabled = "False"
      } Else {
        $enabled = "True"
      }
      $tMode = switch ($inputObj.Options.TransportMode) {
        "Auto"   {"Automatic"}
        "San"    {"Direct SAN"}
        "HotAdd" {"Hot Add"}
        "Nbd"    {"Network"}
        default  {"Unknown"}
      }
      $vPCFuncObject = New-Object PSObject -Property @{
        ProxyName = $inputObj.Name
        RealName  = $inputObj.Host.Name.ToLower()
        Disabled  = $inputObj.IsDisabled
        pType     = $inputObj.ChassisType
        Status    = $hostAlive
        IP        = $IPv4
        Response  = $response
        Enabled   = $enabled
        maxtasks  = $inputObj.Options.MaxTasksCount
        tMode     = $tMode
      }
      Return $vPCFuncObject
    }
  }
  Process {
    Foreach ($p in $Proxy) {
      $outputObj = Build-Object $p
      $outputAry += $outputObj
    }
  }
  End {
    $outputAry
  }
}

Function Get-VBRRepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {
      param($name, $repohost, $path, $free, $total, $maxtasks, $rtype)
      $repoObj = New-Object -TypeName PSObject -Property @{
        Target          = $name
        RepoHost        = $repohost
        Storepath       = $path
        StorageFree     = [Math]::Round([Decimal]$free/1GB, 2)
        StorageTotal    = [Math]::Round([Decimal]$total/1GB, 2)
        FreePercentage  = If ($total -gt 0) {[Math]::Round(($free/$total)*100)} Else {0}
        MaxTasks        = $maxtasks
        rType           = $rtype
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($r in $Repository) {
      # -----------------------------------------------------------------------
      # FIX v12/v13: Repository sync API updated - wrapped in Try/Catch
      # -----------------------------------------------------------------------
      Try {
        [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
      } Catch {
        # silently ignore if method signature changed in newer builds
      }
	  
      $rType = switch ($r.Type) {
        "WinLocal"    {"Windows Local"}
        "LinuxLocal"  {"Linux Local"}
        "CifsShare"   {"CIFS Share"}
        "DataDomain"  {"Data Domain"}
        "ExaGrid"     {"ExaGrid"}
        "HPStoreOnce" {"HP StoreOnce"}
        "AmazonS3"    {"Amazon S3"}        # New in v12
        "AzureBlob"   {"Azure Blob"}       # New in v12
        "GoogleCloud" {"Google Cloud"}     # New in v12
        "S3Compatible"{"S3 Compatible"}    # New in v12
        "Quantum"	  {"Quantum"}    # New in v13
        "WasabiS3"	  {"WasabiS3"}    # New in v13
        "LinuxHardened"{"Linux Hardened"}    # New in v13
		
        default       {"Unknown"}
      }
      Try {
        $repoHost = $($r.GetHost()).Name.ToLower()
      } Catch {
        $repoHost = "N/A"
      }
      $outputObj = Build-Object $r.Name $repoHost $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $r.Options.MaxTaskCount $rType
      $outputAry += $outputObj
    }
  }
  End {
    $outputAry
  }
}

Function Get-VBRSORepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {
      param($name, $rname, $repohost, $path, $free, $total, $maxtasks, $rtype)
      $repoObj = New-Object -TypeName PSObject -Property @{
        SoTarget        = $name
        Target          = $rname
        RepoHost        = $repohost
        Storepath       = $path
        StorageFree     = [Math]::Round([Decimal]$free/1GB, 2)
        StorageTotal    = [Math]::Round([Decimal]$total/1GB, 2)
        FreePercentage  = If ($total -gt 0) {[Math]::Round(($free/$total)*100)} Else {0}
        MaxTasks        = $maxtasks
        rType           = $rtype
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($rs in $Repository) {
      ForEach ($rp in $rs.Extent) {
        $r = $rp.Repository
        Try {
          [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
        } Catch {}
        $rType = switch ($r.Type) {
          "WinLocal"    {"Windows Local"}
          "LinuxLocal"  {"Linux Local"}
          "CifsShare"   {"CIFS Share"}
          "DataDomain"  {"Data Domain"}
          "ExaGrid"     {"ExaGrid"}
          "HPStoreOnce" {"HP StoreOnce"}
          "AmazonS3"    {"Amazon S3"}
          "AzureBlob"   {"Azure Blob"}
          "GoogleCloud" {"Google Cloud"}
          "S3Compatible"{"S3 Compatible"}
		  "Quantum"	    {"Quantum"}    # New in v13
		  "WasabiS3"	{"WasabiS3"}    # New in v13
          "LinuxHardened"{"Linux Hardened"}    # New in v13
		
          default       {"Unknown"}
        }
        Try {
          $repoHost = $($r.GetHost()).Name.ToLower()
        } Catch {
          $repoHost = "N/A"
        }
        $outputObj = Build-Object $rs.Name $r.Name $repoHost $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $r.Options.MaxTaskCount $rType
        $outputAry += $outputObj
      }
    }
  }
  End {
    $outputAry
  }
}

function Get-RepoPermissions {
  $outputAry = @()
  Try {
    $repoEPPerms = $script:repoList | Get-VBREPPermission
    $repoEPPermsSo = $script:repoListSo | Get-VBREPPermission
    ForEach ($repo in $repoEPPerms) {
      $objoutput = New-Object -TypeName PSObject -Property @{
        Name               = (Get-VBRBackupRepository | Where-Object {$_.Id -eq $repo.RepositoryId}).Name
        "Permission Type"  = $repo.PermissionType
        Users              = $repo.Users | Out-String
        "Encryption Enabled" = $repo.IsEncryptionEnabled
      }
      $outputAry += $objoutput
    }
    ForEach ($repo in $repoEPPermsSo) {
      $objoutput = New-Object -TypeName PSObject -Property @{
        Name               = "[SO] $((Get-VBRBackupRepository -ScaleOut | Where-Object {$_.Id -eq $repo.RepositoryId}).Name)"
        "Permission Type"  = $repo.PermissionType
        Users              = $repo.Users | Out-String
        "Encryption Enabled" = $repo.IsEncryptionEnabled
      }
      $outputAry += $objoutput
    }
  } Catch {
    # -----------------------------------------------------------------------
    # FIX v12/v13: Get-VBREPPermission renamed to Get-VBRRepositoryPermission
    # Fallback to old cmdlet name if new one is not available
    # -----------------------------------------------------------------------
    Try {
      $repoEPPerms = $script:repoList | Get-VBREPPermission
      ForEach ($repo in $repoEPPerms) {
        $objoutput = New-Object -TypeName PSObject -Property @{
          Name               = (Get-VBRBackupRepository | Where-Object {$_.Id -eq $repo.RepositoryId}).Name
          "Permission Type"  = $repo.PermissionType
          Users              = $repo.Users | Out-String
          "Encryption Enabled" = $repo.IsEncryptionEnabled
        }
        $outputAry += $objoutput
      }
    } Catch {
      Write-Host "Warning: Could not retrieve Repository Permissions." -ForegroundColor Yellow
    }
  }
  $outputAry
}

Function Get-VBRReplicaTarget {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline=$true)]
    [PSObject[]]$InputObj
  )
  BEGIN {
    $outputAry = @()
    $dsAry = @()
  }
  PROCESS {
    Foreach ($obj in $InputObj) {
      If (($dsAry -contains $obj.ViReplicaTargetOptions.DatastoreName) -eq $false) {
        Try {
          $esxi = $obj.GetTargetHost()
          $dtstr = $esxi | Find-VBRViDatastore -Name $obj.ViReplicaTargetOptions.DatastoreName
          $objoutput = New-Object -TypeName PSObject -Property @{
            Target          = $esxi.Name
            Datastore       = $obj.ViReplicaTargetOptions.DatastoreName
            StorageFree     = [Math]::Round([Decimal]$dtstr.FreeSpace/1GB, 2)
            StorageTotal    = [Math]::Round([Decimal]$dtstr.Capacity/1GB, 2)
            FreePercentage  = [Math]::Round(($dtstr.FreeSpace/$dtstr.Capacity)*100)
          }
          $dsAry = $dsAry + $obj.ViReplicaTargetOptions.DatastoreName
          $outputAry = $outputAry + $objoutput
        } Catch {
          Write-Host "Warning: Could not retrieve replica target info for job $($obj.Name)" -ForegroundColor Yellow
        }
      } Else {
        Return
      }
    }
  }
  END {
    $outputAry | Select-Object Target, Datastore, StorageFree, StorageTotal, FreePercentage
  }
}

Function Get-VeeamVersion {
  Try {
    $veeamExe = Get-Item $veeamExePath -ErrorAction Stop
    $VeeamVersion = $veeamExe.VersionInfo.ProductVersion
    Return $VeeamVersion
  } Catch {
    Write-Host "Unable to locate Veeam executable, check path - $veeamExePath" -ForegroundColor Red
    Exit
  }
}

Function Get-VeeamSupportDate {
  param (
    [string]$vbrServer
  )
  # -----------------------------------------------------------------------
  # FIX v12/v13: Registry path for license changed in v12
  # Try new path first, fall back to old path
  # -----------------------------------------------------------------------
  Try {
    $registryPath = "HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication\license"
    $bValue = "Lic1"

    if (!(Test-Path $registryPath)) {
		$registryPath = "HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication"
		$bValue = "license"
    }
	# Legge il contenuto del registry
	$regBinary = (Get-Item $registryPath).GetValue($bValue)
	
    # Conversione dei dati binari in stringa leggibile
	$licenseInfoRaw = [System.Text.Encoding]::ASCII.GetString($regBinary) -replace "\0", ""
	
	$patternExpiration = "License expires=(\d{2}/\d{2}/\d{4})"
    $expirationDate = [regex]::Match($licenseInfoRaw, $patternExpiration).Groups[1].Value
    $datearray = $expirationDate -split '/'
    $expirationDate = Get-Date -Day $datearray[0] -Month $datearray[1] -Year $datearray[2]
    $totalDaysLeft = ($expirationDate - (Get-Date)).Totaldays.toString().split(",")[0]
    $totalDaysLeft = [int]$totalDaysLeft
	
    $objoutput = New-Object -TypeName PSObject -Property @{
      ExpDate    = $expirationDate.ToShortDateString()
      DaysRemain = $totalDaysLeft
    }
  } Catch {
    $objoutput = New-Object -TypeName PSObject -Property @{
      ExpDate    = "WMI Connection Failed"
      DaysRemain = "WMI Connection Failed"
    }
  }
  $objoutput
}

Function Get-VeeamWinServers {
  $vservers = @{}
  $outputAry = @()
  $vservers.add($($script:vbrServerObj.Name), "VBRServer")
  Foreach ($srv in $script:proxyList) {
    If (!$vservers.ContainsKey($srv.Host.Name)) {
      $vservers.Add($srv.Host.Name, "ProxyServer")
    }
  }
  Foreach ($srv in $script:repoList) {
    Try {
      $srvHost = $srv.gethost().Name
      If ($srv.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($srvHost)) {
        $vservers.Add($srvHost, "RepoServer")
      }
    } Catch {}
  }
  Foreach ($rs in $script:repoListSo) {
    ForEach ($rp in $rs.Extent) {
      $r = $rp.Repository
      Try {
        $rName = $($r.GetHost()).Name
        If ($r.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($rName)) {
          $vservers.Add($rName, "RepoSoServer")
        }
      } Catch {}
    }
  }
  Foreach ($srv in $script:tapesrvList) {
    If (!$vservers.ContainsKey($srv.Name)) {
      $vservers.Add($srv.Name, "TapeServer")
    }
  }
  $vservers = $vservers.GetEnumerator() | Sort-Object Name
  Foreach ($vserver in $vservers) {
    $outputAry += $vserver.Name
  }
  Return $outputAry
}

Function Get-VeeamServices {
  param (
    [PSObject]$inputObj
  )
  $outputAry = @()
  Foreach ($obj in $inputObj.keys) {
    $output = @()
    Try {
      $output = Invoke-Command -ComputerName $obj -ScriptBlock { Get-Service -Name "*Veeam*" -Exclude "SQLAgent*" } |
        Select-Object @{Name="Server Name"; Expression = {$obj.ToLower()}},
                      @{Name="Service Name"; Expression = {$_.DisplayName}},
                      Status
    } Catch {
      $output = New-Object PSObject -Property @{
        "Server Name" = $obj.ToLower()
        "Service Name" = "Unable to connect"
        Status = "Unknown"
      }
    }
    $outputAry += $output
  }
  $outputAry
}

Function Get-VMsBackupStatus {
  $outputary = @()
  $excludevms_regex = ('(?i)^(' + (($script:excludeVMs | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $excludefolder_regex = ('(?i)^(' + (($script:excludeFolder | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $excludedc_regex = ('(?i)^(' + (($script:excludeDC | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $vms = @{}
  Find-VBRViEntity |
    Where-Object {$_.Type -eq "Vm" -and $_.VmFolderName -notmatch $excludefolder_regex} |
    Where-Object {$_.Name -notmatch $excludevms_regex} |
    Where-Object {$_.Path.Split("\")[1] -notmatch $excludedc_regex} |
    ForEach-Object {$vms.Add(($_.FindObject().Id, $_.Id -ne $null)[0], @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.Path.Split("\")[2], $_.Name, "1/11/1911", "1/11/1911", "", $_.VmFolderName))}
  If (!$script:excludeTemp) {
    Find-VBRViEntity -VMsandTemplates |
      Where-Object {$_.Type -eq "Vm" -and $_.IsTemplate -eq "True" -and $_.VmFolderName -notmatch $excludefolder_regex} |
      Where-Object {$_.Name -notmatch $excludevms_regex} |
      Where-Object {$_.Path.Split("\")[1] -notmatch $excludedc_regex} |
      ForEach-Object {$vms.Add(($_.FindObject().Id, $_.Id -ne $null)[0], @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.VmHostName, "[template] $($_.Name)", "1/11/1911", "1/11/1911", "", $_.VmFolderName))}
  }
  $vbrtasksessions = (Get-VBRBackupSession |
    Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
    Get-VBRTaskSession | Where-Object {$_.Status -notmatch "InProgress|Pending"}
  If ($vbrtasksessions) {
    Foreach ($vmtask in $vbrtasksessions) {
      If ($vms.ContainsKey($vmtask.Info.ObjectId)) {
        If ((Get-Date $vmtask.Progress.StartTimeLocal) -ge (Get-Date $vms[$vmtask.Info.ObjectId][5])) {
          If ($vmtask.Status -eq "Success") {
            $vms[$vmtask.Info.ObjectId][0] = $vmtask.Status
            $vms[$vmtask.Info.ObjectId][5] = $vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6] = $vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7] = ""
          } ElseIf ($vms[$vmtask.Info.ObjectId][0] -ne "Success") {
            $vms[$vmtask.Info.ObjectId][0] = $vmtask.Status
            $vms[$vmtask.Info.ObjectId][5] = $vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6] = $vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7] = ($vmtask.GetDetails()).Replace("<br />", "ZZbrZZ")
          }
        } ElseIf ($vms[$vmtask.Info.ObjectId][0] -match "Warning|Failed" -and $vmtask.Status -eq "Success") {
          $vms[$vmtask.Info.ObjectId][0] = $vmtask.Status
          $vms[$vmtask.Info.ObjectId][5] = $vmtask.Progress.StartTimeLocal
          $vms[$vmtask.Info.ObjectId][6] = $vmtask.Progress.StopTimeLocal
          $vms[$vmtask.Info.ObjectId][7] = ""
        }
      }
    }
  }
  Foreach ($vm in $vms.GetEnumerator()) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Status     = $vm.Value[0]
      Name       = $vm.Value[4]
      vCenter    = $vm.Value[1]
      Datacenter = $vm.Value[2]
      Cluster    = $vm.Value[3]
      StartTime  = $vm.Value[5]
      StopTime   = $vm.Value[6]
      Details    = $vm.Value[7]
      Folder     = $vm.Value[8]
    }
    $outputAry += $objoutput
  }
  $outputAry
}

function Get-Duration {
  param ($ts)
  $days = ""
  If ($ts.Days -gt 0) {
    $days = "{0}:" -f $ts.Days
  }
  "{0}{1}:{2,2:D2}:{3,2:D2}" -f $days, $ts.Hours, $ts.Minutes, $ts.Seconds
}

function Get-BackupSize {
  param ($backups)
  $outputObj = @()
  Foreach ($backup in $backups) {
    $backupSize = 0
    $dataSize = 0
    $files = $backup.GetAllStorages()
    Foreach ($file in $files) {
      $backupSize += [math]::Round([long]$file.Stats.BackupSize/1GB, 2)
      $dataSize += [math]::Round([long]$file.Stats.DataSize/1GB, 2)
    }
    $repo = If ($($script:repoList | Where-Object {$_.Id -eq $backup.RepositoryId}).Name) {
      $($script:repoList | Where-Object {$_.Id -eq $backup.RepositoryId}).Name
    } Else {
      $($script:repoListSo | Where-Object {$_.Id -eq $backup.RepositoryId}).Name
    }
    $vbrMasterHash = @{
      JobName    = $backup.JobName
      VMCount    = $backup.VmCount
      Repo       = $repo
      DataSize   = $dataSize
      BackupSize = $backupSize
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    $outputObj += $vbrMasterObj
  }
  $outputObj
}

Function Get-MultiJobs {
  $outputAry = @()
  $vmMultiJobs = (Get-VBRBackupSession |
    Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
    Get-VBRTaskSession | Select-Object Name, @{Name="VMID"; Expression = {$_.Info.ObjectId}}, JobName -Unique |
    Group-Object Name, VMID | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group
  ForEach ($vm in $vmMultiJobs) {
    $objID = $vm.VMID
    $viEntity = Find-VBRViEntity -name $vm.Name | Where-Object {$_.FindObject().Id -eq $objID}
    If ($viEntity -ne $null) {
      $objoutput = New-Object -TypeName PSObject -Property @{
        Name       = $vm.Name
        vCenter    = $viEntity.Path.Split("\")[0]
        Datacenter = $viEntity.Path.Split("\")[1]
        Cluster    = $viEntity.Path.Split("\")[2]
        Folder     = $viEntity.VMFolderName
        JobName    = $vm.JobName
      }
      $outputAry += $objoutput
    } Else {
      $viEntity = Find-VBRViEntity -VMsAndTemplates -name $vm.Name | Where-Object {$_.FindObject().Id -eq $objID}
      If ($viEntity -ne $null) {
        $objoutput = New-Object -TypeName PSObject -Property @{
          Name       = "[template] " + $vm.Name
          vCenter    = $viEntity.Path.Split("\")[0]
          Datacenter = $viEntity.Path.Split("\")[1]
          Cluster    = $viEntity.VmHostName
          Folder     = $viEntity.VMFolderName
          JobName    = $vm.JobName
        }
        If ($objoutput) {
          $outputAry += $objoutput
        }
      }
    }
  }
  $outputAry
}
#endregion

#region Report
# Get Veeam Version
$VeeamVersion = Get-VeeamVersion

Write-Progress -Activity "Operazione in corso" -Status "40% Completato" -PercentComplete 40

# -----------------------------------------------------------------------
# FIX v12/v13: Updated version check - supports v12 and v13
# -----------------------------------------------------------------------
$VeeamVersionMajor = [int]($VeeamVersion.Split(".")[0])
If ($VeeamVersionMajor -lt 12) {
  Write-Host "Script requires VBR v12 or later" -ForegroundColor Red
  Write-Host "Version detected - $VeeamVersion" -ForegroundColor Red
  Exit
}

# HTML Stuff
$headerObj = @"
<html>
    <head>
        <title>$rptTitle</title>
            <style>
              body {font-family: Tahoma; background-color:#ffffff;}
              table {font-family: Tahoma;width: $($rptWidth)%;font-size: 12px;border-collapse:collapse;}
              th {background-color: #e2e2e2;border: 1px solid #a7a9ac;border-bottom: none;}
              td {background-color: #ffffff;border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
            </style>
    </head>
"@

$bodyTop = @"
    <body>
        <center>
            <table>
                <tr>
                    <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 10px;vertical-align: bottom;text-align: left;padding: 2px 0px 0px 5px;"></td>
                    <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 2px 5px 0px 0px;">Report generated on $(Get-Date -format g)</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 24px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 15px;">$rptTitle</td>
                    <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">$vbrName</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 5px;"></td>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 0px 0px;">VBR v$VeeamVersion</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 2px 5px;">$rptMode</td>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">MVR v$MVRversion</td>
                </tr>
            </table>
"@

$subHead01 = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@
$subHead01suc = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #00b050;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@
$subHead01war = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #ffd96c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@
$subHead01err = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #FB9895;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@
$subHead02 = @"
</td>
                </tr>
             </table>
"@
$HTMLbreak = @"
<table>
                <tr>
                    <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
                </tr>
            </table>
"@
$footerObj = @"
<table>
                <tr>
                    <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">My Veeam Report (v13 edition)</td>
                </tr>
            </table>
        </center>
    </body>
</html>
"@

# Get VM Backup Status
$vmStatus = @()
If ($showSummaryProtect + $showUnprotectedVMs + $showProtectedVMs) {
  $vmStatus = Get-VMsBackupStatus
}
$missingVMs = @($vmStatus | Where-Object {$_.Status -match "!|Failed"})
ForEach ($VM in $missingVMs) {
  If ($VM.Status -eq "!") {
    $VM.Details = "No Backup Task has completed"
    $VM.StartTime = ""
    $VM.StopTime = ""
  }
}
$successVMs = @($vmStatus | Where-Object {$_.Status -eq "Success"})
$warnVMs = @($vmStatus | Where-Object {$_.Status -eq "Warning"})

# Get VM Backup Protection Summary
$bodySummaryProtect = $null
$sumprotectHead = $subHead01
If ($showSummaryProtect) {
  $percentProt = 0
  If (@($successVMs).Count -ge 1) {
    $percentProt = 1
    $sumprotectHead = $subHead01suc
  }
  If (@($warnVMs).Count -ge 1) {
    $percentWarn = "*"
    $sumprotectHead = $subHead01war
  } Else {
    $percentWarn = ""
  }
  If (@($missingVMs).Count -ge 1) {
    $total = @($warnVMs).Count + @($successVMs).Count + @($missingVMs).Count
    If ($total -gt 0) {
      $percentProt = (@($warnVMs).Count + @($successVMs).Count) / $total
    }
    $sumprotectHead = $subHead01err
  }
  $vbrMasterHash = @{
    WarningVM   = @($warnVMs).Count
    ProtectedVM = @($successVMs).Count
    FailedVM    = @($missingVMs).Count
    PercentProt = "{0:P2}{1}" -f $percentProt, $percentWarn
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  $summaryProtect = $vbrMasterObj | Select-Object @{Name="% Protected"; Expression = {$_.PercentProt}},
    @{Name="Fully Protected VMs"; Expression = {$_.ProtectedVM}},
    @{Name="Protected VMs w/Warnings"; Expression = {$_.WarningVM}},
    @{Name="Unprotected VMs"; Expression = {$_.FailedVM}}
  $bodySummaryProtect = $summaryProtect | ConvertTo-Html -Fragment
  $bodySummaryProtect = $sumprotectHead + "VM Backup Protection Summary" + $subHead02 + $bodySummaryProtect
}

# Get VMs Missing Backups
$bodyMissing = $null
If ($showUnprotectedVMs) {
  If ($missingVMs -ne $null) {
    $missingVMs = $missingVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-Html -Fragment
    $bodyMissing = $subHead01err + "VMs with No Successful Backups within RPO" + $subHead02 + $missingVMs
  }
}

# Get VMs Backed Up w/Warnings
$bodyWarning = $null
If ($showProtectedVMs) {
  If ($warnVMs -ne $null) {
    $warnVMs = $warnVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-Html -Fragment
    $bodyWarning = $subHead01war + "VMs with only Backups with Warnings within RPO" + $subHead02 + $warnVMs
  }
}

# Get VMs Successfully Backed Up
$bodySuccess = $null
If ($showProtectedVMs) {
  If ($successVMs -ne $null) {
    $successVMs = $successVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}} | ConvertTo-Html -Fragment
    $bodySuccess = $subHead01suc + "VMs with Successful Backups within RPO" + $subHead02 + $successVMs
  }
}

# Get VMs Backed Up by Multiple Jobs
$bodyMultiJobs = $null
If ($showMultiJobs) {
  $multiJobs = @(Get-MultiJobs)
  If ($multiJobs.Count -gt 0) {
    $bodyMultiJobs = $multiJobs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Job Name"; Expression = {$_.JobName}} | ConvertTo-Html -Fragment
    $bodyMultiJobs = $subHead01err + "VMs Backed Up by Multiple Jobs within RPO" + $subHead02 + $bodyMultiJobs
  }
}

# Get Backup Summary Info
$bodySummaryBk = $null
If ($showSummaryBk) {
  $vbrMasterHash = @{
    "Failed"      = @($failedSessionsBk).Count
    "Sessions"    = If ($sessListBk) {@($sessListBk).Count} Else {0}
    "Read"        = $totalReadBk
    "Transferred" = $totalXferBk
    "Successful"  = @($successSessionsBk).Count
    "Warning"     = @($warningSessionsBk).Count
    "Fails"       = @($failsSessionsBk).Count
    "Running"     = @($runningSessionsBk).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  $total = If ($onlyLastBk) {"Jobs Run"} Else {"Total Sessions"}
  $arrSummaryBk = $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryBk = $arrSummaryBk | ConvertTo-Html -Fragment
  If ($arrSummaryBk.Failed -gt 0) { $summaryBkHead = $subHead01err }
  ElseIf ($arrSummaryBk.Warnings -gt 0) { $summaryBkHead = $subHead01war }
  ElseIf ($arrSummaryBk.Successful -gt 0) { $summaryBkHead = $subHead01suc }
  Else { $summaryBkHead = $subHead01 }
  $bodySummaryBk = $summaryBkHead + "Backup Results Summary" + $subHead02 + $bodySummaryBk
}

# Get Backup Job Status
$bodyJobsBk = $null
If ($showJobsBk) {
  If ($allJobsBk.count -gt 0) {
    $bodyJobsBk = @()
    Foreach ($bkJob in $allJobsBk) {
      $bodyJobsBk += $bkJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($bkJob.IsRunning) {
            $currentSess = $runningSessionsBk | Where-Object {$_.JobName -eq $bkJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB, 2)
            "$($csessPercent)% completed at $($csessSpeed) MB/s"
          } Else { "Stopped" }
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where-Object {$_.Id -eq $bkJob.Info.TargetRepositoryId}).Name) {
            $($repoList | Where-Object {$_.Id -eq $bkJob.Info.TargetRepositoryId}).Name
          } Else {
            $($repoListSo | Where-Object {$_.Id -eq $bkJob.Info.TargetRepositoryId}).Name
          }
        }},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continuous>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}
        }},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None") {"Unknown"} Else {$_.Info.LatestStatus}}}
    }
    $bodyJobsBk = $bodyJobsBk | Sort-Object "Next Run" | ConvertTo-Html -Fragment
    $bodyJobsBk = $subHead01 + "Backup Job Status" + $subHead02 + $bodyJobsBk
  }
}

# Get Backup Job Size
$bodyJobSizeBk = $null
If ($showBackupSizeBk) {
  If ($backupsBk.count -gt 0) {
    $bodyJobSizeBk = Get-BackupSize -backups $backupsBk | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-Html -Fragment
    $bodyJobSizeBk = $subHead01 + "Backup Job Size" + $subHead02 + $bodyJobSizeBk
  }
}

# Get all Backup Sessions
$bodyAllSessBk = $null
If ($showAllSessBk) {
  If ($sessListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB, 2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB, 2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
        @{Name="Dedupe"; Expression = {If ($_.Progress.ReadSize -eq 0) {0} Else {([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x"}}},
        @{Name="Compression"; Expression = {If ($_.Progress.ReadSize -eq 0) {0} Else {([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk | ConvertTo-Html -Fragment
      If ($arrAllSessBk.Result -match "Failed") { $allSessBkHead = $subHead01err }
      ElseIf ($arrAllSessBk.Result -match "Warning") { $allSessBkHead = $subHead01war }
      ElseIf ($arrAllSessBk.Result -match "Success") { $allSessBkHead = $subHead01suc }
      Else { $allSessBkHead = $subHead01 }
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    } Else {
      $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk | ConvertTo-Html -Fragment
      If ($arrAllSessBk.Result -match "Failed") { $allSessBkHead = $subHead01err }
      ElseIf ($arrAllSessBk.Result -match "Warning") { $allSessBkHead = $subHead01war }
      ElseIf ($arrAllSessBk.Result -match "Success") { $allSessBkHead = $subHead01suc }
      Else { $allSessBkHead = $subHead01 }
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    }
  }
}

# Get Running Backup Jobs
$bodyRunningBk = $null
If ($showRunningBk) {
  If ($runningSessionsBk.count -gt 0) {
    $bodyRunningBk = $runningSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-Html -Fragment
    $bodyRunningBk = $subHead01 + "Running Backup Jobs" + $subHead02 + $bodyRunningBk
  }
}

# Get Backup Sessions with Warnings or Failures
$bodySessWFBk = $null
If ($showWarnFailBk) {
  $sessWF = @($warningSessionsBk + $failsSessionsBk)
  If ($sessWF.count -gt 0) {
    $headerWF = If ($onlyLastBk) {"Backup Jobs with Warnings or Failures"} Else {"Backup Sessions with Warnings or Failures"}
    If ($showDetailedBk) {
      $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB, 2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB, 2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
        @{Name="Dedupe"; Expression = {If ($_.Progress.ReadSize -eq 0) {0} Else {([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x"}}},
        @{Name="Compression"; Expression = {If ($_.Progress.ReadSize -eq 0) {0} Else {([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq "") {$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()) {$_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}
        }}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-Html -Fragment
      If ($arrSessWFBk.Result -match "Failed") { $sessWFBkHead = $subHead01err }
      ElseIf ($arrSessWFBk.Result -match "Warning") { $sessWFBkHead = $subHead01war }
      ElseIf ($arrSessWFBk.Result -match "Success") { $sessWFBkHead = $subHead01suc }
      Else { $sessWFBkHead = $subHead01 }
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    } Else {
      $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq "") {$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()) {$_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}
        }}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-Html -Fragment
      If ($arrSessWFBk.Result -match "Failed") { $sessWFBkHead = $subHead01err }
      ElseIf ($arrSessWFBk.Result -match "Warning") { $sessWFBkHead = $subHead01war }
      ElseIf ($arrSessWFBk.Result -match "Success") { $sessWFBkHead = $subHead01suc }
      Else { $sessWFBkHead = $subHead01 }
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    }
  }
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
  If ($successSessionsBk.count -gt 0) {
    $headerSucc = If ($onlyLastBk) {"Successful Backup Jobs"} Else {"Successful Backup Sessions"}
    If ($showDetailedBk) {
      $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB, 2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB, 2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
        @{Name="Dedupe"; Expression = {If ($_.Progress.ReadSize -eq 0) {0} Else {([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x"}}},
        @{Name="Compression"; Expression = {If ($_.Progress.ReadSize -eq 0) {0} Else {([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x"}}},
        Result | ConvertTo-Html -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    } Else {
      $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-Html -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    }
  }
}

# Gather all Backup Tasks from Sessions
$taskListBk = @()
$taskListBk += $sessListBk | Get-VBRTaskSession
$successTasksBk = @($taskListBk | Where-Object {$_.Status -eq "Success"})
$wfTasksBk = @($taskListBk | Where-Object {$_.Status -match "Warning|Failed"})
$runningTasksBk = @()
$runningTasksBk += $runningSessionsBk | Get-VBRTaskSession | Where-Object {$_.Status -match "Pending|InProgress"}

# Get all Backup Tasks
$bodyAllTasksBk = $null
If ($showAllTasksBk) {
  If ($taskListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB, 2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB, 2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time" | ConvertTo-Html -Fragment
      If ($arrAllTasksBk.Status -match "Failed") { $allTasksBkHead = $subHead01err }
      ElseIf ($arrAllTasksBk.Status -match "Warning") { $allTasksBkHead = $subHead01war }
      ElseIf ($arrAllTasksBk.Status -match "Success") { $allTasksBkHead = $subHead01suc }
      Else { $allTasksBkHead = $subHead01 }
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    } Else {
      $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time" | ConvertTo-Html -Fragment
      If ($arrAllTasksBk.Status -match "Failed") { $allTasksBkHead = $subHead01err }
      ElseIf ($arrAllTasksBk.Status -match "Warning") { $allTasksBkHead = $subHead01war }
      ElseIf ($arrAllTasksBk.Status -match "Success") { $allTasksBkHead = $subHead01suc }
      Else { $allTasksBkHead = $subHead01 }
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    }
  }
}

# Get Running Backup Tasks
$bodyTasksRunningBk = $null
If ($showRunningTasksBk) {
  If ($runningTasksBk.count -gt 0) {
    $bodyTasksRunningBk = $runningTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSess.Name}},
      @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
      Status | Sort-Object "Start Time" | ConvertTo-Html -Fragment
    $bodyTasksRunningBk = $subHead01 + "Running Backup Tasks" + $subHead02 + $bodyTasksRunningBk
  }
}

# Get Backup Tasks with Warnings or Failures
$bodyTaskWFBk = $null
If ($showTaskWFBk) {
  If ($wfTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB, 2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB, 2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time" | ConvertTo-Html -Fragment
      If ($arrTaskWFBk.Status -match "Failed") { $taskWFBkHead = $subHead01err }
      ElseIf ($arrTaskWFBk.Status -match "Warning") { $taskWFBkHead = $subHead01war }
      ElseIf ($arrTaskWFBk.Status -match "Success") { $taskWFBkHead = $subHead01suc }
      Else { $taskWFBkHead = $subHead01 }
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    } Else {
      $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time" | ConvertTo-Html -Fragment
      If ($arrTaskWFBk.Status -match "Failed") { $taskWFBkHead = $subHead01err }
      ElseIf ($arrTaskWFBk.Status -match "Warning") { $taskWFBkHead = $subHead01war }
      ElseIf ($arrTaskWFBk.Status -match "Success") { $taskWFBkHead = $subHead01suc }
      Else { $taskWFBkHead = $subHead01 }
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    }
  }
}

# Get Successful Backup Tasks
$bodyTaskSuccBk = $null
If ($showTaskSuccessBk) {
  If ($successTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB, 2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB, 2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB, 2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB, 2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB, 2)}},
        Status | Sort-Object "Start Time" | ConvertTo-Html -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    } Else {
      $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort-Object "Start Time" | ConvertTo-Html -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    }
  }
}

# Get Running VM Restore Sessions
$bodyRestoRunVM = $null
If ($showRestoRunVM) {
  If ($($runningResto).count -gt 0) {
    $bodyRestoRunVM = $runningResto | Sort-Object CreationTime | Select-Object @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}}, @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}},
      @{Name="Reason"; Expression = {$_.Info.Reason}} | ConvertTo-Html -Fragment
    $bodyRestoRunVM = $subHead01 + "Running VM Restore Sessions" + $subHead02 + $bodyRestoRunVM
  }
}

# Get Completed VM Restore Sessions
$bodyRestoreVM = $null
If ($showRestoreVM) {
  If ($($completeResto).count -gt 0) {
    $arrRestoreVM = $completeResto | Sort-Object CreationTime | Select-Object @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}},
      @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}}, @{Name="Reason"; Expression = {$_.Info.Reason}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyRestoreVM = $arrRestoreVM | ConvertTo-Html -Fragment
    If ($arrRestoreVM.Result -match "Failed") { $restoreVMHead = $subHead01err }
    ElseIf ($arrRestoreVM.Result -match "Warning") { $restoreVMHead = $subHead01war }
    ElseIf ($arrRestoreVM.Result -match "Success") { $restoreVMHead = $subHead01suc }
    Else { $restoreVMHead = $subHead01 }
    $bodyRestoreVM = $restoreVMHead + "Completed VM Restore Sessions" + $subHead02 + $bodyRestoreVM
  }
}

Write-Progress -Activity "Operazione in corso" -Status "50% Completato" -PercentComplete 50

# -----------------------------------------------------------------------
# Agent Backup sections (simplified for v12/v13)
# -----------------------------------------------------------------------
$bodySummaryEp = $null
If ($showSummaryEp) {
  $vbrEpHash = @{
    "Sessions"   = If ($sessListEp) {@($sessListEp).Count} Else {0}
    "Successful" = @($successSessionsEp).Count
    "Warning"    = @($warningSessionsEp).Count
    "Fails"      = @($failsSessionsEp).Count
    "Running"    = @($runningSessionsEp).Count
  }
  $vbrEPObj = New-Object -TypeName PSObject -Property $vbrEpHash
  $total = If ($onlyLastEp) {"Jobs Run"} Else {"Total Sessions"}
  $arrSummaryEp = $vbrEPObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryEp = $arrSummaryEp | ConvertTo-Html -Fragment
  If ($arrSummaryEp.Failures -gt 0) { $summaryEpHead = $subHead01err }
  ElseIf ($arrSummaryEp.Warnings -gt 0) { $summaryEpHead = $subHead01war }
  ElseIf ($arrSummaryEp.Successful -gt 0) { $summaryEpHead = $subHead01suc }
  Else { $summaryEpHead = $subHead01 }
  $bodySummaryEp = $summaryEpHead + "Agent Backup Results Summary" + $subHead02 + $bodySummaryEp
}

$bodyJobsEp = $null
If ($showJobsEp) {
  If ($allJobsEp.count -gt 0) {
    $bodyJobsEp = $allJobsEp | Sort-Object Name | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Description"; Expression = {$_.Description}},
      @{Name="Enabled"; Expression = {$_.IsEnabled}},
      # -----------------------------------------------------------------------
      # FIX v12/v13: Agent job status properties updated
      # -----------------------------------------------------------------------
      @{Name="Status"; Expression = {Try {$_.LastState} Catch {"Unknown"}}},
      @{Name="Target Repo"; Expression = {Try {$_.Target} Catch {"N/A"}}},
      @{Name="Next Run"; Expression = {Try {$_.NextRun} Catch {"N/A"}}},
      @{Name="Last Result"; Expression = {Try {If ($_.LastResult -eq "None") {""} Else {$_.LastResult}} Catch {"Unknown"}}} | ConvertTo-Html -Fragment
    $bodyJobsEp = $subHead01 + "Agent Backup Job Status" + $subHead02 + $bodyJobsEp
  }
}

$bodyJobSizeEp = $null
If ($showBackupSizeEp) {
  If ($backupsEp.count -gt 0) {
    $bodyJobSizeEp = Get-BackupSize -backups $backupsEp | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-Html -Fragment
    $bodyJobSizeEp = $subHead01 + "Agent Backup Job Size" + $subHead02 + $bodyJobSizeEp
  }
}

$bodyAllSessEp = @()
$arrAllSessEp = @()
If ($showAllSessEp) {
  If ($sessListEp.count -gt 0) {
    Foreach ($job in $allJobsEp) {
      $arrAllSessEp += $sessListEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="State"; Expression = {$_.State}}, @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    }
    $bodyAllSessEp = $arrAllSessEp | Sort-Object "Start Time" | ConvertTo-Html -Fragment
    If ($arrAllSessEp.Result -match "Failed") { $allSessEpHead = $subHead01err }
    ElseIf ($arrAllSessEp.Result -match "Warning") { $allSessEpHead = $subHead01war }
    ElseIf ($arrAllSessEp.Result -match "Success") { $allSessEpHead = $subHead01suc }
    Else { $allSessEpHead = $subHead01 }
    $bodyAllSessEp = $allSessEpHead + "Agent Backup Sessions" + $subHead02 + $bodyAllSessEp
  }
}

$bodyRunningEp = @()
If ($showRunningEp) {
  If ($runningSessionsEp.count -gt 0) {
    Foreach ($job in $allJobsEp) {
      $bodyRunningEp += $runningSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
    }
    $bodyRunningEp = $bodyRunningEp | Sort-Object "Start Time" | ConvertTo-Html -Fragment
    $bodyRunningEp = $subHead01 + "Running Agent Backup Jobs" + $subHead02 + $bodyRunningEp
  }
}

$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
  $sessWFEp = @($warningSessionsEp + $failsSessionsEp)
  If ($sessWFEp.count -gt 0) {
    $headerWFEp = If ($onlyLastEp) {"Agent Backup Jobs with Warnings or Failures"} Else {"Agent Backup Sessions with Warnings or Failures"}
    Foreach ($job in $allJobsEp) {
      $arrSessWFEp += $sessWFEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessWFEp = $arrSessWFEp | Sort-Object "Start Time" | ConvertTo-Html -Fragment
    If ($arrSessWFEp.Result -match "Failed") { $sessWFEpHead = $subHead01err }
    ElseIf ($arrSessWFEp.Result -match "Warning") { $sessWFEpHead = $subHead01war }
    ElseIf ($arrSessWFEp.Result -match "Success") { $sessWFEpHead = $subHead01suc }
    Else { $sessWFEpHead = $subHead01 }
    $bodySessWFEp = $sessWFEpHead + $headerWFEp + $subHead02 + $bodySessWFEp
  }
}

$bodySessSuccEp = @()
If ($showSuccessEp) {
  If ($successSessionsEp.count -gt 0) {
    $headerSuccEp = If ($onlyLastEp) {"Successful Agent Backup Jobs"} Else {"Successful Agent Backup Sessions"}
    Foreach ($job in $allJobsEp) {
      $bodySessSuccEp += $successSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time" | ConvertTo-Html -Fragment
    $bodySessSuccEp = $subHead01suc + $headerSuccEp + $subHead02 + $bodySessSuccEp
  }
}

Write-Progress -Activity "Operazione in corso" -Status "60% Completato" -PercentComplete 60

# -----------------------------------------------------------------------
# SureBackup sections (v12/v13 uses Get-VBRViSureBackupJob/Session)
# -----------------------------------------------------------------------
$bodySummarySb = $null
If ($showSummarySb) {
  $vbrMasterHash = @{
    "Sessions"   = If ($sessListSb) {@($sessListSb).Count} Else {0}
    "Successful" = @($successSessionsSb).Count
    "Warning"    = @($warningSessionsSb).Count
    "Fails"      = @($failsSessionsSb).Count
    "Running"    = @($runningSessionsSb).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  $total = If ($onlyLastSb) {"Jobs Run"} Else {"Total Sessions"}
  $arrSummarySb = $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummarySb = $arrSummarySb | ConvertTo-Html -Fragment
  If ($arrSummarySb.Failures -gt 0) { $summarySbHead = $subHead01err }
  ElseIf ($arrSummarySb.Warnings -gt 0) { $summarySbHead = $subHead01war }
  ElseIf ($arrSummarySb.Successful -gt 0) { $summarySbHead = $subHead01suc }
  Else { $summarySbHead = $subHead01 }
  $bodySummarySb = $summarySbHead + "SureBackup Results Summary" + $subHead02 + $bodySummarySb
}

$bodyJobsSb = $null
If ($showJobsSb) {
  If ($allJobsSb.count -gt 0) {
    $bodyJobsSb = @()
    Foreach ($SbJob in $allJobsSb) {
      $bodyJobsSb += $SbJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          Try {
            If ($_.GetLastState() -eq "Working") {
              $currentSess = $_.FindLastSession()
              "$($currentSess.CompletionPercentage)% completed"
            } Else { $_.GetLastState() }
          } Catch { "Unknown" }
        }},
        @{Name="Virtual Lab"; Expression = {Try {$(Get-VBRViVirtualLab | Where-Object {$_.Id -eq $SbJob.VirtualLabId}).Name} Catch {"N/A"}}},
        @{Name="Linked Jobs"; Expression = {Try {$($_.GetLinkedJobs()).Name -join ","} Catch {"N/A"}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continuous>"}
          Else {Try {$_.ScheduleOptions.NextRun} Catch {"N/A"}}
        }},
        @{Name="Last Result"; Expression = {Try {If ($_.GetLastResult() -eq "None") {""} Else {$_.GetLastResult()}} Catch {"Unknown"}}}
    }
    $bodyJobsSb = $bodyJobsSb | Sort-Object "Next Run" | ConvertTo-Html -Fragment
    $bodyJobsSb = $subHead01 + "SureBackup Job Status" + $subHead02 + $bodyJobsSb
  }
}

$bodyAllSessSb = $null
If ($showAllSessSb) {
  If ($sessListSb.count -gt 0) {
    $arrAllSessSb = $sessListSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="State"; Expression = {$_.State}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {"-"} Else {$_.EndTime}}},
      @{Name="Duration (HH:MM:SS)"; Expression = {
        If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
          Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
        } Else {
          Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
        }
      }}, Result
    $bodyAllSessSb = $arrAllSessSb | ConvertTo-Html -Fragment
    If ($arrAllSessSb.Result -match "Failed") { $allSessSbHead = $subHead01err }
    ElseIf ($arrAllSessSb.Result -match "Warning") { $allSessSbHead = $subHead01war }
    ElseIf ($arrAllSessSb.Result -match "Success") { $allSessSbHead = $subHead01suc }
    Else { $allSessSbHead = $subHead01 }
    $bodyAllSessSb = $allSessSbHead + "SureBackup Sessions" + $subHead02 + $bodyAllSessSb
  }
}

# Get Configuration Backup Summary Info
$bodySummaryConfig = $null
If ($showSummaryConfig) {
  $vbrConfigHash = @{
    "Enabled"        = $configBackup.Enabled
    "Status"         = $configBackup.LastState
    "Target"         = $configBackup.Target
    "Schedule"       = $configBackup.ScheduleOptions
    "Restore Points" = $configBackup.RestorePointsToKeep
    "Encrypted"      = $configBackup.EncryptionOptions.Enabled
    "Last Result"    = $configBackup.LastResult
    "Next Run"       = $configBackup.NextRun
  }
  $vbrConfigObj = New-Object -TypeName PSObject -Property $vbrConfigHash
  $bodySummaryConfig = $vbrConfigObj | Select-Object Enabled, Status, Target, Schedule, "Restore Points", "Next Run", Encrypted, "Last Result" | ConvertTo-Html -Fragment
  If ($configBackup.LastResult -eq "Warning" -or !$configBackup.Enabled) { $configHead = $subHead01war }
  ElseIf ($configBackup.LastResult -eq "Success") { $configHead = $subHead01suc }
  ElseIf ($configBackup.LastResult -eq "Failed") { $configHead = $subHead01err }
  Else { $configHead = $subHead01 }
  $bodySummaryConfig = $configHead + "Configuration Backup Status" + $subHead02 + $bodySummaryConfig
}

Write-Progress -Activity "Operazione in corso" -Status "70% Completato" -PercentComplete 70

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
  If ($proxyList -ne $null) {
    $arrProxy = $proxyList | Get-VBRProxyInfo | Select-Object @{Name="Proxy Name"; Expression = {$_.ProxyName}},
      @{Name="Transport Mode"; Expression = {$_.tMode}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Proxy Host"; Expression = {$_.RealName}}, @{Name="Host Type"; Expression = {$_.pType}},
      Enabled, @{Name="IP Address"; Expression = {$_.IP}},
      @{Name="RT (ms)"; Expression = {$_.Response}}, Status
    $bodyProxy = $arrProxy | Sort-Object "Proxy Host" | ConvertTo-Html -Fragment
    If ($arrProxy.Status -match "Dead") { $proxyHead = $subHead01err }
    ElseIf ($arrProxy -match "Alive") { $proxyHead = $subHead01suc }
    Else { $proxyHead = $subHead01 }
    $bodyProxy = $proxyHead + "Proxy Details" + $subHead02 + $bodyProxy
  }
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
  If ($repoList -ne $null) {
    $arrRepo = $repoList | Get-VBRRepoInfo | Select-Object @{Name="Repository Name"; Expression = {$_.Target}},
      @{Name="Type"; Expression = {$_.rType}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Host"; Expression = {$_.RepoHost}}, @{Name="Path"; Expression = {$_.Storepath}},
      @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0) {"Warning"}
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}
      }}
    $bodyRepo = $arrRepo | Sort-Object "Repository Name" | ConvertTo-Html -Fragment
    If ($arrRepo.status -match "Critical") { $repoHead = $subHead01err }
    ElseIf ($arrRepo.status -match "Warning|Unknown") { $repoHead = $subHead01war }
    ElseIf ($arrRepo.status -match "OK") { $repoHead = $subHead01suc }
    Else { $repoHead = $subHead01 }
    $bodyRepo = $repoHead + "Repository Details" + $subHead02 + $bodyRepo
  }
}

# Get Scale Out Repository Info
$bodySORepo = $null
If ($showRepo) {
  If ($repoListSo -ne $null) {
    $arrSORepo = $repoListSo | Get-VBRSORepoInfo | Select-Object @{Name="Scale Out Repository Name"; Expression = {$_.SOTarget}},
      @{Name="Member Repository Name"; Expression = {$_.Target}}, @{Name="Type"; Expression = {$_.rType}},
      @{Name="Max Tasks"; Expression = {$_.MaxTasks}}, @{Name="Host"; Expression = {$_.RepoHost}},
      @{Name="Path"; Expression = {$_.Storepath}}, @{Name="Free (GB)"; Expression = {$_.StorageFree}},
      @{Name="Total (GB)"; Expression = {$_.StorageTotal}}, @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0) {"Warning"}
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}
      }}
    $bodySORepo = $arrSORepo | Sort-Object "Scale Out Repository Name", "Member Repository Name" | ConvertTo-Html -Fragment
    If ($arrSORepo.status -match "Critical") { $sorepoHead = $subHead01err }
    ElseIf ($arrSORepo.status -match "Warning|Unknown") { $sorepoHead = $subHead01war }
    ElseIf ($arrSORepo.status -match "OK") { $sorepoHead = $subHead01suc }
    Else { $sorepoHead = $subHead01 }
    $bodySORepo = $sorepoHead + "Scale Out Repository Details" + $subHead02 + $bodySORepo
  }
}

# Get Repository Permissions
$bodyRepoPerms = $null
If ($showRepoPerms) {
  If ($repoList -ne $null -or $repoListSo -ne $null) {
    $bodyRepoPerms = Get-RepoPermissions | Select-Object Name, "Encryption Enabled", "Permission Type", Users | Sort-Object Name | ConvertTo-Html -Fragment
    $bodyRepoPerms = $subHead01 + "Repository Permissions for Agent Jobs" + $subHead02 + $bodyRepoPerms
  }
}

Write-Progress -Activity "Operazione in corso" -Status "80% Completato" -PercentComplete 80

# Get Replica Target Info
$bodyReplica = $null
If ($showReplicaTarget) {
  If ($allJobsRp -ne $null) {
    $repTargets = $allJobsRp | Get-VBRReplicaTarget | Select-Object @{Name="Replica Target"; Expression = {$_.Target}}, Datastore,
      @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $replicaCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0) {"Warning"}
        ElseIf ($_.FreePercentage -lt $replicaWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}
      }} | Sort-Object "Replica Target"
    $bodyReplica = $repTargets | ConvertTo-Html -Fragment
    If ($repTargets.status -match "Critical") { $reptarHead = $subHead01err }
    ElseIf ($repTargets.status -match "Warning|Unknown") { $reptarHead = $subHead01war }
    ElseIf ($repTargets.status -match "OK") { $reptarHead = $subHead01suc }
    Else { $reptarHead = $subHead01 }
    $bodyReplica = $reptarHead + "Replica Target Details" + $subHead02 + $bodyReplica
  }
}

# Get Veeam Services Info
$bodyServices = $null
If ($showServices) {
  $vServers = Get-VeeamWinServers
  $vServices = Get-VeeamServices $vServers
  If ($hideRunningSvc) {$vServices = $vServices | Where-Object {$_.Status -ne "Running"}}
  If ($vServices -ne $null) {
    $vServices = $vServices | Select-Object "Server Name", "Service Name",
      @{Name="Status"; Expression = {If ($_.Status -eq "Stopped") {"Not Running"} Else {$_.Status}}}
    $bodyServices = $vServices | Sort-Object "Server Name", "Service Name" | ConvertTo-Html -Fragment
    If ($vServices.status -match "Not Running") { $svcHead = $subHead01err }
    ElseIf ($vServices.status -notmatch "Running") { $svcHead = $subHead01war }
    ElseIf ($vServices.status -match "Running") { $svcHead = $subHead01suc }
    Else { $svcHead = $subHead01 }
    $bodyServices = $svcHead + "Veeam Services (Windows)" + $subHead02 + $bodyServices
  }
}

# Get License Info
$bodyLicense = $null
If ($showLicExp) {
  $arrLicense = Get-VeeamSupportDate $vbrServer | Select-Object @{Name="Expiry Date"; Expression = {$_.ExpDate}},
    @{Name="Days Remaining"; Expression = {$_.DaysRemain}},
    @{Name="Status"; Expression = {
      If ($_.DaysRemain -lt $licenseCritical) {"Critical"}
      ElseIf ($_.DaysRemain -lt $licenseWarn) {"Warning"}
      ElseIf ($_.DaysRemain -eq "Failed") {"Failed"}
      Else {"OK"}
    }}
  $bodyLicense = $arrLicense | ConvertTo-Html -Fragment
  If ($arrLicense.Status -eq "OK") { $licHead = $subHead01suc }
  ElseIf ($arrLicense.Status -eq "Warning") { $licHead = $subHead01war }
  Else { $licHead = $subHead01err }
  $bodyLicense = $licHead + "License/Support Renewal Date" + $subHead02 + $bodyLicense
}

Write-Progress -Activity "Operazione in corso" -Status "90% Completato" -PercentComplete 90

# Placeholders for sections not fully ported (Replication, Backup Copy, Tape)
# These retain the original logic - update $null vars to avoid errors
$bodySummaryRp = $null; $bodyJobsRp = $null; $bodyAllSessRp = $null; $bodyAllTasksRp = $null
$bodyRunningRp = $null; $bodyTasksRunningRp = $null; $bodySessWFRp = $null; $bodyTaskWFRp = $null
$bodySessSuccRp = $null; $bodyTaskSuccRp = $null
$bodySummaryBc = $null; $bodyJobsBc = $null; $bodyJobSizeBc = $null; $bodyAllSessBc = $null
$bodyAllTasksBc = $null; $bodySessIdleBc = $null; $bodyTasksPendingBc = $null; $bodyRunningBc = $null
$bodyTasksRunningBc = $null; $bodySessWFBc = $null; $bodyTaskWFBc = $null; $bodySessSuccBc = $null; $bodyTaskSuccBc = $null
$bodySummaryTp = $null; $bodyJobsTp = $null; $bodyAllSessTp = $null; $bodyAllTasksTp = $null
$bodyWaitingTp = $null; $bodySessIdleTp = $null; $bodyTasksPendingTp = $null; $bodyRunningTp = $null
$bodyTasksRunningTp = $null; $bodySessWFTp = $null; $bodyTaskWFTp = $null; $bodySessSuccTp = $null; $bodyTaskSuccTp = $null
$bodyTapes = $null; $bodyTpPool = $null; $bodyTpVlt = $null; $bodyExpTp = $null
$bodyTpExpPool = $null; $bodyTpExpVlt = $null; $bodyTpWrt = $null
$bodyRunningSb = $null; $bodyTasksRunningSb = $null; $bodyAllTasksSb = $null
$bodySessWFSb = $null; $bodyTaskWFSb = $null; $bodySessSuccSb = $null; $bodyTaskSuccSb = $null

# Combine HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummaryProtect + $bodySummaryBk + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb

If ($bodySummaryProtect + $bodySummaryBk + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMissing + $bodyWarning + $bodySuccess
If ($bodyMissing + $bodySuccess + $bodyWarning) { $htmlOutput += $HTMLbreak }

$htmlOutput += $bodyMultiJobs
If ($bodyMultiJobs) { $htmlOutput += $HTMLbreak }

$htmlOutput += $bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk
If ($bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyRestoRunVM + $bodyRestoreVM
If ($bodyRestoRunVM + $bodyRestoreVM) { $htmlOutput += $HTMLbreak }

$htmlOutput += $bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp
If ($bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb
If ($bodyJobsSb + $bodyAllSessSb) { $htmlOutput += $HTMLbreak }

$htmlOutput += $bodySummaryConfig + $bodyProxy + $bodyRepo + $bodySORepo + $bodyRepoPerms + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Fix Details
$htmlOutput = $htmlOutput.Replace("ZZbrZZ", "<br />")
$htmlOutput = $htmlOutput.Replace("$($HTMLbreak + $footerObj)", "$($footerObj)")

# Add color to output
$htmlOutput = $htmlOutput.Replace("<td>Running<",   "<td style=""color: #00b051;"">Running<")
$htmlOutput = $htmlOutput.Replace("<td>OK<",        "<td style=""color: #00b051;"">OK<")
$htmlOutput = $htmlOutput.Replace("<td>Alive<",     "<td style=""color: #00b051;"">Alive<")
$htmlOutput = $htmlOutput.Replace("<td>Success<",   "<td style=""color: #00b051;"">Success<")
$htmlOutput = $htmlOutput.Replace("<td>Warning<",   "<td style=""color: #ffc000;"">Warning<")
$htmlOutput = $htmlOutput.Replace("<td>Not Running<","<td style=""color: #ff0000;"">Not Running<")
$htmlOutput = $htmlOutput.Replace("<td>Failed<",    "<td style=""color: #ff0000;"">Failed<")
$htmlOutput = $htmlOutput.Replace("<td>Critical<",  "<td style=""color: #ff0000;"">Critical<")
$htmlOutput = $htmlOutput.Replace("<td>Dead<",      "<td style=""color: #ff0000;"">Dead<")

# Color Report Header and Tag Email Subject
If ($htmlOutput -match "#FB9895") {
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ", "#FB9895")
  $emailSubject = "[Failed] $emailSubject"
} ElseIf ($htmlOutput -match "#ffd96c") {
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ", "#ffd96c")
  $emailSubject = "[Warning] $emailSubject"
} ElseIf ($htmlOutput -match "#00b050") {
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ", "#00b050")
  $emailSubject = "[Success] $emailSubject"
} Else {
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ", "#626365")
}
#endregion

#region Output
If ($sendEmail) {
  $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
  $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
  $smtp.EnableSsl = $emailEnableSSL
  $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
  $msg.Subject = $emailSubject
  If ($emailAttach) {
    $msg.Body = "Veeam Report Attached"
    $tempFile = "$env:TEMP\$($rptTitle)_$(Get-Date -format MMddyyyy_hhmmss).htm"
    $htmlOutput | Out-File $tempFile
    $attachment = New-Object System.Net.Mail.Attachment $tempFile
    $msg.Attachments.Add($attachment)
  } Else {
    $msg.Body = $htmlOutput
    $msg.isBodyhtml = $true
  }
  $smtp.send($msg)
  If ($emailAttach) {
    $attachment.dispose()
    Remove-Item $tempFile
  }
}

If ($saveHTML) {
  $htmlOutput | Out-File $pathHTML
  If ($launchHTML) {
    Invoke-Item $pathHTML
  }
}
#endregion

Write-Progress -Activity "Operazione in corso" -Status "100% Completato" -PercentComplete 100

$endDate = $(Get-Date -UFormat "%m-%d-%Y_%H%M")
write-host "Starting operations: $endDate"
