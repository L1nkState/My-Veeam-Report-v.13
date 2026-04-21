
# My Veeam Report v13 modified ny [@L1nkState](https://github.com/L1nkState)

**My Veeam Report** is an advanced and highly configurable PowerShell script that generates comprehensive **HTML reports** for **Veeam Backup & Replication**.

This project originates from the original script created by **Shawn Masterson**, which was abandoned in 2018, and has been modernized and updated thanks to the community (especially **marcohorstmann**) to support modern Veeam versions (v12 and v13).

This script has long been – and still is – an excellent learning tool to understand and work with the **Veeam PowerShell API** and its internal logic.

***

## 📌 Key Features

*   Compatible with **Veeam Backup & Replication v12 and v13**
*   Clean and readable HTML report
*   Automatic email delivery
*   Extremely granular configuration (each report section can be enabled/disabled)
*   Supports **enterprise** and **MSP** environments

***

## ✅ Prerequisites

<img width="594" height="267" alt="image" src="https://github.com/user-attachments/assets/d04dcb52-70df-45bb-bd52-ffdd7587b520" />


Note: The script must be run on the Veeam Backup Server or a machine with the Veeam console and PowerShell snap-in installed.



### 1. Operating System

*   Windows Server or Windows Client supported by Veeam
*   Network connectivity to the VBR server

***

### 2. Veeam Backup & Replication

One of the following must be installed on the system running the script:

*   Veeam Backup & Replication **Server**
*   **Veeam Backup & Replication Console**

Supported versions:

*   ✅ VBR **v12**
*   ✅ VBR **v13**

***

### 3. PowerShell

#### Windows PowerShell (classic)

*   **Minimum version:** PowerShell **5.0**
*   Check version:

```powershell
$PSVersionTable.PSVersion
```

#### ✅ PowerShell 7 (recommended)

The script is also compatible with **PowerShell 7 (Core)**.

Additional requirements:

*   PowerShell 7 installed on the server
*   **Veeam Console must be installed** (required to load the Veeam module)

Check version:

```powershell
pwsh --version
```

⚠️ Note:

> PowerShell 7 does **not automatically include** Veeam modules.  
> This script implements a **multi-strategy loading mechanism**:
>
> *   Auto import by module name
> *   Known installation paths
> *   Recursive search under `Program Files`
> *   Legacy PSSnapIn fallback (v9.5 / v11)

***

### 4. Permissions

*   Run the script with **administrative privileges**
*   User must have access to Veeam Backup & Replication
*   If using the “Services” section, **PowerShell Remoting must be enabled**

***

## ⚙️ Initial Configuration

All configuration options are defined in:

```powershell
#region User-Variables
```
## 🗓️ Scheduling with Windows Task Scheduler

You can automate the validator using **Windows Task Scheduler**:

1. Open **Task Scheduler** and create a new task.
2. Set the **Trigger** (e.g., daily at 06:00 AM after backup jobs complete).
3. Set the **Action** to:
   - **Program:** `powershell.exe`
   - **Arguments:**
     ```
     -NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\veeam-BackupValidator.ps1" "YOUR-JOB" "YOUR-SERVER"
     ```
4. Run the task as a service account with **Veeam Operator** (or higher) permissions.
5. Enable **"Run whether user is logged on or not"**.

<img width="632" height="478" alt="image" src="https://github.com/user-attachments/assets/8a474715-02f6-4750-ad18-e6150a09202d" />

<img width="633" height="480" alt="image" src="https://github.com/user-attachments/assets/e112bf16-7a61-4587-932d-7287b331a449" />

Program/script: "C:\Program Files\PowerShell\7\pwsh.exe" Add arguments: -file "C:\Your-dir-script\\MyVeeamReport.ps1

<img width="633" height="657" alt="image" src="https://github.com/user-attachments/assets/cc2f0db4-9e24-47cd-a986-f3d063a0474d" />


> **Important:** Ensure the executing account has rights to run Veeam PowerShell cmdlets and access `Veeam.Backup.Validator.exe`.

Key parameters:

*   `$vbrServer` → VBR server name / FQDN
*   `$reportMode` → `24`, `48`, `Weekly`, `Monthly`
*   `$saveHTML`, `$pathHTML`
*   `$sendEmail`, `$emailHost`, `$emailTo`
*   Enable/disable **each report section individually**

***

## 📊 Report Capabilities

### ✅ VM Protection Overview

*   VM protection status (Success / Warning / Failed)
*   VMs without backups within RPO
*   VMs with only Warning backups
*   Overall protection percentage
*   Exclusions by:
    *   VM name
    *   vCenter folder
    *   Datacenter
    *   Templates

***

### 🔁 Backup Jobs

*   Job status
*   Backup sessions
*   Per‑VM tasks
*   Running jobs
*   Warning / Failed sessions
*   Last session only option
*   Backup size analysis (Data Size vs Backup Size)

***

### 💻 Agent Backup

*   Windows/Linux Agent jobs
*   Agent backup sessions
*   State, duration, results
*   Backup size reporting
*   Fully aligned with v12/v13 APIs

***

### 🔄 Replication (partial)

*   Replication sessions
*   Job status
*   Target datastore space utilization

***

### 🧪 SureBackup

*   SureBackup jobs
*   Sessions
*   Status and results
*   Virtual Lab and linked jobs

***

### 📦 Veeam Infrastructure

*   **Proxies**
    *   Alive/Dead status
    *   Transport mode
    *   Max tasks
    *   Ping and latency
*   **Repositories**
    *   Standard repositories
    *   Scale-Out repositories
    *   Types (Windows, Linux, Hardened, S3, Azure, Wasabi, Quantum…)
    *   Free space with warning/critical thresholds
*   **Repository Permissions**
    *   Agent job permissions
    *   Encryption status

***

### 🎯 Replica Targets

*   Target datastore
*   Free space
*   Configurable warning/critical thresholds

***

### 🪪 License Status

*   License expiration date
*   Remaining days
*   Status:
    *   OK
    *   Warning
    *   Critical

***

### ⚙️ VBR Configuration Backup

*   Configuration backup status
*   Schedule
*   Repository
*   Restore points
*   Encryption status

***

### 📧 Output Options

*   HTML report:
    *   Saved to file
    *   Optional auto‑open
*   Email:
    *   HTML body or attachment
    *   Dynamic subject line (Success / Warning / Failed)

***

## 🚀 Running the Script

```powershell
pwsh .\MyVeeamReport_v13.ps1
```

or:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\MyVeeamReport_v13.ps1
```

***

## 🤝 Contributing

Contributions are **very welcome** ✅

You can:

*   Open **Issues** for bugs or enhancement requests
*   Submit **Pull Requests**
*   Improve documentation
*   Extend compatibility with future Veeam releases

***

## 🙏 Credits

*   **Shawn Masterson** – original script author
*   **marcohorstmann** and the community – modernization and updates
*   Veeam PowerShell community contributors

***

[![PowerShell](https://img.shields.io/badge/PowerShell-5.0%2B%20%7C%207-blue.svg?logo=powershell)](https://learn.microsoft.com/powershell/)
[![Veeam](https://img.shields.io/badge/Veeam-Backup%20%26%20Replication%20v12%20%7C%20v13-brightgreen)](https://www.veeam.com/)
[![Version](https://img.shields.io/badge/Version-13.0.0-blue)](https://github.com/<YOUR_ORG>/<REPO_NAME>/releases)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
``
