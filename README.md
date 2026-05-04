# ⚡ PowerShell Automation Scripts — Windows Server 2022

Home Lab Project — Part 06

![Windows Server](https://img.shields.io/badge/Windows%20Server-2022-blue?style=flat-square&logo=windows)
![PowerShell](https://img.shields.io/badge/PowerShell-Automation-blue?style=flat-square&logo=powershell)
![Active Directory](https://img.shields.io/badge/Active%20Directory-Integrated-green?style=flat-square)
![PS2EXE](https://img.shields.io/badge/PS2EXE-Compiled-orange?style=flat-square)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen?style=flat-square)

> **Home Lab Project — Part 6**
> Building enterprise-grade PowerShell automation scripts for Active Directory management — bulk user creation from CSV, AD user reporting, password expiry checking, security summary generation, and compiling scripts into standalone executable files using PS2EXE.

---

## 📑 Table of Contents

- [Lab Environment](#lab-environment)
- [Objectives](#objectives)
- [Why PowerShell Automation?](#why-powershell-automation)
- [Script 1 — Bulk AD User Creation](#script-1--bulk-ad-user-creation-from-csv)
- [Script 2 — AD User Report Generator](#script-2--ad-user-report-generator)
- [Script 3 — Password Expiry Checker](#script-3--password-expiry-checker)
- [Script 4 — Master Report Generator](#script-4--master-report-generator-exe)
- [EXE Compilation](#exe-compilation-with-ps2exe)
- [Output Files](#output-files)
- [Skills Demonstrated](#skills-demonstrated)
- [Related Projects](#related-projects)

---

## 🏗️ Lab Environment

| Component | Details |
|---|---|
| **Hypervisor** | Oracle VirtualBox |
| **Domain Controller** | WIN-6ARV79FMB48 — Windows Server 2022 |
| **Domain** | mylab.local |
| **Server IP** | 192.168.1.200 |
| **OU Targeted** | OU=VMUsers,DC=mylab,DC=local |
| **PS2EXE Version** | 1.0.17 |

---

## 🎯 Objectives

- [x] Create bulk user CSV file with department data
- [x] Build Script 1 — Bulk AD User Creation from CSV
- [x] Build Script 2 — AD User Report Generator
- [x] Build Script 3 — Password Expiry Checker
- [x] Build Script 4 — Master Report Generator (all-in-one)
- [x] Install PS2EXE module
- [x] Compile PowerShell scripts into standalone .exe file
- [x] Verify EXE runs and generates all reports with one click
- [x] Export reports to timestamped CSV and TXT files

---

## 📖 Why PowerShell Automation?

In real IT environments, administrators deal with repetitive tasks daily — creating users, generating reports, checking password expiry, and auditing accounts. Doing these manually is:

- ❌ **Time consuming** — 50 users takes hours manually vs seconds with a script
- ❌ **Error prone** — typos happen when typing manually
- ❌ **Not scalable** — impossible to manage 1000 users manually
- ❌ **Not auditable** — manual processes are hard to document

PowerShell automation solves all of these:

- ✅ **Fast** — create 100 users in seconds from a CSV
- ✅ **Consistent** — same result every time, no typos
- ✅ **Scalable** — works for 10 or 10,000 users
- ✅ **Auditable** — scripts are documented and version controlled

---

## 📋 Script 1 — Bulk AD User Creation from CSV

### CSV Template (NewUsers.csv)

```csv
FirstName,LastName,Username,Department,Password,JobTitle
John,Smith,jsmith,IT,Welcome123!,IT Support Specialist
Sarah,Johnson,sjohnson,HR,Welcome123!,HR Manager
Michael,Brown,mbrown,Finance,Welcome123!,Financial Analyst
Emily,Davis,edavis,Management,Welcome123!,Operations Manager
David,Wilson,dwilson,IT,Welcome123!,Network Engineer
Jessica,Martinez,jmartinez,HR,Welcome123!,HR Coordinator
Robert,Anderson,randerson,Finance,Welcome123!,Senior Accountant
Lisa,Taylor,ltaylor,Management,Welcome123!,Project Manager
```

### Script

```powershell
$users = Import-Csv "C:\Users\Administrator\Desktop\NewUsers.csv"
$created = 0
$failed = 0

foreach ($user in $users) {
    try {
        New-ADUser `
          -GivenName $user.FirstName `
          -Surname $user.LastName `
          -Name "$($user.FirstName) $($user.LastName)" `
          -SamAccountName $user.Username `
          -UserPrincipalName "$($user.Username)@mylab.local" `
          -Department $user.Department `
          -Title $user.JobTitle `
          -AccountPassword (ConvertTo-SecureString $user.Password -AsPlainText -Force) `
          -Enabled $true `
          -Path "OU=VMUsers,DC=mylab,DC=local"

        Write-Host "✅ Created: $($user.FirstName) $($user.LastName) ($($user.Username))" -ForegroundColor Green
        $created++
    } catch {
        Write-Host "❌ Failed: $($user.Username) — $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

Write-Host "Total Created: $created | Total Failed: $failed"
```

### Result ✅
```
✅ Created: John Smith (jsmith)
✅ Created: Sarah Johnson (sjohnson)
✅ Created: Michael Brown (mbrown)
✅ Created: Emily Davis (edavis)
✅ Created: David Wilson (dwilson)
✅ Created: Jessica Martinez (jmartinez)
✅ Created: Robert Anderson (randerson)
✅ Created: Lisa Taylor (ltaylor)

Total Created: 8 | Total Failed: 0
```

---

## 📊 Script 2 — AD User Report Generator

Generates a full report of all AD users with their details and exports to CSV.

```powershell
$report = Get-ADUser -Filter * -SearchBase "OU=VMUsers,DC=mylab,DC=local" `
  -Properties DisplayName, SamAccountName, Department, Title, `
              Enabled, LastLogonDate, PasswordLastSet, PasswordExpired |
  Select-Object `
    @{N="Full Name";    E={$_.DisplayName}}, `
    @{N="Username";     E={$_.SamAccountName}}, `
    @{N="Department";   E={$_.Department}}, `
    @{N="Job Title";    E={$_.Title}}, `
    @{N="Enabled";      E={$_.Enabled}}, `
    @{N="Last Logon";   E={$_.LastLogonDate}}, `
    @{N="Password Set"; E={$_.PasswordLastSet}}, `
    @{N="Pwd Expired";  E={$_.PasswordExpired}}

$report | Format-Table -AutoSize
$report | Export-Csv -Path "C:\Users\Administrator\Desktop\ADUserReport.csv" -NoTypeInformation
```

### Result ✅
- **13 users** exported successfully
- All departments, job titles, last logon dates captured
- Saved as **ADUserReport.csv**

---

## 🔐 Script 3 — Password Expiry Checker

Checks all user passwords and flags accounts that are expired or expiring within 14 days.

```powershell
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
$warningDays = 14
$today = Get-Date

$results = Get-ADUser -Filter * -SearchBase "OU=VMUsers,DC=mylab,DC=local" `
  -Properties PasswordLastSet, PasswordNeverExpires, Enabled |
  Where-Object { $_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false } |
  ForEach-Object {
    $expiryDate = $_.PasswordLastSet.AddDays($maxPasswordAge)
    $daysLeft = ($expiryDate - $today).Days
    [PSCustomObject]@{
      Username   = $_.SamAccountName
      ExpiryDate = $expiryDate.ToString("MM/dd/yyyy")
      DaysLeft   = $daysLeft
      Status     = if ($daysLeft -lt 0) { "EXPIRED" }
                   elseif ($daysLeft -le $warningDays) { "EXPIRING SOON" }
                   else { "OK" }
    }
  } | Sort-Object DaysLeft

$results | Export-Csv "C:\Users\Administrator\Desktop\PasswordExpiryReport.csv" -NoTypeInformation
```

### Status Color Codes
| Status | Color | Meaning |
|---|---|---|
| **OK** | 🟢 Green | Password valid, no action needed |
| **EXPIRING SOON** | 🟡 Yellow | Expires within 14 days — notify user |
| **EXPIRED** | 🔴 Red | Password expired — immediate action needed |

### Result ✅
```
nsimon    | Expires: 06/13/2026 | Days Left:  39 | OK
finuser01 | Expires: 06/13/2026 | Days Left:  40 | OK
mgmtuser01| Expires: 06/13/2026 | Days Left:  40 | OK
ituser01  | Expires: 06/13/2026 | Days Left:  40 | OK
hruser01  | Expires: 06/13/2026 | Days Left:  40 | OK
jmartinez | Expires: 06/15/2026 | Days Left:  41 | OK
...
```

---

## 🚀 Script 4 — Master Report Generator EXE

Combines all three reports into a single script that generates everything with one click.

### What it generates

| Report | Format | Contents |
|---|---|---|
| **ADUserReport_[timestamp].csv** | CSV | All AD users with full details |
| **PasswordExpiryReport_[timestamp].csv** | CSV | Password status for all users |
| **SecuritySummary_[timestamp].txt** | TXT | Locked accounts, disabled users, expiring passwords |

### Security Summary Output
```
============================================
SECURITY SUMMARY REPORT
Generated: 05/04/2026 02:08:17
Domain: mylab.local
============================================

TOTAL USERS IN VMUSERS OU : 13
LOCKED OUT ACCOUNTS       : 0
DISABLED ACCOUNTS         : 0
EXPIRED PASSWORDS         : 0
PASSWORDS EXPIRING SOON   : 0

============================================
```

---

## 🔨 EXE Compilation with PS2EXE

### Install PS2EXE

```powershell
Install-Module -Name PS2EXE -Scope CurrentUser -Force
Get-Module -Name PS2EXE -ListAvailable
```

### Compile to EXE

```powershell
Import-Module PS2EXE

Invoke-PS2EXE `
  -InputFile "C:\Users\Administrator\Desktop\ITAdminReports.ps1" `
  -OutputFile "C:\Users\Administrator\Desktop\ITAdminReports.exe" `
  -RequireAdmin `
  -Title "IT Admin Report Generator" `
  -Description "Generates AD User, Password Expiry and Security Reports" `
  -Company "mylab.local" `
  -Version "1.0.1"
```

### EXE Properties

| Property | Value |
|---|---|
| **File Name** | ITAdminReports.exe |
| **Version** | 1.0.1 |
| **Requires Admin** | Yes |
| **Title** | IT Admin Report Generator |
| **Company** | mylab.local |

### How it works
1. Double-click **ITAdminReports.exe**
2. UAC prompt appears — click Yes (requires admin)
3. Console opens and runs all 3 reports automatically
4. Reports saved to **Desktop\Reports** folder with timestamp
5. Press any key to exit

---

## 📁 Output Files

All reports are saved to `C:\Users\Administrator\Desktop\Reports\` with timestamps:

```
Desktop\
└── Reports\
    ├── ADUserReport_2026-05-04_02-08.csv
    ├── PasswordExpiryReport_2026-05-04_02-08.csv
    └── SecuritySummary_2026-05-04_02-08.txt
```

---

## 📚 Skills Demonstrated

- PowerShell Scripting & Automation
- Active Directory Bulk User Management
- CSV Import/Export for IT Operations
- Password Policy Reporting & Monitoring
- Security Summary Report Generation
- PS2EXE — PowerShell to EXE Compilation
- Error Handling in PowerShell (try/catch)
- Timestamped File Output
- IT Tool Development for Non-Technical Staff
- Active Directory Module Commands

---

## 🚀 Future Improvements

- [ ] Add email notification for expiring passwords
- [ ] Add inactive user detection (no login in 30+ days)
- [ ] Auto-disable inactive accounts
- [ ] Add GUI interface using Windows Forms
- [ ] Schedule EXE to run automatically via Task Scheduler
- [ ] Add HTML report output for better readability
- [ ] Export reports directly to SharePoint

---

## 🔗 Related Projects

| Project | Description |
|---|---|
| **Part 1** | [DNS & DHCP Server Setup](https://yuzuki007.github.io/DNS-DHCP-Server-Setup-on-Windows-Server-2022/) |
| **Part 2** | [Windows 10 Client & AD Integration](https://yuzuki007.github.io/Windows-10-Client-Setup-Active-Directory-Integration/) |
| **Part 3** | Group Policy Objects (GPOs) |
| **Part 4** | File Server with NTFS Permissions |
| **Part 5** | Account Lockout & Audit Policy |
| **Part 6** | PowerShell Automation Scripts ← You are here |

---

## 👤 Author

**Neil Marvin Simon**
Home Lab Project — Part 6 | Windows Server 2022 | PowerShell Automation | PS2EXE
*Built for IT Portfolio purposes*
