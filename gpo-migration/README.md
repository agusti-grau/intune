# GPO → Intune Migration Tool

Analyzes Group Policy Object backups, builds a migration plan, creates the equivalent Intune policies via Microsoft Graph API, and generates an HTML report.

---

## Requirements

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or later |
| OS | Windows (run from a domain-joined machine or any machine with network access to the DC) |
| Permissions | **Domain**: Group Policy read access · **Intune**: Global Admin or Intune Administrator + Cloud Device Administrator |
| Module | `Microsoft.Graph.Authentication` — auto-installed on first run if missing |

---

## Step 1 — Export GPO Backups from Active Directory

Run the following on a **domain controller** or any machine with the **Group Policy Management** feature/RSAT installed.

### Option A — Export all GPOs (recommended)

```powershell
# Creates one subfolder per GPO (GUID-named) inside C:\GPOBackups
Backup-GPO -All -Path "C:\GPOBackups"
```

### Option B — Export specific GPOs by name

```powershell
Backup-GPO -Name "Default Domain Policy"      -Path "C:\GPOBackups"
Backup-GPO -Name "Workstation Security Policy" -Path "C:\GPOBackups"
Backup-GPO -Name "Windows Firewall Policy"     -Path "C:\GPOBackups"
```

### Option C — Export all GPOs with a comment (useful for tracking)

```powershell
$date = Get-Date -Format "yyyy-MM-dd"
Backup-GPO -All -Path "C:\GPOBackups" -Comment "Pre-Intune migration backup $date"
```

> **Note:** `Backup-GPO` requires the `GroupPolicy` PowerShell module.
> On a domain controller it is available by default.
> On a workstation, install it via:
> ```powershell
> Add-WindowsCapability -Online -Name "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0"
> ```

### What the backup folder looks like

```
C:\GPOBackups\
├── {4A8B3C1D-...}\          ← each GPO gets a GUID-named folder
│   ├── bkupInfo.xml         ← metadata (GPO name, domain, date)
│   ├── gpreport.xml         ← full settings report  ← parsed by this script
│   ├── Backup.xml
│   └── manifest.xml
├── {7F2E9A0B-...}\
│   └── ...
└── manifest.xml
```

The migration script reads `gpreport.xml` from each subfolder. All subfolders without that file are silently skipped.

---

## Step 2 — Copy the Backup Folder to the Target Machine

Transfer `C:\GPOBackups` to the machine you will run the migration from (can be the same DC, a jump host, or a technician laptop with internet access to Microsoft Graph).

---

## Step 3 — Run the Migration Script

### Dry run first — see the plan without making any changes

```powershell
.\Invoke-GPOtoIntuneMigration.ps1 -GPOBackupPath "C:\GPOBackups" -PlanOnly
```

### Full migration

```powershell
.\Invoke-GPOtoIntuneMigration.ps1 -GPOBackupPath "C:\GPOBackups"
```

The script will:
1. Parse all GPO backups
2. Display the full migration plan with actions (`CREATE` / `PARTIAL` / `MANUAL`)
3. Ask for confirmation before making any changes
4. Authenticate interactively to the customer's Azure AD tenant
5. Create all automatable policies in Intune
6. Open an HTML report in your browser

### All parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-GPOBackupPath` | Yes | — | Path to the GPO backup root folder |
| `-TenantId` | No | Auto-detected | Azure AD tenant ID |
| `-PlanOnly` | No | `$false` | Show plan, create nothing |
| `-PolicyPrefix` | No | `"MIGRATED - "` | Prefix added to all policy names |
| `-ReportPath` | No | Current dir | Output directory for the HTML report |
| `-SkipModuleInstall` | No | `$false` | Skip automatic module installation |

### Example — customer-specific run

```powershell
.\Invoke-GPOtoIntuneMigration.ps1 `
    -GPOBackupPath "C:\GPOBackups" `
    -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -PolicyPrefix  "CONTOSO - " `
    -ReportPath    "C:\Reports\Contoso"
```

---

## What Gets Migrated

| GPO Setting | Intune Policy Type | Result |
|---|---|---|
| Password Policy | Compliance Policy | AUTO |
| Account Lockout Policy | Settings Catalog | AUTO |
| Audit Policy | Settings Catalog | AUTO |
| Windows Firewall Profiles | Endpoint Security – Firewall | AUTO |
| Firewall Rules | Custom Configuration (OMA-URI) | AUTO |
| Administrative Templates (ADMX) | Group Policy Configurations | AUTO* |
| Security Options | Report only | PARTIAL |
| User Rights Assignment | Report only | PARTIAL |
| BitLocker | Report only | PARTIAL |
| Startup / Logon Scripts | Report only | MANUAL |

\* ADMX settings are matched by display name against your tenant's policy definitions. Settings with no matching definition in Intune are flagged in the report.

---

## After the Migration

1. **Open the HTML report** — review all `PARTIAL` and `MANUAL` items.
2. **Assign policies** — newly created policies are unassigned. Go to Intune → the relevant policy → **Properties → Assignments** and assign to the appropriate groups.
3. **Security Options / User Rights** — the report lists the raw registry keys and values. Create a **Custom Configuration Profile** in Intune and add the OMA-URI entries manually.
4. **Scripts** — review each script listed in the report for security, then upload via **Intune → Devices → Scripts**.
5. **BitLocker** — create a **Disk Encryption** policy under Endpoint Security with the settings noted in the report.
6. **Test on a pilot group** before rolling out to all devices.

---

## Troubleshooting

**`Backup-GPO` not found**
Install RSAT Group Policy tools (see Option A note above).

**`gpreport.xml` missing from backup folders**
Re-run `Backup-GPO`. Incomplete backups (e.g. interrupted copy) may lack this file.

**Graph authentication fails**
Ensure the account used has `DeviceManagementConfiguration.ReadWrite.All` and `DeviceManagementServiceConfig.ReadWrite.All` permissions, and that MFA is satisfied.

**ADMX settings show 0 matched / all missing**
The GPO uses custom (third-party) ADMX templates not ingested into your tenant. Ingest the ADMX files first via **Intune → Devices → Configuration → Import ADMX**, then re-run.

**`Install-Module` is blocked by execution policy**
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```
