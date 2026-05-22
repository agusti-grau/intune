#Requires -Version 5.1
<#
.SYNOPSIS
    Migrates Group Policy Objects (GPOs) to Microsoft Intune via Microsoft Graph API.

.DESCRIPTION
    1. Discovers GPO backup folders (GUID-named, each containing gpreport.xml)
    2. Parses and maps settings to Intune policy equivalents
    3. Displays the full migration plan
    4. On confirmation, connects to Graph and creates all policies
    5. Generates an HTML report

    Supported migrations:
      - Password Policy          → Compliance Policy
      - Account Lockout          → Settings Catalog
      - Audit Policy             → Settings Catalog
      - Windows Firewall         → Endpoint Security Firewall Policy
      - Administrative Templates → Group Policy Configurations (ADMX)
      - Startup/Logon Scripts    → Device Management Scripts (flagged for review)
      - Security Options         → Custom OMA-URI (flagged for review)
      - User Rights Assignment   → Custom OMA-URI (flagged for review)
      - BitLocker                → Disk Encryption (Endpoint Security)

.PARAMETER GPOBackupPath
    Path to the root folder containing GUID-named GPO backup subfolders.
    Each subfolder must contain a gpreport.xml file.

.PARAMETER TenantId
    Azure AD Tenant ID. Discovered automatically from the connected account if omitted.

.PARAMETER PlanOnly
    Show the migration plan without creating any Intune policies.

.PARAMETER ReportPath
    Directory where the HTML report will be saved. Defaults to current directory.

.PARAMETER PolicyPrefix
    String prepended to every created policy name. Default: "MIGRATED - ".

.PARAMETER SkipModuleInstall
    Skip automatic installation of the Microsoft.Graph.Authentication module.

.EXAMPLE
    .\Invoke-GPOtoIntuneMigration.ps1 -GPOBackupPath "C:\GPOBackups" -PlanOnly

.EXAMPLE
    .\Invoke-GPOtoIntuneMigration.ps1 -GPOBackupPath "C:\GPOBackups" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    .\Invoke-GPOtoIntuneMigration.ps1 -GPOBackupPath "C:\GPOBackups" -PolicyPrefix "CONTOSO - " -ReportPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to GPO backup root folder")]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$GPOBackupPath,

    [string]$TenantId,

    [switch]$PlanOnly,

    [string]$ReportPath = (Get-Location).Path,

    [string]$PolicyPrefix = "MIGRATED - ",

    [switch]$SkipModuleInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:Log             = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:MigrationResults = [System.Collections.Generic.List[PSCustomObject]]::new()

$GRAPH_BASE = "https://graph.microsoft.com/beta"

# ──────────────────────────────────────────────────────────────
# CONSOLE HELPERS
# ──────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red" }
        default   { "Cyan" }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
    $script:Log.Add([PSCustomObject]@{ Time = $ts; Level = $Level; Message = $Message })
}

function Write-Section {
    param([string]$Title)
    $line = "─" * 68
    Write-Host "`n$line" -ForegroundColor DarkGray
    Write-Host "  $Title" -ForegroundColor White
    Write-Host "$line" -ForegroundColor DarkGray
}

# ──────────────────────────────────────────────────────────────
# MODULE & GRAPH AUTH
# ──────────────────────────────────────────────────────────────

function Initialize-GraphConnection {
    Write-Section "Connecting to Microsoft Graph"

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        if ($SkipModuleInstall) {
            Write-Log "Microsoft.Graph.Authentication not found. Install it or remove -SkipModuleInstall." "ERROR"
            throw "Module not available"
        }
        Write-Log "Installing Microsoft.Graph.Authentication (CurrentUser)..." "WARN"
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $scopes = @(
        "DeviceManagementConfiguration.ReadWrite.All",
        "DeviceManagementApps.ReadWrite.All",
        "DeviceManagementServiceConfig.ReadWrite.All",
        "DeviceManagementManagedDevices.ReadWrite.All"
    )

    $params = @{ Scopes = $scopes }
    if ($TenantId) { $params.TenantId = $TenantId }

    Connect-MgGraph @params -ErrorAction Stop
    $ctx = Get-MgContext
    Write-Log "Connected  Tenant: $($ctx.TenantId)  Account: $($ctx.Account)" "SUCCESS"
}

# ──────────────────────────────────────────────────────────────
# GRAPH API WRAPPERS
# ──────────────────────────────────────────────────────────────

function Invoke-GPost {
    param([string]$Uri, [hashtable]$Body)
    $json = $Body | ConvertTo-Json -Depth 20 -Compress
    return Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $json -ContentType "application/json" -ErrorAction Stop
}

function Invoke-GGet {
    param([string]$Uri)
    return Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
}

# ──────────────────────────────────────────────────────────────
# GPO BACKUP DISCOVERY
# ──────────────────────────────────────────────────────────────

function Get-GPOBackupList {
    param([string]$Path)

    $list = @()
    foreach ($dir in Get-ChildItem -Path $Path -Directory) {
        $report  = Join-Path $dir.FullName "gpreport.xml"
        $bkupInfo = Join-Path $dir.FullName "bkupInfo.xml"
        if (Test-Path $report) {
            $list += [PSCustomObject]@{
                FolderName   = $dir.Name
                Path         = $dir.FullName
                ReportFile   = $report
                BkupInfoFile = if (Test-Path $bkupInfo) { $bkupInfo } else { $null }
            }
        }
    }

    Write-Log "Found $($list.Count) GPO backup(s) in '$Path'"
    return $list
}

# ──────────────────────────────────────────────────────────────
# XML PARSING HELPERS
# ──────────────────────────────────────────────────────────────

function Get-XmlValue {
    param($Node, [string]$LocalName, [string]$Default = $null)
    if ($null -eq $Node) { return $Default }
    $child = $Node.SelectSingleNode(".//*[local-name()='$LocalName']")
    if ($child) { return $child.InnerText }
    # Try direct attribute
    if ($Node.$LocalName) { return $Node.$LocalName }
    return $Default
}

function Get-IntOrNull {
    param([string]$Val)
    $n = 0
    if ([int]::TryParse($Val, [ref]$n)) { return $n }
    return $null
}

# ──────────────────────────────────────────────────────────────
# GPO REPORT PARSER
# ──────────────────────────────────────────────────────────────

function Read-GPOReport {
    param([PSCustomObject]$Backup)

    try {
        [xml]$xml = Get-Content $Backup.ReportFile -Encoding UTF8 -Raw
    } catch {
        Write-Log "Cannot parse '$($Backup.ReportFile)': $_" "ERROR"
        return $null
    }

    # Resolve GPO display name from multiple sources
    $root = $xml.DocumentElement
    $gpoName = $root.Name
    if (-not $gpoName -or $gpoName -eq "GPO") {
        $nameNode = $root.SelectSingleNode(".//*[local-name()='Name'][not(*)]")
        if ($nameNode) { $gpoName = $nameNode.InnerText }
    }
    if (-not $gpoName -and $Backup.BkupInfoFile) {
        try {
            [xml]$bkup = Get-Content $Backup.BkupInfoFile -Encoding UTF8 -Raw
            $nameNode = $bkup.SelectSingleNode("//*[local-name()='GPODisplayName']")
            if ($nameNode) { $gpoName = $nameNode.InnerText }
        } catch { }
    }
    if (-not $gpoName) { $gpoName = $Backup.FolderName }

    $report = [PSCustomObject]@{
        Name             = $gpoName
        Guid             = $Backup.FolderName
        Path             = $Backup.Path
        PasswordPolicy   = $null
        LockoutPolicy    = $null
        AuditSettings    = [System.Collections.Generic.List[PSCustomObject]]::new()
        UserRights        = [System.Collections.Generic.List[PSCustomObject]]::new()
        SecurityOptions  = [System.Collections.Generic.List[PSCustomObject]]::new()
        FirewallProfiles = $null
        FirewallRules    = [System.Collections.Generic.List[PSCustomObject]]::new()
        RegistrySettings = [System.Collections.Generic.List[PSCustomObject]]::new()
        BitLockerSettings = $null
        Scripts          = [System.Collections.Generic.List[PSCustomObject]]::new()
    }

    # All Extension nodes (Computer + User)
    $extensions = $root.SelectNodes(".//*[local-name()='Extension']")
    foreach ($ext in $extensions) {
        $outerXml = $ext.OuterXml
        $isUser   = Is-UserExtension -ExtNode $ext

        if ($outerXml -match "SystemAccess|UserRightsAssignment|EventAudit|AuditLog|SecurityOptions|RegistryValues") {
            Parse-SecurityExtension -Ext $ext -Report $report
        }
        if ($outerXml -match "DomainProfile|PrivateProfile|PublicProfile|EnableFirewall|FirewallRule|WindowsFirewall") {
            Parse-FirewallExtension -Ext $ext -Report $report
        }
        if ($outerXml -match "<q\d*:Policy>|RegistrySettings|NtUserSetting|<Policy>") {
            Parse-RegistryExtension -Ext $ext -Report $report -IsUser $isUser
        }
        if ($outerXml -match "Script[^s]|ScriptsSettings|<Script>") {
            Parse-ScriptsExtension -Ext $ext -Report $report -IsUser $isUser
        }
        if ($outerXml -match "BitLocker|BDE") {
            Parse-BitLockerExtension -Ext $ext -Report $report
        }
    }

    return $report
}

function Is-UserExtension {
    param($ExtNode)
    $parent = $ExtNode.ParentNode
    while ($null -ne $parent) {
        if ($parent.LocalName -eq "User") { return $true }
        if ($parent.LocalName -eq "Computer") { return $false }
        $parent = $parent.ParentNode
    }
    return $false
}

# ── Security Extension ──────────────────────────────────────

function Parse-SecurityExtension {
    param($Ext, $Report)

    # ── SystemAccess: password + lockout ──
    $sa = $Ext.SelectSingleNode(".//*[local-name()='SystemAccess']")
    if ($sa) {
        $Report.PasswordPolicy = [PSCustomObject]@{
            MinPasswordAge     = (Get-IntOrNull (Get-XmlValue $sa "MinimumPasswordAge" "0"))
            MaxPasswordAge     = (Get-IntOrNull (Get-XmlValue $sa "MaximumPasswordAge" "0"))
            MinPasswordLength  = (Get-IntOrNull (Get-XmlValue $sa "MinimumPasswordLength" "0"))
            PasswordComplexity = (Get-IntOrNull (Get-XmlValue $sa "PasswordComplexity" "0"))
            PasswordHistory    = (Get-IntOrNull (Get-XmlValue $sa "PasswordHistorySize" "0"))
            ClearTextPassword  = (Get-IntOrNull (Get-XmlValue $sa "ClearTextPassword" "0"))
            LockoutThreshold   = (Get-IntOrNull (Get-XmlValue $sa "LockoutBadCount" "0"))
            LockoutDuration    = (Get-IntOrNull (Get-XmlValue $sa "LockoutDuration" "0"))
            ResetLockoutCount  = (Get-IntOrNull (Get-XmlValue $sa "ResetLockoutCount" "0"))
        }
    }

    # Alternative format: <Account><Name>...</Name><SettingNumber>...</SettingNumber></Account>
    $accountNodes = $Ext.SelectNodes(".//*[local-name()='Account']")
    foreach ($acc in $accountNodes) {
        $name = Get-XmlValue $acc "Name"
        $val  = Get-IntOrNull (Get-XmlValue $acc "SettingNumber" "0")
        if (-not $name -or $null -eq $val) { continue }

        if ($null -eq $Report.PasswordPolicy) {
            $Report.PasswordPolicy = [PSCustomObject]@{
                MinPasswordAge=0; MaxPasswordAge=0; MinPasswordLength=0
                PasswordComplexity=0; PasswordHistory=0; ClearTextPassword=0
                LockoutThreshold=0; LockoutDuration=0; ResetLockoutCount=0
            }
        }
        switch ($name) {
            "MinimumPasswordAge"    { $Report.PasswordPolicy.MinPasswordAge     = $val }
            "MaximumPasswordAge"    { $Report.PasswordPolicy.MaxPasswordAge     = $val }
            "MinimumPasswordLength" { $Report.PasswordPolicy.MinPasswordLength  = $val }
            "PasswordComplexity"    { $Report.PasswordPolicy.PasswordComplexity = $val }
            "PasswordHistorySize"   { $Report.PasswordPolicy.PasswordHistory    = $val }
            "LockoutBadCount"       { $Report.PasswordPolicy.LockoutThreshold   = $val }
            "LockoutDuration"       { $Report.PasswordPolicy.LockoutDuration    = $val }
            "ResetLockoutCount"     { $Report.PasswordPolicy.ResetLockoutCount  = $val }
        }
    }

    # ── EventAudit ──
    $eventAudit = $Ext.SelectSingleNode(".//*[local-name()='EventAudit']")
    if ($eventAudit) {
        # 0=None, 1=Success, 2=Failure, 3=Both
        $auditFields = @(
            @{ Name="AccountLogon";         Xml="AuditLogonEvents" }
            @{ Name="AccountManagement";    Xml="AuditAccountManage" }
            @{ Name="DSAccess";             Xml="AuditDSAccess" }
            @{ Name="LogonLogoff";          Xml="AuditLogonEvents" }
            @{ Name="ObjectAccess";         Xml="AuditObjectAccess" }
            @{ Name="PolicyChange";         Xml="AuditPolicyChange" }
            @{ Name="PrivilegeUse";         Xml="AuditPrivilegeUse" }
            @{ Name="ProcessTracking";      Xml="AuditProcessTracking" }
            @{ Name="SystemEvents";         Xml="AuditSystemEvents" }
        )
        foreach ($f in $auditFields) {
            $raw = Get-XmlValue $eventAudit $f.Xml
            if ($null -ne $raw) {
                $v = Get-IntOrNull $raw
                if ($null -ne $v) {
                    $Report.AuditSettings.Add([PSCustomObject]@{ Category=$f.Name; Value=$v }) | Out-Null
                }
            }
        }
    }

    # ── User Rights Assignment ──
    $urNodes = $Ext.SelectNodes(".//*[local-name()='UserRightsAssignment']")
    foreach ($ur in $urNodes) {
        $right = Get-XmlValue $ur "Name"
        if (-not $right) { continue }
        $members = $ur.SelectNodes(".//*[local-name()='Name'][not(*)]") |
                   ForEach-Object { $_.InnerText } |
                   Select-Object -Unique
        $Report.UserRights.Add([PSCustomObject]@{
            Right   = $right
            Members = $members -join "; "
        }) | Out-Null
    }

    # ── Security Options ──
    $soNodes = $Ext.SelectNodes(".//*[local-name()='SecurityOptions']")
    foreach ($so in $soNodes) {
        $key = Get-XmlValue $so "KeyName"
        if (-not $key) { continue }
        $Report.SecurityOptions.Add([PSCustomObject]@{
            KeyName = $key
            Value   = (Get-XmlValue $so "SettingNumber") ?? (Get-XmlValue $so "SettingStrings")
        }) | Out-Null
    }

    # ── RegistryValues (security-related) ──
    $rvNodes = $Ext.SelectNodes(".//*[local-name()='RegistryValues']//*[local-name()='RegistryValue']")
    foreach ($rv in $rvNodes) {
        $path = Get-XmlValue $rv "Name"
        $val  = Get-XmlValue $rv "Value"
        if (-not $path) { continue }
        $Report.SecurityOptions.Add([PSCustomObject]@{
            KeyName = $path
            Value   = $val
        }) | Out-Null
    }
}

# ── Firewall Extension ──────────────────────────────────────

function Parse-FirewallExtension {
    param($Ext, $Report)

    if ($null -eq $Report.FirewallProfiles) {
        $Report.FirewallProfiles = [PSCustomObject]@{
            Domain  = [PSCustomObject]@{ Enabled=$null; DefaultInbound=$null; DefaultOutbound=$null; Notifications=$null }
            Private = [PSCustomObject]@{ Enabled=$null; DefaultInbound=$null; DefaultOutbound=$null; Notifications=$null }
            Public  = [PSCustomObject]@{ Enabled=$null; DefaultInbound=$null; DefaultOutbound=$null; Notifications=$null }
        }
    }

    foreach ($profileName in @("Domain","Private","Public")) {
        $pNode = $Ext.SelectSingleNode(".//*[local-name()='${profileName}Profile']")
        if (-not $pNode) {
            $pNode = $Ext.SelectSingleNode(".//*[local-name()='Profile'][@Name='$profileName']")
        }
        if ($pNode) {
            $p = $Report.FirewallProfiles.$profileName
            $p.Enabled         = Get-XmlValue $pNode "EnableFirewall"
            $p.DefaultInbound  = Get-XmlValue $pNode "DefaultInboundAction"
            $p.DefaultOutbound = Get-XmlValue $pNode "DefaultOutboundAction"
            $p.Notifications   = Get-XmlValue $pNode "DisableNotifications"
        }
    }

    # Firewall rules
    $ruleNodes = $Ext.SelectNodes(".//*[local-name()='FirewallRule' or local-name()='Rule']")
    foreach ($rule in $ruleNodes) {
        $ruleName = Get-XmlValue $rule "Name"
        if (-not $ruleName) { $ruleName = Get-XmlValue $rule "Id" }
        $Report.FirewallRules.Add([PSCustomObject]@{
            Name       = $ruleName
            Direction  = Get-XmlValue $rule "Dir"
            Action     = Get-XmlValue $rule "Action"
            Protocol   = Get-XmlValue $rule "Protocol"
            LocalPort  = Get-XmlValue $rule "LPort"
            RemotePort = Get-XmlValue $rule "RPort"
            Enabled    = Get-XmlValue $rule "Active"
            Profile    = Get-XmlValue $rule "Profile"
            App        = Get-XmlValue $rule "App"
        }) | Out-Null
    }
}

# ── Registry/ADMX Extension ────────────────────────────────

function Parse-RegistryExtension {
    param($Ext, $Report, [bool]$IsUser)

    $policyNodes = $Ext.SelectNodes(".//*[local-name()='Policy']")
    foreach ($p in $policyNodes) {
        $settingName = Get-XmlValue $p "Name"
        if (-not $settingName) { continue }
        $Report.RegistrySettings.Add([PSCustomObject]@{
            Name     = $settingName
            State    = (Get-XmlValue $p "State" "Unknown")
            Category = (Get-XmlValue $p "Category")
            KeyName  = (Get-XmlValue $p "KeyName")
            ValueName = (Get-XmlValue $p "ValueName")
            IsUser   = $IsUser
        }) | Out-Null
    }
}

# ── Scripts Extension ───────────────────────────────────────

function Parse-ScriptsExtension {
    param($Ext, $Report, [bool]$IsUser)

    $scriptNodes = $Ext.SelectNodes(".//*[local-name()='Script']")
    foreach ($s in $scriptNodes) {
        $cmd = Get-XmlValue $s "Command"
        if (-not $cmd) { continue }
        $Report.Scripts.Add([PSCustomObject]@{
            Command    = $cmd
            Parameters = (Get-XmlValue $s "Parameters")
            Type       = if ($IsUser) { "User" } else { "Computer" }
            RunOrder   = (Get-XmlValue $s "RunOrder" "Startup")
        }) | Out-Null
    }
}

# ── BitLocker Extension ─────────────────────────────────────

function Parse-BitLockerExtension {
    param($Ext, $Report)

    # Detect BitLocker settings presence (detailed mapping handled separately)
    $requireEncryption   = Get-XmlValue $Ext "RequireDeviceEncryption"
    $encryptionMethod    = Get-XmlValue $Ext "EncryptionMethodByDriveType"
    $recoveryOptions     = Get-XmlValue $Ext "RecoveryOptions"

    if ($requireEncryption -or $encryptionMethod -or $recoveryOptions) {
        $Report.BitLockerSettings = [PSCustomObject]@{
            RequireEncryption = $requireEncryption
            EncryptionMethod  = $encryptionMethod
            RecoveryOptions   = $recoveryOptions
            RawXml            = $Ext.OuterXml
        }
    }
}

# ──────────────────────────────────────────────────────────────
# MIGRATION PLAN BUILDER
# ──────────────────────────────────────────────────────────────

function Build-MigrationPlan {
    param([array]$Reports)

    $plan = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($gpo in $Reports) {
        $policies = [System.Collections.Generic.List[PSCustomObject]]::new()

        # ── Password / Compliance ──────────────────────────
        $pp = $gpo.PasswordPolicy
        if ($pp) {
            $hasPasswordSetting = ($pp.MinPasswordLength -gt 0) -or ($pp.MaxPasswordAge -gt 0) -or
                                  ($pp.PasswordComplexity -gt 0) -or ($pp.PasswordHistory -gt 0)
            $hasLockoutSetting  = ($pp.LockoutThreshold -gt 0)

            if ($hasPasswordSetting) {
                $policies.Add([PSCustomObject]@{
                    PolicyType  = "Compliance Policy (Password)"
                    PolicyName  = "$PolicyPrefix$($gpo.Name) - Password & Compliance"
                    Description = "Password requirements migrated from GPO '$($gpo.Name)'"
                    Settings    = $pp
                    Action      = "CREATE"
                    ApiType     = "compliance"
                    Details     = @(
                        if ($pp.MinPasswordLength -gt 0) { "Min length: $($pp.MinPasswordLength)" }
                        if ($pp.MaxPasswordAge -gt 0)    { "Max age: $($pp.MaxPasswordAge) days" }
                        if ($pp.PasswordComplexity -eq 1) { "Complexity: Enabled" }
                        if ($pp.PasswordHistory -gt 0)   { "History: $($pp.PasswordHistory)" }
                    )
                }) | Out-Null
            }

            if ($hasLockoutSetting) {
                $policies.Add([PSCustomObject]@{
                    PolicyType  = "Settings Catalog (Account Lockout)"
                    PolicyName  = "$PolicyPrefix$($gpo.Name) - Account Lockout"
                    Description = "Account lockout policy migrated from GPO '$($gpo.Name)'"
                    Settings    = $pp
                    Action      = "CREATE"
                    ApiType     = "lockout"
                    Details     = @(
                        "Threshold: $($pp.LockoutThreshold)"
                        "Duration: $($pp.LockoutDuration) min"
                        "Reset after: $($pp.ResetLockoutCount) min"
                    )
                }) | Out-Null
            }
        }

        # ── Audit Policy ──────────────────────────────────
        if ($gpo.AuditSettings.Count -gt 0) {
            $auditEnabled = $gpo.AuditSettings | Where-Object { $_.Value -gt 0 }
            if ($auditEnabled) {
                $policies.Add([PSCustomObject]@{
                    PolicyType  = "Settings Catalog (Audit Policy)"
                    PolicyName  = "$PolicyPrefix$($gpo.Name) - Audit Policy"
                    Description = "$($auditEnabled.Count) audit categorie(s) from GPO '$($gpo.Name)'"
                    Settings    = $gpo.AuditSettings
                    Action      = "CREATE"
                    ApiType     = "audit"
                    Details     = $auditEnabled | ForEach-Object { "$($_.Category): $(Get-AuditLabel $_.Value)" }
                }) | Out-Null
            }
        }

        # ── Windows Firewall ──────────────────────────────
        if ($gpo.FirewallProfiles) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "Endpoint Security (Firewall Profiles)"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - Firewall"
                Description = "Firewall Domain/Private/Public profile settings from GPO '$($gpo.Name)'"
                Settings    = $gpo.FirewallProfiles
                Action      = "CREATE"
                ApiType     = "firewall"
                Details     = @(
                    "Domain:  Enabled=$($gpo.FirewallProfiles.Domain.Enabled ?? 'not set')"
                    "Private: Enabled=$($gpo.FirewallProfiles.Private.Enabled ?? 'not set')"
                    "Public:  Enabled=$($gpo.FirewallProfiles.Public.Enabled ?? 'not set')"
                )
            }) | Out-Null
        }

        if ($gpo.FirewallRules.Count -gt 0) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "Endpoint Security (Firewall Rules)"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - Firewall Rules"
                Description = "$($gpo.FirewallRules.Count) firewall rule(s) from GPO '$($gpo.Name)'"
                Settings    = $gpo.FirewallRules
                Action      = "CREATE"
                ApiType     = "firewallRules"
                Details     = $gpo.FirewallRules | Select-Object -First 5 |
                              ForEach-Object { "$($_.Name) [$($_.Direction)/$($_.Action)]" }
            }) | Out-Null
        }

        # ── Admin Templates / ADMX ────────────────────────
        if ($gpo.RegistrySettings.Count -gt 0) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "Administrative Templates (ADMX)"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - Admin Templates"
                Description = "$($gpo.RegistrySettings.Count) ADMX policy setting(s) from GPO '$($gpo.Name)'"
                Settings    = $gpo.RegistrySettings
                Action      = "CREATE"
                ApiType     = "admx"
                Details     = $gpo.RegistrySettings | Select-Object -First 5 |
                              ForEach-Object { "$($_.Name) [$($_.State)]" }
            }) | Out-Null
        }

        # ── BitLocker ─────────────────────────────────────
        if ($gpo.BitLockerSettings) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "Endpoint Security (Disk Encryption / BitLocker)"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - BitLocker"
                Description = "BitLocker settings from GPO '$($gpo.Name)'"
                Settings    = $gpo.BitLockerSettings
                Action      = "PARTIAL"
                ApiType     = "bitlocker"
                Details     = @("BitLocker settings detected — review generated policy before assigning")
            }) | Out-Null
        }

        # ── Security Options → OMA-URI ────────────────────
        if ($gpo.SecurityOptions.Count -gt 0) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "Custom Profile (Security Options via OMA-URI)"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - Security Options"
                Description = "$($gpo.SecurityOptions.Count) security option(s) from GPO '$($gpo.Name)'"
                Settings    = $gpo.SecurityOptions
                Action      = "PARTIAL"
                ApiType     = "securityOptions"
                Details     = @("Security options exported to report. Create Custom Profile manually or via OMA-URI.")
            }) | Out-Null
        }

        # ── User Rights → OMA-URI ─────────────────────────
        if ($gpo.UserRights.Count -gt 0) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "Custom Profile (User Rights via OMA-URI)"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - User Rights"
                Description = "$($gpo.UserRights.Count) user rights assignment(s) from GPO '$($gpo.Name)'"
                Settings    = $gpo.UserRights
                Action      = "PARTIAL"
                ApiType     = "userRights"
                Details     = $gpo.UserRights | ForEach-Object { "$($_.Right): $($_.Members)" }
            }) | Out-Null
        }

        # ── Scripts ───────────────────────────────────────
        foreach ($s in $gpo.Scripts) {
            $policies.Add([PSCustomObject]@{
                PolicyType  = "PowerShell Script"
                PolicyName  = "$PolicyPrefix$($gpo.Name) - Script ($($s.RunOrder) / $($s.Type))"
                Description = "Script: $($s.Command)"
                Settings    = $s
                Action      = "MANUAL"
                ApiType     = "script"
                Details     = @("Command: $($s.Command)","Params: $($s.Parameters)","Review before deploying!")
            }) | Out-Null
        }

        $plan.Add([PSCustomObject]@{
            GPOName  = $gpo.Name
            GPOGuid  = $gpo.Guid
            Policies = $policies
        }) | Out-Null
    }

    return $plan
}

function Get-AuditLabel {
    param([int]$Val)
    switch ($Val) {
        0 { return "No Auditing" }
        1 { return "Success" }
        2 { return "Failure" }
        3 { return "Success + Failure" }
        default { return "Unknown ($Val)" }
    }
}

# ──────────────────────────────────────────────────────────────
# MIGRATION PLAN DISPLAY
# ──────────────────────────────────────────────────────────────

function Show-MigrationPlan {
    param([array]$Plan)

    Write-Section "MIGRATION PLAN"

    $counts = @{ CREATE = 0; PARTIAL = 0; MANUAL = 0 }

    foreach ($item in $Plan) {
        Write-Host "`n  GPO: " -NoNewline -ForegroundColor DarkGray
        Write-Host $item.GPOName -ForegroundColor Yellow
        Write-Host "  GUID: $($item.GPOGuid)" -ForegroundColor DarkGray

        if ($item.Policies.Count -eq 0) {
            Write-Host "    (no supported settings found)" -ForegroundColor DarkGray
            continue
        }

        foreach ($pol in $item.Policies) {
            $counts[$pol.Action]++
            $actionColor = switch ($pol.Action) {
                "CREATE"  { "Green" }
                "PARTIAL" { "Yellow" }
                "MANUAL"  { "Magenta" }
                default   { "White" }
            }
            $label = "[$($pol.Action.PadRight(7))]"
            Write-Host "    $label " -NoNewline -ForegroundColor $actionColor
            Write-Host "$($pol.PolicyType)" -ForegroundColor Cyan
            Write-Host "             Name: $($pol.PolicyName)" -ForegroundColor White
            if ($pol.Details) {
                foreach ($d in $pol.Details | Select-Object -First 4) {
                    Write-Host "             · $d" -ForegroundColor DarkGray
                }
            }
        }
    }

    $totalCreate  = $counts.CREATE
    $totalPartial = $counts.PARTIAL
    $totalManual  = $counts.MANUAL
    $totalAll     = $totalCreate + $totalPartial + $totalManual

    Write-Host "`n$('─' * 68)" -ForegroundColor DarkGray
    Write-Host "  Total policies: $totalAll" -ForegroundColor White
    Write-Host "  CREATE  (auto)    : $totalCreate"  -ForegroundColor Green
    Write-Host "  PARTIAL (review)  : $totalPartial" -ForegroundColor Yellow
    Write-Host "  MANUAL  (skipped) : $totalManual"  -ForegroundColor Magenta
    Write-Host "`n  Legend:" -ForegroundColor DarkGray
    Write-Host "    CREATE  = Fully automated" -ForegroundColor Green
    Write-Host "    PARTIAL = Created with available settings; manual review recommended" -ForegroundColor Yellow
    Write-Host "    MANUAL  = Script/unsupported — exported to report only" -ForegroundColor Magenta
}

# ──────────────────────────────────────────────────────────────
# POLICY CREATION FUNCTIONS
# ──────────────────────────────────────────────────────────────

function New-CompliancePolicy {
    param([string]$Name, [string]$Desc, $S)

    $body = @{
        "@odata.type" = "#microsoft.graph.windows10CompliancePolicy"
        displayName   = $Name
        description   = $Desc
    }

    if ($S.MinPasswordLength  -gt 0) { $body.passwordRequired = $true; $body.passwordMinimumLength = $S.MinPasswordLength }
    if ($S.PasswordComplexity -eq 1) { $body.passwordRequiredType = "alphanumeric"; $body.passwordMinimumCharacterSetCount = 3 }
    if ($S.MaxPasswordAge     -gt 0) { $body.passwordExpirationDays = $S.MaxPasswordAge }
    if ($S.PasswordHistory    -gt 0) { $body.passwordPreviousPasswordBlockCount = $S.PasswordHistory }
    if ($S.ClearTextPassword  -eq 0) { $body.passwordBlockSimple = $true }

    return Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/deviceCompliancePolicies" -Body $body
}

function New-LockoutPolicy {
    param([string]$Name, [string]$Desc, $S)

    $settings = [System.Collections.Generic.List[hashtable]]::new()

    function Add-IntSetting {
        param([string]$DefId, [int]$Value)
        $settings.Add(@{
            "@odata.type"   = "#microsoft.graph.deviceManagementConfigurationSetting"
            settingInstance = @{
                "@odata.type"       = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                settingDefinitionId = $DefId
                simpleSettingValue  = @{
                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                    value         = $Value
                }
            }
        }) | Out-Null
    }

    if ($S.LockoutThreshold  -gt 0) { Add-IntSetting "device_vendor_msft_policy_config_accountpolicies_accountlockoutpolicy_accountlockoutthreshold"      $S.LockoutThreshold }
    if ($null -ne $S.LockoutDuration) { Add-IntSetting "device_vendor_msft_policy_config_accountpolicies_accountlockoutpolicy_accountlockoutduration"        $S.LockoutDuration }
    if ($S.ResetLockoutCount -gt 0) { Add-IntSetting "device_vendor_msft_policy_config_accountpolicies_accountlockoutpolicy_resetaccountlockoutcounterafter" $S.ResetLockoutCount }

    if ($settings.Count -eq 0) { return $null }

    $body = @{ name = $Name; description = $Desc; platforms = "windows10"; technologies = "mdm"; settings = @($settings) }
    return Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/configurationPolicies" -Body $body
}

function New-AuditPolicySettings {
    param([string]$Name, [string]$Desc, $S)

    # SettingDefinitionId → category name (Intune audit settings catalog IDs)
    # NOTE: Verify these IDs against your tenant via:
    #   GET /beta/deviceManagement/configurationSettings?$filter=startsWith(id,'device_vendor_msft_policy_config_audit')
    $auditMap = @{
        "AccountLogon"      = "device_vendor_msft_policy_config_audit_accountlogon_auditcredentialvalidation"
        "AccountManagement" = "device_vendor_msft_policy_config_audit_accountmanagement_auditsecuritygroupmanagement"
        "LogonLogoff"       = "device_vendor_msft_policy_config_audit_logonlogoff_auditlogon"
        "ObjectAccess"      = "device_vendor_msft_policy_config_audit_objectaccess_auditfileshare"
        "PolicyChange"      = "device_vendor_msft_policy_config_audit_policychange_auditauditpolicychange"
        "PrivilegeUse"      = "device_vendor_msft_policy_config_audit_privilegeuse_auditsensitiveprivilegeuse"
        "SystemEvents"      = "device_vendor_msft_policy_config_audit_system_auditsystemintegrity"
    }
    $auditValues = @{ 0="off"; 1="success"; 2="failure"; 3="successtandfailure" }

    $settings = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($item in $S) {
        $defId = $auditMap[$item.Category]
        if (-not $defId) { continue }
        $v = $auditValues[[int]$item.Value]
        if (-not $v) { $v = "off" }

        $settings.Add(@{
            "@odata.type"   = "#microsoft.graph.deviceManagementConfigurationSetting"
            settingInstance = @{
                "@odata.type"       = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                settingDefinitionId = $defId
                choiceSettingValue  = @{
                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue"
                    value         = "${defId}_${v}"
                    children      = @()
                }
            }
        }) | Out-Null
    }

    if ($settings.Count -eq 0) { return $null }

    $body = @{ name = $Name; description = $Desc; platforms = "windows10"; technologies = "mdm"; settings = @($settings) }
    return Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/configurationPolicies" -Body $body
}

function New-FirewallProfilePolicy {
    param([string]$Name, [string]$Desc, $Profiles)

    $settings = [System.Collections.Generic.List[hashtable]]::new()

    function Add-FwChoice {
        param([string]$DefId, [string]$ValueSuffix)
        $settings.Add(@{
            "@odata.type"   = "#microsoft.graph.deviceManagementConfigurationSetting"
            settingInstance = @{
                "@odata.type"       = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                settingDefinitionId = $DefId
                choiceSettingValue  = @{
                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue"
                    value         = "${DefId}_${ValueSuffix}"
                    children      = @()
                }
            }
        }) | Out-Null
    }

    # NOTE: These settingDefinitionIds are from the Windows Firewall CSP.
    # Verify against your tenant: GET /beta/deviceManagement/configurationSettings?$filter=startsWith(id,'vendor_msft_firewall')
    $profileMap = @(
        @{ Key = "Domain";  Prefix = "vendor_msft_firewall_mdmstore_domainprofile" }
        @{ Key = "Private"; Prefix = "vendor_msft_firewall_mdmstore_privateprofile" }
        @{ Key = "Public";  Prefix = "vendor_msft_firewall_mdmstore_publicprofile" }
    )

    foreach ($pm in $profileMap) {
        $p = $Profiles.($pm.Key)
        if (-not $p) { continue }

        if ($p.Enabled -match "^(true|1|yes)$") { Add-FwChoice "$($pm.Prefix)_enablefirewall" "true" }
        elseif ($p.Enabled -match "^(false|0|no)$") { Add-FwChoice "$($pm.Prefix)_enablefirewall" "false" }

        if ($p.DefaultInbound -match "block|0")  { Add-FwChoice "$($pm.Prefix)_defaultinboundaction" "block" }
        elseif ($p.DefaultInbound -match "allow|1") { Add-FwChoice "$($pm.Prefix)_defaultinboundaction" "allow" }

        if ($p.DefaultOutbound -match "block|0") { Add-FwChoice "$($pm.Prefix)_defaultoutboundaction" "block" }
        elseif ($p.DefaultOutbound -match "allow|1") { Add-FwChoice "$($pm.Prefix)_defaultoutboundaction" "allow" }

        if ($p.Notifications -match "^(true|1)$") { Add-FwChoice "$($pm.Prefix)_disablenotifications" "true" }
    }

    if ($settings.Count -eq 0) { return $null }

    $body = @{
        name         = $Name
        description  = $Desc
        platforms    = "windows10"
        technologies = "mdm,microsoftSense"
        settings     = @($settings)
    }
    return Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/configurationPolicies" -Body $body
}

function New-FirewallRulesPolicy {
    param([string]$Name, [string]$Desc, $Rules)

    # Firewall rules use OMA-URI via custom configuration profile
    # Each rule: ./Vendor/MSFT/Firewall/MdmStore/FirewallRules/{RuleId}/...
    $omaSettings = [System.Collections.Generic.List[hashtable]]::new()
    $idx = 1

    foreach ($rule in $Rules) {
        $ruleId = "GPO_Rule_$idx"
        $idx++
        $baseUri = "./Vendor/MSFT/Firewall/MdmStore/FirewallRules/$ruleId"

        $addOma = {
            param([string]$Path, [string]$Value, [string]$Type = "String")
            $omaSettings.Add(@{
                "@odata.type" = "#microsoft.graph.omaSettingString"
                displayName   = "$ruleId - $($Path.Split('/')[-1])"
                omaUri        = "$baseUri/$Path"
                value         = $Value
            }) | Out-Null
        }

        if ($rule.Name)      { & $addOma "Name"         $rule.Name }
        if ($rule.Action)    { & $addOma "Action"        (if ($rule.Action -match "allow") { "1" } else { "0" }) "Integer" }
        if ($rule.Direction) { & $addOma "Direction"     (if ($rule.Direction -match "in")  { "In" } else { "Out" }) }
        if ($rule.Enabled)   { & $addOma "Enabled"       (if ($rule.Enabled -match "true|1") { "true" } else { "false" }) "Boolean" }
        if ($rule.LocalPort) { & $addOma "LocalPortRanges" $rule.LocalPort }
        if ($rule.RemotePort){ & $addOma "RemotePortRanges" $rule.RemotePort }
        if ($rule.App)       { & $addOma "App/FilePath"  $rule.App }
        if ($rule.Profile)   { & $addOma "ProfilesSpecified" $rule.Profile }
    }

    if ($omaSettings.Count -eq 0) { return $null }

    $body = @{
        "@odata.type"  = "#microsoft.graph.windows10CustomConfiguration"
        displayName    = $Name
        description    = $Desc
        omaSettings    = @($omaSettings)
    }
    return Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/deviceConfigurations" -Body $body
}

function New-AdmxPolicy {
    param([string]$Name, [string]$Desc, $Settings)

    # Create a Group Policy Configuration (Administrative Templates)
    $gpConfig = Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/groupPolicyConfigurations" -Body @{
        displayName = $Name
        description = $Desc
    }

    $configId = $gpConfig.id
    $matched  = 0
    $missing  = 0

    foreach ($s in $Settings) {
        if (-not $s.Name) { $missing++; continue }
        try {
            # Search for matching definition
            $encoded = [System.Uri]::EscapeDataString($s.Name)
            $search  = Invoke-GGet "$GRAPH_BASE/deviceManagement/groupPolicyDefinitions?`$filter=displayName eq '$encoded'&`$top=1"
            if ($search.value -and $search.value.Count -gt 0) {
                $defId = $search.value[0].id
                $isEnabled = $s.State -ne "Disabled"
                Invoke-GPost -Uri "$GRAPH_BASE/deviceManagement/groupPolicyConfigurations/$configId/definitionValues" -Body @{
                    enabled                    = $isEnabled
                    "definition@odata.bind"    = "$GRAPH_BASE/deviceManagement/groupPolicyDefinitions('$defId')"
                } | Out-Null
                $matched++
            } else {
                $missing++
                Write-Log "  ADMX setting not found: '$($s.Name)'" "WARN"
            }
        } catch {
            $missing++
            Write-Log "  ADMX setting error '$($s.Name)': $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "  Admin Template: $matched configured / $missing not found in tenant"
    return $gpConfig
}

# ──────────────────────────────────────────────────────────────
# MIGRATION EXECUTOR
# ──────────────────────────────────────────────────────────────

function Invoke-MigratePolicies {
    param([array]$Plan)

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $gpoIndex = 0

    foreach ($item in $Plan) {
        $gpoIndex++
        $pct = [int](($gpoIndex / $Plan.Count) * 100)
        Write-Progress -Activity "Migrating GPOs" -Status "[$gpoIndex/$($Plan.Count)] $($item.GPOName)" -PercentComplete $pct

        Write-Log "▶ GPO: $($item.GPOName)"

        foreach ($pol in $item.Policies) {
            $result = [PSCustomObject]@{
                GPO        = $item.GPOName
                PolicyType = $pol.PolicyType
                PolicyName = $pol.PolicyName
                Status     = "PENDING"
                IntuneId   = "-"
                Note       = ""
            }

            if ($pol.Action -eq "MANUAL") {
                $result.Status = "SKIPPED"
                $result.Note   = "Manual review required before deployment"
                $results.Add($result) | Out-Null
                Write-Log "  [SKIP] $($pol.PolicyName) — manual action needed" "WARN"
                continue
            }

            Write-Log "  [→] Creating: $($pol.PolicyName)"

            try {
                $created = switch ($pol.ApiType) {
                    "compliance"     { New-CompliancePolicy  -Name $pol.PolicyName -Desc $pol.Description -S $pol.Settings }
                    "lockout"        { New-LockoutPolicy     -Name $pol.PolicyName -Desc $pol.Description -S $pol.Settings }
                    "audit"          { New-AuditPolicySettings -Name $pol.PolicyName -Desc $pol.Description -S $pol.Settings }
                    "firewall"       { New-FirewallProfilePolicy -Name $pol.PolicyName -Desc $pol.Description -Profiles $pol.Settings }
                    "firewallRules"  { New-FirewallRulesPolicy -Name $pol.PolicyName -Desc $pol.Description -Rules $pol.Settings }
                    "admx"           { New-AdmxPolicy        -Name $pol.PolicyName -Desc $pol.Description -Settings $pol.Settings }
                    "securityOptions" {
                        $result.Status = "SKIPPED"
                        $result.Note   = "Export to report only. Create Custom Profile manually via Intune portal using OMA-URI values below."
                        $null
                    }
                    "userRights" {
                        $result.Status = "SKIPPED"
                        $result.Note   = "User rights exported to report. Map to ./Vendor/MSFT/Policy/Config/UserRights/* OMA-URI."
                        $null
                    }
                    "bitlocker" {
                        $result.Status = "PARTIAL"
                        $result.Note   = "BitLocker settings detected. Create Disk Encryption policy in Endpoint Security manually."
                        $null
                    }
                    default { $null }
                }

                if ($created) {
                    $result.Status   = "SUCCESS"
                    $result.IntuneId = if ($created.id)  { $created.id  }
                                       elseif ($created.Id) { $created.Id }
                                       else { "-" }
                    Write-Log "  [OK] $($result.IntuneId)" "SUCCESS"
                } elseif ($result.Status -eq "PENDING") {
                    $result.Status = "SKIPPED"
                    $result.Note   = "No mappable settings"
                }
            } catch {
                $result.Status = "FAILED"
                $result.Note   = $_.Exception.Message
                Write-Log "  [FAIL] $($_.Exception.Message)" "ERROR"
            }

            $results.Add($result) | Out-Null
        }
    }

    Write-Progress -Activity "Migrating GPOs" -Completed
    return $results
}

# ──────────────────────────────────────────────────────────────
# HTML REPORT
# ──────────────────────────────────────────────────────────────

function New-HtmlReport {
    param([array]$Results, [array]$Plan, [string]$OutDir)

    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $fname = "GPO-Intune-Migration-$(Get-Date -Format 'yyyyMMdd-HHmm').html"
    $fpath = Join-Path $OutDir $fname

    $successCount = ($Results | Where-Object Status -eq "SUCCESS").Count
    $failedCount  = ($Results | Where-Object Status -eq "FAILED").Count
    $skippedCount = ($Results | Where-Object Status -in @("SKIPPED","PARTIAL")).Count

    # Build main results table
    $rows = ""
    foreach ($r in $Results) {
        $color = switch ($r.Status) {
            "SUCCESS" { "#155724"; $bg = "#d4edda" }
            "FAILED"  { "#721c24"; $bg = "#f8d7da" }
            "PARTIAL" { "#856404"; $bg = "#fff3cd" }
            "SKIPPED" { "#383d41"; $bg = "#e2e3e5" }
            default   { "#000"; $bg = "#fff" }
        }
        $rows += "<tr style='background:$bg'>
            <td>$($r.GPO)</td>
            <td>$($r.PolicyType)</td>
            <td>$($r.PolicyName)</td>
            <td style='color:$color;font-weight:600'>$($r.Status)</td>
            <td style='font-family:monospace;font-size:0.85em'>$($r.IntuneId)</td>
            <td>$($r.Note)</td></tr>"
    }

    # Collect manual action items
    $manualSection = ""
    foreach ($item in $Plan) {
        $manualPols = $item.Policies | Where-Object { $_.Action -in @("MANUAL","PARTIAL") }
        foreach ($mp in $manualPols) {
            $details = ($mp.Details -join "<br>")
            $manualSection += "<tr><td>$($item.GPOName)</td><td>$($mp.PolicyType)</td><td>$($mp.PolicyName)</td><td>$details</td></tr>"
        }
    }

    # Log table
    $logRows = ""
    foreach ($l in $script:Log) {
        $lc = switch ($l.Level) {
            "SUCCESS" { "#155724" }
            "WARN"    { "#856404" }
            "ERROR"   { "#721c24" }
            default   { "#333" }
        }
        $logRows += "<tr><td>$($l.Time)</td><td style='color:$lc;font-weight:600'>$($l.Level)</td><td>$($l.Message)</td></tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>GPO → Intune Migration Report</title>
<style>
  *, *::before, *::after { box-sizing: border-box; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; margin: 0; padding: 32px 48px; background: #f4f6f9; color: #2b2b2b; }
  h1   { color: #0078d4; margin: 0 0 4px; }
  .sub { color: #666; font-size: 0.9em; margin: 0 0 32px; }
  h2   { color: #0063b1; border-bottom: 3px solid #0078d4; padding-bottom: 6px; margin-top: 40px; }
  .cards { display: flex; gap: 16px; margin: 24px 0; }
  .card { padding: 20px 28px; border-radius: 10px; color: #fff; min-width: 130px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,.15); }
  .card .n { font-size: 2.4em; font-weight: 700; line-height: 1; }
  .card .l { font-size: 0.85em; margin-top: 4px; opacity: .9; }
  .green  { background: #107c10; }
  .red    { background: #c50f1f; }
  .yellow { background: #ca5010; }
  .gray   { background: #616161; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); font-size: 0.9em; }
  thead th { background: #0078d4; color: #fff; padding: 12px 14px; text-align: left; font-weight: 600; }
  tbody td { padding: 9px 14px; border-bottom: 1px solid #e8ecf0; vertical-align: top; }
  tbody tr:last-child td { border-bottom: none; }
  tbody tr:hover { filter: brightness(0.97); }
  .warn-box { background:#fff3cd; border:1px solid #ffc107; border-radius:6px; padding:12px 16px; margin:16px 0; font-size:.9em; }
</style>
</head>
<body>
<h1>GPO → Intune Migration Report</h1>
<p class="sub">Generated: $ts</p>

<div class="cards">
  <div class="card green" ><div class="n">$successCount</div><div class="l">Succeeded</div></div>
  <div class="card red"   ><div class="n">$failedCount</div> <div class="l">Failed</div></div>
  <div class="card yellow"><div class="n">$skippedCount</div><div class="l">Skipped / Partial</div></div>
  <div class="card gray"  ><div class="n">$($Results.Count)</div><div class="l">Total</div></div>
</div>

<h2>Migration Results</h2>
<table>
<thead><tr><th>GPO</th><th>Policy Type</th><th>Policy Name</th><th>Status</th><th>Intune ID</th><th>Notes</th></tr></thead>
<tbody>$rows</tbody>
</table>

<h2>Manual Action Required</h2>
<div class="warn-box">⚠️ The following settings could <strong>not</strong> be fully automated. Review each item and configure manually in Intune.</div>
<table>
<thead><tr><th>GPO</th><th>Policy Type</th><th>Policy Name</th><th>Details / OMA-URI hints</th></tr></thead>
<tbody>$manualSection</tbody>
</table>

<h2>Execution Log</h2>
<table>
<thead><tr><th>Time</th><th>Level</th><th>Message</th></tr></thead>
<tbody>$logRows</tbody>
</table>
</body>
</html>
"@

    $html | Out-File -FilePath $fpath -Encoding UTF8 -Force
    return $fpath
}

# ──────────────────────────────────────────────────────────────
# MAIN FLOW
# ──────────────────────────────────────────────────────────────

Write-Host @"

  ╔══════════════════════════════════════════════════════════╗
  ║           GPO → Intune Migration Tool                   ║
  ║           Version 1.0  |  Windows 10/11 MDM             ║
  ╚══════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

Write-Host "  Backup path : $GPOBackupPath" -ForegroundColor White
Write-Host "  Mode        : $(if ($PlanOnly) { 'Plan Only (no changes)' } else { 'Full Migration' })" -ForegroundColor White
Write-Host "  Prefix      : '$PolicyPrefix'" -ForegroundColor White
Write-Host "  Report path : $ReportPath`n" -ForegroundColor White

# 1. Discover
Write-Section "Step 1 / 5 — Discovering GPO Backups"
$backups = Get-GPOBackupList -Path $GPOBackupPath
if ($backups.Count -eq 0) {
    Write-Log "No GPO backups found. Expected GUID-named subfolders containing gpreport.xml." "ERROR"
    exit 1
}

# 2. Parse
Write-Section "Step 2 / 5 — Parsing GPO Settings"
$reports = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($b in $backups) {
    Write-Log "Parsing: $($b.FolderName)"
    $r = Read-GPOReport -Backup $b
    if ($r) {
        $reports.Add($r) | Out-Null
        $settingCount = (
            ($r.AuditSettings.Count) + ($r.UserRights.Count) + ($r.SecurityOptions.Count) +
            ($r.FirewallRules.Count) + ($r.RegistrySettings.Count) + ($r.Scripts.Count) +
            (if ($r.PasswordPolicy) { 1 } else { 0 }) +
            (if ($r.FirewallProfiles) { 1 } else { 0 })
        )
        Write-Log "  → '$($r.Name)'  ($settingCount setting group(s) found)" "SUCCESS"
    }
}

if ($reports.Count -eq 0) {
    Write-Log "No GPO reports could be parsed." "ERROR"
    exit 1
}

# 3. Plan
Write-Section "Step 3 / 5 — Building Migration Plan"
$plan = Build-MigrationPlan -Reports $reports
Show-MigrationPlan -Plan $plan

if ($PlanOnly) {
    Write-Log "Plan-only mode. No changes made." "WARN"
    Write-Host "`n  Run without -PlanOnly to execute the migration.`n" -ForegroundColor Yellow
    exit 0
}

# 4. Confirm
Write-Host ""
$confirm = Read-Host "  ► Proceed with migration? Type 'yes' to continue"
if ($confirm -notmatch "^yes$") {
    Write-Log "Migration cancelled by user." "WARN"
    exit 0
}

# 5. Connect & Execute
Write-Section "Step 4 / 5 — Connecting to Microsoft Graph"
Initialize-GraphConnection

Write-Section "Step 5 / 5 — Creating Intune Policies"
$results = Invoke-MigratePolicies -Plan $plan
$script:MigrationResults.AddRange($results)

# Summary
Write-Section "Migration Complete"
$s = ($results | Where-Object Status -eq "SUCCESS").Count
$f = ($results | Where-Object Status -eq "FAILED").Count
$k = ($results | Where-Object Status -in @("SKIPPED","PARTIAL")).Count

Write-Host "  Succeeded : $s" -ForegroundColor Green
Write-Host "  Failed    : $f" -ForegroundColor Red
Write-Host "  Skipped   : $k" -ForegroundColor Yellow

$reportFile = New-HtmlReport -Results $results -Plan $plan -OutDir $ReportPath
Write-Log "HTML Report: $reportFile" "SUCCESS"

# Open report on Windows
if ($IsWindows -or $env:OS -match "Windows") {
    try { Start-Process $reportFile } catch { }
}

Write-Host "`n  Done. Review the report and complete any MANUAL/PARTIAL items in Intune.`n" -ForegroundColor Cyan
