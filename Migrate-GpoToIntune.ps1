<#
.SYNOPSIS
    Migrates on-prem GPOs to Intune Settings Catalog via Group Policy Analytics.

.DESCRIPTION
    Pipeline:
      1. Preflight (modules, RSAT, Graph SDK)
      2. Connect-MgGraph (device login)
      3. Enumerate linked+enabled GPOs from live AD
      4. Skip GPOs whose "MIG-<Name>" already exists in Intune
      5. Submit XML to Group Policy Analytics (beta API)
      6. Classify settings: Edge / Chrome / Firefox / Other / Unmapped
      7. Per-policy confirmation, create Settings Catalog policy
      8. Unmapped -> Unmapped.json with OMA-URI hints for manual work

    Each step asks for confirmation. Per-run folder:
        .\GPOMigration\<timestamp>\
            transcript.log        (full PS transcript)
            migration.log         (structured NDJSON, one event per line)
            <PolicyName>.json     (preview of body before push)
            Edge.json, Chrome.json, Firefox.json, Other.json, Unmapped.json

.NOTES
    Requires:
      - RSAT: GroupPolicy module
      - Microsoft.Graph.Authentication module
      - Run on domain-joined host with rights to read all GPOs
      - Graph scopes: DeviceManagementConfiguration.ReadWrite.All, Group.Read.All

    Caveats:
      - Group Policy Analytics + configurationPolicies are beta endpoints.
      - Chrome/Firefox map fully only if their ADMX has been ingested in the
        tenant (Devices > Configuration > Import ADMX). Otherwise their settings
        will fall through to Unmapped.json.
      - Script creates policies UNASSIGNED. Original GPO links are stamped in
        the policy description for later targeting.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$OutputRoot = ".\GPOMigration",
    [string]$NamePrefix = "MIG-",
    [string]$Domain     = $env:USERDNSDOMAIN,
    [int]   $AnalyticsTimeoutSec = 180,
    [switch]$SkipPreflight
)

#region --- Setup -------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$runDir    = Join-Path $OutputRoot $timestamp
$null      = New-Item -Path $runDir -ItemType Directory -Force

$transcript    = Join-Path $runDir 'transcript.log'
$structuredLog = Join-Path $runDir 'migration.log'
$bucketFiles   = @{
    Edge     = Join-Path $runDir 'Edge.json'
    Chrome   = Join-Path $runDir 'Chrome.json'
    Firefox  = Join-Path $runDir 'Firefox.json'
    Other    = Join-Path $runDir 'Other.json'
    Unmapped = Join-Path $runDir 'Unmapped.json'
}

Start-Transcript -Path $transcript -Force | Out-Null

function Write-Log {
    param(
        [Parameter(Mandatory)][ValidateSet('INFO','WARN','ERROR','STEP','OK','SKIP')][string]$Level,
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Data = @{}
    )
    $entry = [ordered]@{
        ts    = (Get-Date).ToString('o')
        level = $Level
        msg   = $Message
        data  = $Data
    }
    ($entry | ConvertTo-Json -Compress -Depth 10) | Add-Content -Path $structuredLog
    $color = @{INFO='Gray';WARN='Yellow';ERROR='Red';STEP='Cyan';OK='Green';SKIP='DarkGray'}[$Level]
    Write-Host "[$Level] $Message" -ForegroundColor $color
}

function Confirm-Step {
    param([Parameter(Mandatory)][string]$Prompt)
    Write-Host ""
    Write-Host ">>> $Prompt" -ForegroundColor Magenta
    $r = Read-Host "    Proceed? [y/N]"
    return $r -match '^(y|yes)$'
}
#endregion

#region --- Step 1: Preflight -------------------------------------------------
function Invoke-Preflight {
    Write-Log STEP "Preflight: checking modules and tooling"
    foreach ($m in 'GroupPolicy','Microsoft.Graph.Authentication') {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            throw "Required module missing: $m. Run: Install-Module $m -Scope CurrentUser"
        }
        Import-Module $m -ErrorAction Stop | Out-Null
    }
    if (-not (Get-Command Get-GPO -ErrorAction SilentlyContinue)) {
        throw "Get-GPO not available. Install RSAT: GPMC."
    }
    Write-Log OK "Preflight passed" @{ ps = $PSVersionTable.PSVersion.ToString() }
}
#endregion

#region --- Step 2: Connect to Graph ------------------------------------------
function Connect-Graph {
    Write-Log STEP "Connecting to Microsoft Graph (device login)"
    $scopes = @(
        'DeviceManagementConfiguration.ReadWrite.All',
        'Group.Read.All'
    )
    Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Graph connection failed." }
    Write-Log OK "Connected" @{ tenant = $ctx.TenantId; account = $ctx.Account }
}
#endregion

#region --- Step 3: Enumerate linked+enabled GPOs -----------------------------
function Get-LinkedEnabledGpo {
    Write-Log STEP "Enumerating linked+enabled GPOs"
    $candidates = @()
    foreach ($gpo in (Get-GPO -All)) {
        if ($gpo.GpoStatus -eq 'AllSettingsDisabled') { continue }
        try {
            [xml]$rpt = Get-GPOReport -Guid $gpo.Id -ReportType Xml
        } catch {
            Write-Log WARN "Skipping GPO (XML report failed)" @{ gpo = $gpo.DisplayName; err = $_.Exception.Message }
            continue
        }
        $links = @($rpt.GPO.LinksTo)
        if (-not $links) { continue }
        $hasEnabled = @($links | Where-Object { $_.Enabled -eq 'true' }).Count -gt 0
        if (-not $hasEnabled) { continue }

        $candidates += [pscustomobject]@{
            Id          = $gpo.Id
            DisplayName = $gpo.DisplayName
            GpoStatus   = $gpo.GpoStatus
            LinkPaths   = ($links | ForEach-Object { $_.SOMPath }) -join '; '
            Xml         = $rpt.OuterXml
        }
    }
    Write-Log OK "Found $($candidates.Count) linked+enabled GPOs" @{ count = $candidates.Count }
    return ,$candidates
}
#endregion

#region --- Step 4: Detect already-migrated -----------------------------------
function Get-ExistingMigratedName {
    Write-Log STEP "Fetching existing Intune policies with prefix '$NamePrefix'"
    $names = @()
    $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$select=name,id&`$top=100"
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
        $names += ($resp.value | Where-Object { $_.name -like "$NamePrefix*" } | Select-Object -ExpandProperty name)
        $uri = $resp.'@odata.nextLink'
    } while ($uri)
    Write-Log OK "$($names.Count) existing migrated policies found"
    return $names
}
#endregion

#region --- Step 5: Group Policy Analytics ------------------------------------
function Submit-AnalyticsReport {
    param(
        [Parameter(Mandatory)][string]$GpoName,
        [Parameter(Mandatory)][string]$Xml,
        [string]$OuDistinguishedName = ''
    )
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Xml))
    $body = @{
        groupPolicyObjectFile = @{
            ouDistinguishedName = $OuDistinguishedName
            content             = $b64
        }
    } | ConvertTo-Json -Depth 5

    return Invoke-MgGraphRequest -Method POST `
        -Uri 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyMigrationReports/createMigrationReport' `
        -Body $body -ContentType 'application/json'
}

function Wait-AnalyticsReport {
    param([Parameter(Mandatory)][string]$ReportId)
    $deadline = (Get-Date).AddSeconds($AnalyticsTimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $r = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyMigrationReports/$ReportId"
        if ($r.migrationReadiness -ne 'processing') { return $r }
        Start-Sleep -Seconds 4
    }
    throw "Analytics report $ReportId did not complete within $AnalyticsTimeoutSec s"
}

function Get-AnalyticsMapping {
    param([Parameter(Mandatory)][string]$ReportId)
    $items = @()
    $uri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyMigrationReports/$ReportId/groupPolicySettingMappings"
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
        $items += $resp.value
        $uri = $resp.'@odata.nextLink'
    } while ($uri)
    return $items
}
#endregion

#region --- Step 6: Classify --------------------------------------------------
function Get-Bucket {
    param([Parameter(Mandatory)]$Mapping)
    $haystack = @(
        $Mapping.admxExtensionName,
        $Mapping.parentSectionDisplayName,
        $Mapping.settingName,
        $Mapping.adminTemplateCategoryPathDisplayName
    ) -join ' '

    if ($haystack -match '(?i)Microsoft\s*Edge')         { return 'Edge' }
    if ($haystack -match '(?i)Google|Chrome')            { return 'Chrome' }
    if ($haystack -match '(?i)Firefox|Mozilla')          { return 'Firefox' }
    if (($Mapping.PSObject.Properties.Name -contains 'mappedToService' -and -not $Mapping.mappedToService) `
        -or -not $Mapping.settingInstance) {
        return 'Unmapped'
    }
    return 'Other'
}
#endregion

#region --- Step 7: Build / push Settings Catalog policy ----------------------
function ConvertTo-SettingInstanceWrapper {
    param([Parameter(Mandatory)]$Mapping)
    if (-not $Mapping.settingInstance) { return $null }
    return @{
        '@odata.type'   = '#microsoft.graph.deviceManagementConfigurationSetting'
        settingInstance = $Mapping.settingInstance
    }
}

function New-SettingsCatalogPolicy {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][object[]]$Settings
    )
    $body = @{
        name         = $Name
        description  = $Description
        platforms    = 'windows10'
        technologies = 'mdm'
        settings     = $Settings
    } | ConvertTo-Json -Depth 30

    return Invoke-MgGraphRequest -Method POST `
        -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' `
        -Body $body -ContentType 'application/json'
}
#endregion

#region --- Step 8: Unmapped export -------------------------------------------
function Export-UnmappedRecord {
    param(
        [Parameter(Mandatory)][string]$GpoName,
        [Parameter(Mandatory)]$Mapping
    )
    [pscustomobject]@{
        gpo               = $GpoName
        settingName       = $Mapping.settingName
        category          = $Mapping.adminTemplateCategoryPathDisplayName
        admxExtension     = $Mapping.admxExtensionName
        parentSection     = $Mapping.parentSectionDisplayName
        settingValue      = $Mapping.settingValueAsString
        suggestion        = "Manual review: ingest matching ADMX and re-run, OR create custom OMA-URI under ./Device/Vendor/MSFT/Policy/Config/<area>/<setting> using values above."
    }
}
#endregion

#region --- Main orchestration ------------------------------------------------
try {
    Write-Log INFO "Run started" @{ outputDir = $runDir; domain = $Domain; prefix = $NamePrefix }

    if (-not $SkipPreflight) {
        if (-not (Confirm-Step "Step 1/4: Run preflight checks")) { throw "Aborted by user." }
        Invoke-Preflight
    }

    if (-not (Confirm-Step "Step 2/4: Connect to Microsoft Graph (device login)")) { throw "Aborted by user." }
    Connect-Graph

    if (-not (Confirm-Step "Step 3/4: Enumerate linked+enabled GPOs from $Domain")) { throw "Aborted by user." }
    $gpos = Get-LinkedEnabledGpo
    if ($gpos.Count -eq 0) { Write-Log WARN "Nothing to do."; return }

    $existing = Get-ExistingMigratedName

    if (-not (Confirm-Step "Step 4/4: Process $($gpos.Count) GPOs (you'll confirm each policy push)")) {
        throw "Aborted by user."
    }

    foreach ($gpo in $gpos) {
        Write-Host ""
        Write-Log STEP "GPO: $($gpo.DisplayName)" @{ id = $gpo.Id; links = $gpo.LinkPaths }

        # Skip already-migrated (any bucket policy starting with MIG-<Name>)
        $alreadyDone = @($existing | Where-Object { $_ -like "$NamePrefix$($gpo.DisplayName)*" }).Count -gt 0
        if ($alreadyDone) {
            Write-Log SKIP "Already migrated (matching '$NamePrefix$($gpo.DisplayName)*' exists)"
            continue
        }

        if (-not (Confirm-Step "Submit '$($gpo.DisplayName)' to Group Policy Analytics")) {
            Write-Log SKIP "User skipped" @{ gpo = $gpo.DisplayName }
            continue
        }

        try {
            $report   = Submit-AnalyticsReport -GpoName $gpo.DisplayName -Xml $gpo.Xml
            $reportId = $report.id
            $report   = Wait-AnalyticsReport -ReportId $reportId
            $maps     = Get-AnalyticsMapping  -ReportId $reportId
        } catch {
            Write-Log ERROR "Analytics failed" @{ gpo = $gpo.DisplayName; err = $_.Exception.Message }
            continue
        }
        Write-Log OK "Analytics complete" @{ gpo = $gpo.DisplayName; settings = $maps.Count; readiness = $report.migrationReadiness }

        # Classify
        $byBucket = @{ Edge=@(); Chrome=@(); Firefox=@(); Other=@(); Unmapped=@() }
        foreach ($m in $maps) { $byBucket[(Get-Bucket -Mapping $m)] += $m }

        Write-Log INFO "Classification" @{
            gpo      = $gpo.DisplayName
            edge     = $byBucket.Edge.Count
            chrome   = $byBucket.Chrome.Count
            firefox  = $byBucket.Firefox.Count
            other    = $byBucket.Other.Count
            unmapped = $byBucket.Unmapped.Count
        }

        foreach ($b in 'Edge','Chrome','Firefox','Other') {
            $items = $byBucket[$b]
            if ($items.Count -eq 0) { continue }

            $instances = @($items | ForEach-Object { ConvertTo-SettingInstanceWrapper -Mapping $_ } | Where-Object { $_ })
            if ($instances.Count -eq 0) {
                Write-Log WARN "Bucket has items but no settingInstance objects" @{ gpo = $gpo.DisplayName; bucket = $b }
                continue
            }

            $policyName  = "$NamePrefix$($gpo.DisplayName)-$b"
            $description = "Migrated from GPO '$($gpo.DisplayName)' [$($gpo.Id)]. Original links: $($gpo.LinkPaths). Generated $timestamp."
            $previewPath = Join-Path $runDir ("preview-" + ($policyName -replace '[^\w\-]','_') + '.json')

            @{ name=$policyName; description=$description; platforms='windows10'; technologies='mdm'; settings=$instances } |
                ConvertTo-Json -Depth 30 | Set-Content -Path $previewPath -Encoding UTF8

            if (-not (Confirm-Step "Create '$policyName' ($($instances.Count) settings) - preview: $previewPath")) {
                Write-Log SKIP "User skipped policy push" @{ policy = $policyName }
                continue
            }

            try {
                $created = New-SettingsCatalogPolicy -Name $policyName -Description $description -Settings $instances
                Write-Log OK "Policy created" @{ policy = $policyName; id = $created.id; bucket = $b }
                Add-Content -Path $bucketFiles[$b] -Value (
                    @{ gpo=$gpo.DisplayName; gpoId=$gpo.Id; policy=$policyName; policyId=$created.id; settings=$instances.Count } |
                        ConvertTo-Json -Compress
                )
            } catch {
                Write-Log ERROR "Policy push failed" @{ policy = $policyName; err = $_.Exception.Message }
            }
        }

        if ($byBucket.Unmapped.Count -gt 0) {
            foreach ($m in $byBucket.Unmapped) {
                Add-Content -Path $bucketFiles.Unmapped -Value (
                    (Export-UnmappedRecord -GpoName $gpo.DisplayName -Mapping $m) | ConvertTo-Json -Compress
                )
            }
            Write-Log WARN "$($byBucket.Unmapped.Count) settings need manual review" @{ gpo = $gpo.DisplayName; file = $bucketFiles.Unmapped }
        }
    }

    Write-Log OK "Run complete" @{ outputDir = $runDir }
}
catch {
    Write-Log ERROR $_.Exception.Message
    throw
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Stop-Transcript | Out-Null
}
#endregion
