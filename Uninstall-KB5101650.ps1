<#
.SYNOPSIS
    Desinstal·la l'actualització conflictiva KB5101650 de Windows 11 24H2
.DESCRIPTION
    Script per eliminar l'actualització KB5101650 (14/07) d'equips amb Windows 11 24H2.
    Compatible amb execució local, SCCM i Intune.
    Gestiona estats "Instal·lat" i "Pendent de reinici".
    No força el reinici (retorna 3010 si cal).

    Exit Codes:
      0    - KB no present o desinstal·lat correctament sense reinici pendent
      170  - KB desinstal·lat, es requereix reinici (compatible SCCM)
      3010 - KB desinstal·lat, cal reiniciar (soft reboot)
      1    - Error

.PARAMETER KB
    Número de KB a desinstal·lar. Per defecte: KB5101650

.PARAMETER LogPath
    Ruta completa del fitxer de log.
    Per defecte: %SystemRoot%\Logs\Uninstall-KB5101650.log

.PARAMETER NoRestart
    No reiniciar l'equip en cap cas (activat per defecte)

.EXAMPLE
    # Execució local
    .\Uninstall-KB5101650.ps1

.EXAMPLE
    # Execució amb log personalitzat
    .\Uninstall-KB5101650.ps1 -LogPath "C:\Temp\remove-kb.log"

.EXAMPLE
    # SCCM deployment (el client ja captura stdout)
    powershell.exe -ExecutionPolicy Bypass -File "Uninstall-KB5101650.ps1"

.EXAMPLE
    # Intune deployment (script únic)
    .\Uninstall-KB5101650.ps1 -Verbose

.NOTES
    Versió:  1.0
    Autor:   Script generat per OpenCode
    Data:    2026-07-17
    SO:      Windows 11 24H2
    KB:      KB5101650
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$KB = "KB5101650",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:SystemRoot\Logs\Uninstall-KB5101650.log",

    [Parameter(Mandatory = $false)]
    [switch]$NoRestart = $true
)

# ============================================================================
# REGIÓ: Inicialització i helpers
# ============================================================================

$ScriptVersion  = "1.0"
$ScriptStartTime = Get-Date
$RebootRequired  = $false
$ExitCode        = 1
$KB_Numeric      = $KB -replace '[^0-9]', ''

# Determinar el context d'execució
$IsLocal  = $false
$IsSCCM   = $false
$IsIntune = $false

if (Test-Path "env:SMS_LOCAL_DIR") {
    $IsSCCM = $true
}
if ((Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue) -or
    (Test-Path "$env:ProgramData\Microsoft\IntuneManagementExtension")) {
    $IsIntune = $true
}
if (-not $IsSCCM -and -not $IsIntune) {
    $IsLocal = $true
}

# Assegurar que el directori de log existeix
$LogDir = Split-Path -Path $LogPath -Parent
try {
    if (-not (Test-Path -Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force -ErrorAction Stop | Out-Null
    }
} catch {
    Write-Warning "No s'ha pogut crear el directori de log: $LogDir. Usant ubicació alternativa."
    $LogPath = "$env:TEMP\Uninstall-KB5101650.log"
    $LogDir  = Split-Path -Path $LogPath -Parent
    if (-not (Test-Path -Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

# Funció de logging dual (consola + fitxer)
function Write-TSLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine   = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "ERROR"   { Write-Host $logLine -ForegroundColor Red }
        "WARN"    { Write-Host $logLine -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logLine -ForegroundColor Green }
        "DEBUG"   { Write-Host $logLine -ForegroundColor Gray }
        default   { Write-Host $logLine }
    }

    try {
        Add-Content -Path $LogPath -Value $logLine -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Warning "No s'ha pogut escriure al fitxer de log: $_"
    }
}

# Verificar privilegis d'administrador
function Test-IsAdministrator {
    $identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ============================================================================
# REGIÓ: Funcions de detecció
# ============================================================================

function Get-UpdateStatus {
    <#
    .SYNOPSIS
        Detecta si el KB esta instal·lat i en quin estat es troba.
        Retorna un objecte amb les propietats IsInstalled, IsPending, PackageName, DetectMethod.
    #>
    [CmdletBinding()]
    param([string]$KB)

    $result = [PSCustomObject]@{
        IsInstalled    = $false
        IsPending      = $false
        PackageName    = $null
        DetectMethod   = $null
        InstalledOn    = $null
        Details        = @()
    }

    # Mètode 1: Get-HotFix (Win32_QuickFixEngineering)
    try {
        $hotfix = Get-HotFix -Id $KB -ErrorAction Stop
        if ($hotfix) {
            $result.IsInstalled  = $true
            $result.InstalledOn  = $hotfix.InstalledOn
            $result.DetectMethod = "Get-HotFix"
            $result.Details     += "Get-HotFix: InstalledOn=$($hotfix.InstalledOn), InstalledBy=$($hotfix.InstalledBy), Description=$($hotfix.Description)"
            Write-TSLog "KB $KB detectat via Get-HotFix (InstalledOn: $($hotfix.InstalledOn))" -Level "INFO"
        }
    } catch {
        Write-TSLog "Get-HotFix: KB $KB no detectat per aquesta via" -Level "DEBUG"
    }

    # Mètode 2: DISM Get-WindowsPackage (direct name match)
    try {
        $dismPkgs = Get-WindowsPackage -Online -ErrorAction Stop |
            Where-Object { $_.PackageName -like "*$($KB_Numeric)*" -or
                           $_.PackageName -like "*$($KB)*" }

        if ($dismPkgs) {
            foreach ($pkg in $dismPkgs) {
                Write-TSLog "DISM direct: $($pkg.PackageName) | State=$($pkg.PackageState) | ReleaseType=$($pkg.ReleaseType) | InstallTime=$($pkg.InstallTime)" -Level "DEBUG"
                $result.Details += "DISM direct: $($pkg.PackageName) | State=$($pkg.PackageState)"

                switch ($pkg.PackageState) {
                    "Installed"        {
                        $result.IsInstalled = $true
                        if (-not $result.PackageName) { $result.PackageName = $pkg.PackageName }
                    }
                    "InstallPending"   {
                        $result.IsInstalled = $true
                        $result.IsPending   = $true
                        if (-not $result.PackageName) { $result.PackageName = $pkg.PackageName }
                    }
                    "Staged"           {
                        $result.IsInstalled = $true
                        $result.IsPending   = $true
                        if (-not $result.PackageName) { $result.PackageName = $pkg.PackageName }
                    }
                    "UninstallPending" {
                        $result.IsPending = $true
                        Write-TSLog "DISM: KB $KB ja té una desinstal·lació pendent de reinici" -Level "WARN"
                    }
                    "Superseded"       {
                        Write-TSLog "DISM: KB $KB trobat però en estat Superseded" -Level "DEBUG"
                    }
                    default {
                        Write-TSLog "DISM: KB $KB trobat amb estat desconegut: $($pkg.PackageState)" -Level "WARN"
                    }
                }
            }
            if (-not $result.DetectMethod) { $result.DetectMethod = "DISM" }
        } else {
            Write-TSLog "DISM direct: Cap paquet amb nom coincident amb $KB_Numeric" -Level "DEBUG"
        }
    } catch {
        Write-TSLog "DISM Get-WindowsPackage ha fallat: $_" -Level "WARN"
    }

    # Mètode 3: DISM.exe /format:list amb scanning complet de tots els camps
    Write-TSLog "DISM.exe: Escanejant tots els paquets (pot trigar uns segons)..." -Level "DEBUG"
    try {
        $dismExeOutput = & dism.exe /online /get-packages /format:list /english 2>&1
        $packageBlock  = $null
        $captureBlock  = $false
        $lastBlockProcessed = $false

        foreach ($line in $dismExeOutput) {
            if ($line -match "Package Identity") {
                # Processem el bloc anterior abans de començar-ne un de nou
                if ($captureBlock -and $packageBlock) {
                    _Process-DismBlock -Block $packageBlock -KB $KB -KB_Numeric $KB_Numeric -Result $result
                }
                $packageBlock = $line
                $captureBlock = $true
            } elseif ($captureBlock) {
                $packageBlock += "`n$line"
            }
        }
        # Processar l'últim bloc capturat (aquest és el fix clau)
        if ($captureBlock -and $packageBlock) {
            _Process-DismBlock -Block $packageBlock -KB $KB -KB_Numeric $KB_Numeric -Result $result
        }
    } catch {
        Write-TSLog "DISM.exe ha fallat: $_" -Level "WARN"
    }

    # Mètode 4: Comprovar pending.xml (pendent de reinici)
    $pendingXmlPath = "$env:SystemRoot\WinSxS\pending.xml"
    if (Test-Path -Path $pendingXmlPath) {
        try {
            [xml]$pendingXml = Get-Content -Path $pendingXmlPath -ErrorAction Stop
            $nodes = $pendingXml.SelectNodes("//*")
            $foundInPending = $false
            foreach ($node in $nodes) {
                if ($node.OuterXml -match $KB_Numeric) {
                    $foundInPending = $true
                    break
                }
            }
            if ($foundInPending) {
                $result.IsPending = $true
                Write-TSLog "pending.xml: KB $KB detectat en operacions pendents" -Level "WARN"
                $result.Details += "pending.xml: operacions pendents detectades"
            }
        } catch {
            Write-TSLog "Error al llegir pending.xml: $_" -Level "DEBUG"
        }
    }

    # Mètode 5: CBS Registry (Component Based Servicing) — PSChildName matching
    $cbsRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
    if (Test-Path -Path $cbsRegPath) {
        try {
            $cbsPkgs = Get-ChildItem -Path $cbsRegPath -ErrorAction Stop |
                Where-Object { $_.PSChildName -like "*$($KB_Numeric)*" } |
                Select-Object -First 5

            if ($cbsPkgs) {
                Write-TSLog "CBS PSChildName: Trobats $($cbsPkgs.Count) registres per KB $KB" -Level "DEBUG"
                foreach ($cbs in $cbsPkgs) {
                    $props = Get-ItemProperty -Path $cbs.PSPath -ErrorAction SilentlyContinue
                    $installState = $props.CurrentState
                    $result.Details += "CBS PSChildName: $($cbs.PSChildName) | CurrentState=$installState"
                    Write-TSLog "CBS PSChildName: $($cbs.PSChildName) | CurrentState=$installState" -Level "DEBUG"

                    if (-not $result.PackageName) { $result.PackageName = $cbs.PSChildName }
                    if ($installState -in @(0x40, 0x50, 0x70)) {
                        $result.IsInstalled = $true
                        if (-not $result.DetectMethod) { $result.DetectMethod = "CBS Registry" }
                    }
                    if ($installState -eq 0x40) { $result.IsPending = $true }
                    if ($installState -in @(0x50, 0x60)) { $result.IsPending = $true }
                }
            }
        } catch {
            Write-TSLog "Error al llegir CBS Registry: $_" -Level "DEBUG"
        }

        # Mètode 5b: CBS Registry per InstallName (clau per quan el KB no surt al PSChildName)
        if ($result.IsInstalled -and -not $result.PackageName) {
            try {
                Write-TSLog "CBS: Buscant per InstallName='$KB' als registres CBS..." -Level "DEBUG"
                $allCbsKeys = Get-ChildItem -Path $cbsRegPath -ErrorAction Stop
                $foundInCbsInstallName = $false
                foreach ($cbsKey in $allCbsKeys) {
                    try {
                        $installName = (Get-ItemProperty -Path $cbsKey.PSPath -Name "InstallName" -ErrorAction SilentlyContinue).InstallName
                        if ($installName -eq $KB -or $installName -eq $KB_Numeric) {
                            $result.PackageName = $cbsKey.PSChildName
                            $result.IsInstalled = $true
                            if (-not $result.DetectMethod) { $result.DetectMethod = "CBS InstallName" }
                            $result.Details += "CBS InstallName: $($cbsKey.PSChildName) -> $installName"
                            Write-TSLog "CBS InstallName: Trobat! $($cbsKey.PSChildName) -> $installName" -Level "SUCCESS"
                            $foundInCbsInstallName = $true
                            break
                        }
                    } catch {
                        # Alguns claus no tenen InstallName, ignorar
                    }
                }
                if (-not $foundInCbsInstallName) {
                    Write-TSLog "CBS InstallName: No s'ha trobat cap clau amb InstallName=$KB" -Level "DEBUG"
                }
            } catch {
                Write-TSLog "CBS InstallName search error: $_" -Level "DEBUG"
            }
        }
    }

    # Mètode 6: Correlació per InstallTime (fallback final)
    if ($result.IsInstalled -and -not $result.PackageName -and $result.InstalledOn) {
        try {
            Write-TSLog "InstallTime: Buscant paquets instal·lats prop de $($result.InstalledOn)..." -Level "DEBUG"
            $allPkgs = Get-WindowsPackage -Online -ErrorAction Stop
            $bestMatch = $null
            $bestDiffMinutes = 99999

            foreach ($pkg in $allPkgs) {
                if ($pkg.InstallTime -and $pkg.ReleaseType -in @('Update', 'Security Update', 'SecurityUpdate', 'Hotfix', 'UpdateEx')) {
                    $diff = [Math]::Abs(($pkg.InstallTime - $result.InstalledOn).TotalMinutes)
                    $sameDay = ($pkg.InstallTime.Date -eq $result.InstalledOn.Date)
                    if (($diff -lt 1500 -or $sameDay) -and $diff -lt $bestDiffMinutes) {
                        $bestMatch = $pkg
                        $bestDiffMinutes = $diff
                    }
                }
            }

            if ($bestMatch) {
                $result.PackageName = $bestMatch.PackageName
                Write-TSLog "InstallTime: Trobat! $($bestMatch.PackageName) (diff=$([math]::Round($bestDiffMinutes,1)) min, ReleaseType=$($bestMatch.ReleaseType))" -Level "SUCCESS"
                $result.Details += "InstallTime match: $($bestMatch.PackageName) | diff=$([math]::Round($bestDiffMinutes,1)) min | ReleaseType=$($bestMatch.ReleaseType)"
            } else {
                Write-TSLog "InstallTime: No s'ha trobat cap paquet dins d'una finestra de 25 hores ni al mateix dia" -Level "WARN"
                # Debug: mostrar paquets recents
                $recentPkgs = $allPkgs | Where-Object { $_.InstallTime } | Sort-Object InstallTime -Descending | Select-Object -First 10
                foreach ($rp in $recentPkgs) {
                    Write-TSLog "InstallTime debug: $($rp.PackageName) | $($rp.InstallTime) | $($rp.ReleaseType)" -Level "DEBUG"
                }
            }
        } catch {
            Write-TSLog "InstallTime correlation error: $_" -Level "WARN"
        }
    }

    return $result
}

# Funció auxiliar per processar blocs de sortida de DISM.exe
function _Process-DismBlock {
    param(
        [string]$Block,
        [string]$KB,
        [string]$KB_Numeric,
        [PSObject]$Result
    )

    # Buscar el KB en qualsevol part del bloc (Package Identity, Install Name, etc.)
    if ($Block -match $KB_Numeric -or $Block -match $KB) {
        Write-TSLog "DISM.exe block: Trobat KB en el bloc de sortida" -Level "DEBUG"

        if ($Block -match "Package Identity\s*:\s*(.+)" ) {
            $pkgId = $Matches[1].Trim()
            if (-not $Result.PackageName) { $Result.PackageName = $pkgId }
            $Result.IsInstalled = $true
            if (-not $Result.DetectMethod) { $Result.DetectMethod = "DISM.exe scan" }
            $Result.Details += "DISM.exe scan: $pkgId"
            Write-TSLog "DISM.exe: Trobat paquet $pkgId" -Level "INFO"
        }

        if ($Block -match "State\s*:\s*(.+)" ) {
            $state = $Matches[1].Trim()
            Write-TSLog "DISM.exe: Package state = $state" -Level "DEBUG"
            if ($state -match "Install Pending|Staged") {
                $Result.IsPending = $true
            }
        }

        if ($Block -match "Release Type\s*:\s*(.+)" ) {
            $relType = $Matches[1].Trim()
            Write-TSLog "DISM.exe: Release type = $relType" -Level "DEBUG"
        }

        if ($Block -match "Install Time\s*:\s*(.+)" ) {
            Write-TSLog "DISM.exe: Install time = $($Matches[1].Trim())" -Level "DEBUG"
        }
    }
}

# ============================================================================
# REGIÓ: Funcions de desinstal·lació
# ============================================================================

function Remove-UpdateViaWUSA {
    <#
    .SYNOPSIS
        Desinstal·la el KB usant wusa.exe
        Retorna $true si té èxit, $false en cas contrari.
    #>
    [CmdletBinding()]
    param([string]$KB)

    $kbNum = $KB -replace '[^0-9]', ''

    # Intent 1: wusa.exe amb /quiet /norestart
    $wusaArgs1 = @("/uninstall", "/kb:$kbNum", "/quiet", "/norestart")
    Write-TSLog "Intentant desinstal·lar via wusa.exe: wusa.exe $($wusaArgs1 -join ' ')" -Level "INFO"

    try {
        $proc = Start-Process -FilePath "wusa.exe" -ArgumentList $wusaArgs1 -Wait -NoNewWindow -PassThru
        $exitCode = $proc.ExitCode
        Write-TSLog "wusa.exe ha sortit amb codi: $exitCode" -Level "INFO"

        $knownCodes = @{
            0       = "Desinstal·lat correctament"
            3010    = "Desinstal·lat, cal reiniciar"
            87      = "ERROR_INVALID_PARAMETER — wusa.exe no suporta aquest tipus d'update; es farà fallback a DISM"
            2359302 = "Update no instal·lat"
        }

        switch ($exitCode) {
            0 {
                Write-TSLog "wusa.exe: KB $KB desinstal·lat correctament" -Level "SUCCESS"
                return $true
            }
            3010 {
                Write-TSLog "wusa.exe: KB $KB desinstal·lat. Cal reiniciar." -Level "SUCCESS"
                $script:RebootRequired = $true
                return $true
            }
            87 {
                Write-TSLog "wusa.exe: Exit 87 - L'update no es pot desinstal·lar via wusa (possiblement paquet acumulatiu o pending). Continuant amb DISM..." -Level "WARN"
                return $false
            }
            2359302 {
                Write-TSLog "wusa.exe: KB $KB no està instal·lat en aquest equip" -Level "INFO"
                return $false
            }
            default {
                Write-TSLog "wusa.exe: Error inesperat (exit code: $exitCode)" -Level "ERROR"
                return $false
            }
        }
    } catch {
        Write-TSLog "wusa.exe: Excepció - $_" -Level "ERROR"
        return $false
    }
}

function Remove-UpdateViaDISM {
    <#
    .SYNOPSIS
        Desinstal·la el KB usant DISM (PowerShell cmdlet o dism.exe)
        Retorna $true si té èxit, $false en cas contrari.
    #>
    [CmdletBinding()]
    param([string]$PackageName)

    if (-not $PackageName) {
        Write-TSLog "DISM: No s'ha proporcionat PackageName per a la desinstal·lació" -Level "ERROR"
        return $false
    }

    Write-TSLog "Intentant desinstal·lar via DISM: $PackageName" -Level "INFO"

    # Intent 1: Remove-WindowsPackage (PowerShell cmdlet)
    try {
        Write-TSLog "DISM: Usant Remove-WindowsPackage..." -Level "INFO"
        $result = Remove-WindowsPackage -Online -PackageName $PackageName -NoRestart -ErrorAction Stop

        Write-TSLog "DISM Remove-WindowsPackage: $($result | Out-String)" -Level "DEBUG"

        if ($result.RestartNeeded -eq $true) {
            $script:RebootRequired = $true
            Write-TSLog "DISM: Desinstal·lació completada. Cal reiniciar." -Level "SUCCESS"
        } else {
            Write-TSLog "DISM Remove-WindowsPackage: Desinstal·lació completada." -Level "SUCCESS"
        }
        return $true

    } catch {
        Write-TSLog "DISM Remove-WindowsPackage ha fallat: $_" -Level "WARN"
    }

    # Intent 2: dism.exe directe (fallback)
    try {
        $dismArgs = @(
            "/online",
            "/remove-package",
            "/packagename:$PackageName",
            "/quiet",
            "/norestart"
        )

        Write-TSLog "DISM: Usant dism.exe: dism.exe $($dismArgs -join ' ')" -Level "INFO"

        $dismProc = Start-Process -FilePath "dism.exe" `
            -ArgumentList $dismArgs `
            -Wait -NoNewWindow -PassThru

        $dismExitCode = $dismProc.ExitCode

        Write-TSLog "dism.exe ha sortit amb codi: $dismExitCode" -Level "INFO"

        if ($dismExitCode -eq 0) {
            Write-TSLog "dism.exe: Desinstal·lació completada correctament" -Level "SUCCESS"
            return $true
        } elseif ($dismExitCode -eq 3010) {
            $script:RebootRequired = $true
            Write-TSLog "dism.exe: Desinstal·lació completada. Cal reiniciar." -Level "SUCCESS"
            return $true
        } else {
            Write-TSLog "dism.exe: Error (exit code: $dismExitCode)" -Level "ERROR"
            return $false
        }
    } catch {
        Write-TSLog "dism.exe: Excepció - $_" -Level "ERROR"
        return $false
    }
}

# ============================================================================
# REGIÓ: Funció de verificació post-desinstal·lació
# ============================================================================

function Test-RebootPending {
    <#
    .SYNOPSIS
        Comprova si hi ha un reinici pendent al sistema.
        Retorna $true si cal reiniciar.
    #>
    [CmdletBinding()]
    param()

    $rebootPending = $false
    $reasons = @()

    # Comprovar CBS (Component Based Servicing)
    $cbsRebootKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
    if (Test-Path -Path $cbsRebootKey) {
        $rebootPending = $true
        $reasons += "CBS RebootPending key exists"
    }

    # Comprovar Windows Update
    $wuRebootKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    if (Test-Path -Path $wuRebootKey) {
        $rebootPending = $true
        $reasons += "Windows Update RebootRequired key exists"
    }

    # Comprovar pending.xml
    if (Test-Path -Path "$env:SystemRoot\WinSxS\pending.xml") {
        $rebootPending = $true
        $reasons += "pending.xml exists"
    }

    # Comprovar PendingFileRenameOperations
    $pfroKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"
    try {
        $pfro = Get-ItemProperty -Path $pfroKey -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
        if ($pfro -and $pfro.PendingFileRenameOperations -and $pfro.PendingFileRenameOperations.Count -gt 0) {
            $rebootPending = $true
            $reasons += "PendingFileRenameOperations has entries"
        }
    } catch {
        # Key or value might not exist
    }

    if ($reasons.Count -gt 0) {
        Write-TSLog "Reinici pendent detectat per: $($reasons -join ', ')" -Level "INFO"
    }

    return $rebootPending
}

# ============================================================================
# REGIÓ: Execució principal
# ============================================================================

Write-TSLog "========== INICI SCRIPT DESINSTAL·LACIÓ $KB ==========" -Level "INFO"
Write-TSLog "Versió script : $ScriptVersion" -Level "INFO"
Write-TSLog "Data execució : $($ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
Write-TSLog "Hostname      : $env:COMPUTERNAME" -Level "INFO"
Write-TSLog "Usuari        : $env:USERDOMAIN\$env:USERNAME" -Level "INFO"
Write-TSLog "Context       : $(if($IsSCCM){'SCCM'}elseif($IsIntune){'Intune'}else{'Local'})" -Level "INFO"
Write-TSLog "Log path      : $LogPath" -Level "INFO"

# Obtenir informació del SO
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    Write-TSLog "OS            : $($os.Caption) ($($os.Version)) Build $($os.BuildNumber)" -Level "INFO"
    Write-TSLog "Arquitectura  : $($os.OSArchitecture)" -Level "INFO"
    Write-TSLog "Últim reinici : $($os.LastBootUpTime)" -Level "INFO"
} catch {
    Write-TSLog "No s'ha pogut obtenir informació del SO: $_" -Level "WARN"
}

# Verificar privilegis d'administrador
if (-not (Test-IsAdministrator)) {
    Write-TSLog "ERROR: El script requereix privilegis d'administrador" -Level "ERROR"
    Write-TSLog "Executeu-lo com a administrador o mitjançant SCCM/Intune (SYSTEM)." -Level "ERROR"
    $ExitCode = 1
    Write-TSLog "========== FI SCRIPT (EXIT CODE: $ExitCode) ==========" -Level "ERROR"
    exit $ExitCode
}

Write-TSLog "Privilegis d'administrador: OK" -Level "INFO"

# ----------------------------------------------------------------------------
# FASE 1: Detecció
# ----------------------------------------------------------------------------
Write-TSLog "---------- FASE 1: DETECCIÓ ----------" -Level "INFO"

$updateStatus = Get-UpdateStatus -KB $KB

Write-TSLog "Resultat detecció -> IsInstalled=$($updateStatus.IsInstalled) | IsPending=$($updateStatus.IsPending) | DetectMethod=$($updateStatus.DetectMethod)" -Level "INFO"
Write-TSLog "PackageName       -> $($updateStatus.PackageName)" -Level "DEBUG"

if (-not $updateStatus.IsInstalled -and -not $updateStatus.IsPending) {
    Write-TSLog "KB $KB NO està instal·lat en aquest equip. No cal fer res." -Level "SUCCESS"
    Write-TSLog "========== FI SCRIPT (EXIT CODE: 0 - Ja conforme) ==========" -Level "SUCCESS"
    exit 0
}

# ----------------------------------------------------------------------------
# FASE 2: Desinstal·lació
# ----------------------------------------------------------------------------
Write-TSLog "---------- FASE 2: DESINSTAL·LACIÓ ----------" -Level "INFO"
$removalSuccess = $false

# Via 1: wusa.exe (mètode preferent per KB)
Write-TSLog "Intentant via primària: wusa.exe" -Level "INFO"
$removalSuccess = Remove-UpdateViaWUSA -KB $KB

# Via 2: DISM (fallback si wusa falla o no troba el paquet)
if (-not $removalSuccess) {
    Write-TSLog "Fallback a via secundària: DISM" -Level "INFO"
    if ($updateStatus.PackageName) {
        $removalSuccess = Remove-UpdateViaDISM -PackageName $updateStatus.PackageName
    } else {
        # Intentar trobar el nom del paquet amb pattern matching genèric
        Write-TSLog "DISM: Cercant paquet amb patrons alternatius..." -Level "INFO"
        $packageFound = $false

        # Sub-via 2a: pattern matching al PackageName
        try {
            $genericPkg = Get-WindowsPackage -Online -ErrorAction Stop |
                Where-Object { $_.PackageName -like "*$($KB_Numeric)*" -or
                               $_.PackageName -like "*$($KB)*" } |
                Select-Object -First 1

            if ($genericPkg) {
                Write-TSLog "DISM: Trobat paquet per patron: $($genericPkg.PackageName)" -Level "INFO"
                $removalSuccess = Remove-UpdateViaDISM -PackageName $genericPkg.PackageName
                $packageFound = $true
            }
        } catch {
            Write-TSLog "DISM cerca per patró ha fallat: $_" -Level "ERROR"
        }

        # Sub-via 2b: CBS InstallName (runtime search)
        if (-not $packageFound) {
            Write-TSLog "CBS InstallName runtime: Buscant clau amb InstallName=$KB..." -Level "INFO"
            try {
                $cbsRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
                $allCbsKeys = Get-ChildItem -Path $cbsRegPath -ErrorAction Stop
                foreach ($cbsKey in $allCbsKeys) {
                    $installName = (Get-ItemProperty -Path $cbsKey.PSPath -Name "InstallName" -ErrorAction SilentlyContinue).InstallName
                    if ($installName -eq $KB -or $installName -eq $KB_Numeric) {
                        Write-TSLog "CBS InstallName runtime: Trobat! $($cbsKey.PSChildName)" -Level "SUCCESS"
                        $removalSuccess = Remove-UpdateViaDISM -PackageName $cbsKey.PSChildName
                        $packageFound = $true
                        break
                    }
                }
            } catch {
                Write-TSLog "CBS InstallName runtime search error: $_" -Level "WARN"
            }
        }

        # Sub-via 2c: Correlació per InstallTime (últim recurs)
        if (-not $packageFound -and $updateStatus.InstalledOn) {
            Write-TSLog "InstallTime runtime: Buscant paquets propers a $($updateStatus.InstalledOn)..." -Level "INFO"
            try {
                $allPkgs = Get-WindowsPackage -Online -ErrorAction Stop
                $bestMatch = $null
                $bestDiff = 99999

                foreach ($pkg in $allPkgs) {
                    if ($pkg.InstallTime -and $pkg.ReleaseType -in @('Update', 'Security Update', 'SecurityUpdate', 'Hotfix', 'UpdateEx')) {
                        $diff = [Math]::Abs(($pkg.InstallTime - $updateStatus.InstalledOn).TotalMinutes)
                        $sameDay = ($pkg.InstallTime.Date -eq $updateStatus.InstalledOn.Date)
                        if (($diff -lt 1500 -or $sameDay) -and $diff -lt $bestDiff) {
                            $bestMatch = $pkg
                            $bestDiff = $diff
                        }
                    }
                }

                if ($bestMatch) {
                    Write-TSLog "InstallTime runtime: Trobat! $($bestMatch.PackageName) (diff=$([math]::Round($bestDiff,1)) min)" -Level "SUCCESS"
                    $removalSuccess = Remove-UpdateViaDISM -PackageName $bestMatch.PackageName
                    $packageFound = $true
                } else {
                    Write-TSLog "InstallTime runtime: No s'ha trobat cap paquet dins de +/-2h" -Level "WARN"
                }
            } catch {
                Write-TSLog "InstallTime runtime search error: $_" -Level "WARN"
            }
        }

        if (-not $packageFound) {
            Write-TSLog "DISM: No s'ha pogut trobar el paquet per cap via disponible" -Level "WARN"
        }
    }
}

# ----------------------------------------------------------------------------
# FASE 3: Verificació
# ----------------------------------------------------------------------------
Write-TSLog "---------- FASE 3: VERIFICACIÓ ----------" -Level "INFO"

if ($removalSuccess) {
    Write-TSLog "Desinstal·lació completada. Verificant estat final..." -Level "INFO"

    # Petita pausa perquè el sistema processi el canvi
    Start-Sleep -Seconds 3

    # Verificar que el KB ja no hi és
    $finalStatus = Get-UpdateStatus -KB $KB

    if ($finalStatus.IsInstalled) {
        Write-TSLog "ATENCIÓ: KB $KB encara apareix com a instal·lat després de la desinstal·lació." -Level "WARN"
        Write-TSLog "Possiblement calgui un reinici per completar l'eliminació." -Level "WARN"
        $script:RebootRequired = $true
    } else {
        Write-TSLog "Verificació OK: KB $KB ja no apareix com a instal·lat." -Level "SUCCESS"
    }

    # Comprovar si el mateix sistema reporta reinici pendent
    if (Test-RebootPending) {
        $script:RebootRequired = $true
    }

} else {
    Write-TSLog "ERROR: No s'ha pogut desinstal·lar KB $KB per cap via disponible." -Level "ERROR"

    # Verificar si l'update és potser permanent (no desinstal·lable)
    if ($updateStatus.IsInstalled) {
        Write-TSLog "L'actualització pot ser permanent o estar bloquejada." -Level "ERROR"
        Write-TSLog "Verifiqueu que KB $KB no sigui un Servicing Stack Update (SSU)." -Level "ERROR"
    }
}

# ----------------------------------------------------------------------------
# RESULTAT FINAL
# ----------------------------------------------------------------------------
Write-TSLog "---------- RESULTAT FINAL ----------" -Level "INFO"

if ($removalSuccess) {
    if ($RebootRequired) {
        $ExitCode = 3010
        Write-TSLog "KB $KB desinstal·lat. CAL REINICIAR l'equip per completar." -Level "SUCCESS"
    } else {
        $ExitCode = 0
        Write-TSLog "KB $KB desinstal·lat correctament. No cal reinici." -Level "SUCCESS"
    }
} else {
    if (-not $updateStatus.IsInstalled) {
        $ExitCode = 0
        Write-TSLog "KB $KB no estava instal·lat. Estat conforme." -Level "SUCCESS"
    } else {
        $ExitCode = 1
        Write-TSLog "ERROR: Fallada en la desinstal·lació de KB $KB" -Level "ERROR"
    }
}

$duration = [math]::Round(((Get-Date) - $ScriptStartTime).TotalSeconds, 1)
Write-TSLog "Temps total d'execució: $duration segons" -Level "INFO"
Write-TSLog "Codi de sortida     : $ExitCode" -Level "INFO"
Write-TSLog "========== FI SCRIPT (EXIT CODE: $ExitCode) ==========" -Level "INFO"

exit $ExitCode
