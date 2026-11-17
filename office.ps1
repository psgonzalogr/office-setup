# ==============================================================================
#  SafeTech - Instalador de Office
#  Instala Office 2024, Office 2021 o Microsoft 365 con la herramienta oficial de
#  Microsoft (ODT) y activa Office 2021/2024 por telefono de forma automatica con
#  el codigo de activacion SafeTech (getcid.safe-tech.cl).
#
#  Uso (PowerShell):
#    irm https://raw.githubusercontent.com/psgonzalogr/office-setup/main/office.ps1 | iex
#
#  Repositorio: https://github.com/psgonzalogr/office-setup
#  SafeTech SpA - https://safe-tech.cl
# ==============================================================================

$StConfig = @{
    Version     = '1.0.5'
    Repo        = 'psgonzalogr/office-setup'
    ScriptUrl   = 'https://raw.githubusercontent.com/psgonzalogr/office-setup/main/office.ps1'
    RawBase     = 'https://raw.githubusercontent.com/psgonzalogr/office-setup/main'
    ApiBase     = 'https://getcid.safe-tech.cl'
    OdtUrl      = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
    PortalUrl   = 'https://portal.office.com'
    HelpUrl     = 'https://ayuda.safe-tech.cl'
    WhatsAppUrl = 'https://wa.me/56951017168'
    WhatsApp    = '+56 9 5101 7168'
    Email       = 'contacto@safe-tech.cl'
    DataDir     = (Join-Path $env:ProgramData 'SafeTech\Office')
}
if ($env:SAFETECH_OFFICE_URL) {
    # Permite probar una rama distinta: $env:SAFETECH_OFFICE_URL = '<url raw de office.ps1>'
    $StConfig.ScriptUrl = $env:SAFETECH_OFFICE_URL
    $StConfig.RawBase = $env:SAFETECH_OFFICE_URL -replace '/[^/]+$', ''
}

$StMode = if ($env:SAFETECH_OFFICE_MODE) { $env:SAFETECH_OFFICE_MODE } else { 'app' }

# ------------------------------------------------------------------------------
#  Arranque: pedir permisos de administrador y abrir el instalador en una
#  ventana propia (Windows PowerShell 5.1, STA). La consola del cliente queda libre.
# ------------------------------------------------------------------------------
if ($StMode -eq 'app') {
    if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
        Write-Host 'Este instalador funciona solo en Windows.' -ForegroundColor Yellow
        return
    }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }

    $stIsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)

    $stPsOk = ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -ge 1)
    if ($env:SAFETECH_OFFICE_CHILD -eq '1' -and (-not $stIsAdmin -or -not $stPsOk -or [Environment]::OSVersion.Version.Major -lt 10)) {
        # Ya es la ventana relanzada y aun asi no cumple: avisar en vez de relanzar otra vez.
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show('Este instalador necesita Windows 10 u 11 con Windows PowerShell 5.1 y permisos de administrador.', 'SafeTech', 'OK', 'Warning') | Out-Null
        return
    }
    if ($env:SAFETECH_OFFICE_CHILD -ne '1' -or -not $stIsAdmin -or -not $stPsOk) {
        $stInner = @"
`$env:SAFETECH_OFFICE_CHILD = '1'
`$env:SAFETECH_OFFICE_URL = '$($StConfig.ScriptUrl)'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
try {
    `$r = Invoke-WebRequest -Uri ('$($StConfig.ScriptUrl)' + '?t=' + [DateTime]::UtcNow.Ticks) -UseBasicParsing
    `$t = [Text.Encoding]::UTF8.GetString(`$r.RawContentStream.ToArray()).TrimStart([char]0xFEFF)
    Invoke-Expression `$t
} catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(('No pudimos abrir el instalador de SafeTech. Revisa tu conexion a Internet e intenta de nuevo.' + [Environment]::NewLine + [Environment]::NewLine + `$_.Exception.Message), 'SafeTech', 'OK', 'Error') | Out-Null
}
"@
        $stEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($stInner))
        $stPs = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        Write-Host ''
        Write-Host '  SafeTech - Instalador de Office' -ForegroundColor Red
        Write-Host '  Windows te preguntara si permites que la aplicacion haga cambios. Responde "Si".' -ForegroundColor Gray
        try {
            Start-Process -FilePath $stPs -Verb RunAs -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden', '-EncodedCommand', $stEncoded) | Out-Null
            Write-Host '  Listo: el instalador se abrio en una ventana nueva. Ya puedes cerrar esta ventana.' -ForegroundColor Green
        } catch {
            Write-Host '  No se otorgaron los permisos de administrador.' -ForegroundColor Yellow
            Write-Host '  Vuelve a ejecutar el comando y responde "Si" cuando Windows lo pregunte.' -ForegroundColor Yellow
        }
        Write-Host ''
        return
    }
}

# ==============================================================================
#  MOTOR: deteccion, ODT, licencias (OSPP) y API de SafeTech.
#  Se carga en la interfaz y en cada tarea en segundo plano.
# ==============================================================================
$StEngine = {
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Get-StFirst {
    # Primer elemento de una lista, o $null. No usar "Select-Object -First" en el motor: en Windows
    # PowerShell 5.1, dentro de una tarea en segundo plano, detiene la tarea completa sin avisar.
    param($Items)
    foreach ($item in @($Items)) { return $item }
    return $null
}

function Write-StLog {
    param([string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $file = $null
    if ($sync -and $sync.ContainsKey('LogFile')) { $file = $sync['LogFile'] }
    if ($file) {
        try { [IO.File]::AppendAllText($file, $line + [Environment]::NewLine, [Text.Encoding]::UTF8) } catch { }
    }
}

function Set-StStep {
    param([string]$Id, [string]$State, [string]$Detail)
    if ($sync) {
        $sync.Steps[$Id] = $State
        if ($PSBoundParameters.ContainsKey('Detail')) { $sync.Detail = $Detail }
    }
    Write-StLog ("Paso {0}: {1} {2}" -f $Id, $State, $Detail)
}

function Get-StProduct {
    param([Parameter(Mandatory)][string]$Id)
    switch ($Id) {
        '2024' {
            return @{
                Id = '2024'; Name = 'Office 2024'; FullName = 'Office LTSC Profesional Plus 2024'
                OdtId = 'ProPlus2024Volume'; Channel = 'PerpetualVL2024'; Family = 'ProPlus2024'
                Year = '2024'; Short = '24'; NeedsKey = $true; Exclude = @('Lync')
            }
        }
        '2021' {
            return @{
                Id = '2021'; Name = 'Office 2021'; FullName = 'Office Profesional Plus 2021'
                OdtId = 'ProPlus2021Retail'; Channel = 'Current'; Family = 'ProPlus2021'
                Year = '2021'; Short = '21'; NeedsKey = $true; Exclude = @('Lync')
            }
        }
        '365' {
            return @{
                Id = '365'; Name = 'Microsoft 365'; FullName = 'Aplicaciones de Microsoft 365'
                OdtId = 'O365ProPlusRetail'; Channel = 'Current'; Family = 'O365ProPlus'
                Year = ''; Short = '365'; NeedsKey = $false; Exclude = @('Lync', 'Groove', 'OneDrive')
            }
        }
    }
    throw "Producto desconocido: $Id"
}

# ----- Formatos ----------------------------------------------------------------
function Format-StProductKey {
    param([string]$Raw)
    $clean = ($Raw -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    if ($clean.Length -gt 25) { $clean = $clean.Substring(0, 25) }
    $parts = for ($i = 0; $i -lt $clean.Length; $i += 5) { $clean.Substring($i, [Math]::Min(5, $clean.Length - $i)) }
    return (@($parts) -join '-')
}

function Test-StProductKey {
    param([string]$Key)
    return ($Key -cmatch '^[A-Z0-9]{5}(-[A-Z0-9]{5}){4}$')
}

function Get-StKeyLast5 {
    param([string]$Key)
    $clean = ($Key -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    if ($clean.Length -lt 5) { return $clean }
    return $clean.Substring($clean.Length - 5)
}

function ConvertTo-StAccessCode {
    param([string]$Raw)
    $clean = ($Raw -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    if ($clean.StartsWith('ST') -and $clean.Length -eq 8) { $clean = $clean.Substring(2) }
    if ($clean.Length -ne 6) { return $null }
    return "ST-$clean"
}

function Test-StAccessCode {
    param([string]$Code)
    return ($Code -cmatch '^ST-[A-Z0-9]{6}$')
}

function Hide-StSecret {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = $Text -replace '[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-([A-Za-z0-9]{5})', 'XXXXX-XXXXX-XXXXX-XXXXX-$1'
    $t = $t -replace '(?i)\bST-([A-Z0-9]{2})[A-Z0-9]{4}\b', 'ST-$1****'
    return $t
}

function Format-StDigitGroups {
    param([string]$Digits, [int]$Size = 6)
    $d = $Digits -replace '\D', ''
    $parts = for ($i = 0; $i -lt $d.Length; $i += $Size) { $d.Substring($i, [Math]::Min($Size, $d.Length - $i)) }
    return (@($parts) -join ' ')
}

# ----- Deteccion de Office -------------------------------------------------------
function Get-StFriendlyName {
    param([string]$Id)
    $year = ''
    if ($Id -match '(20\d\d)') { $year = ' ' + $matches[1] }
    switch -Regex ($Id) {
        '^O365HomePrem' { return 'Microsoft 365 Familia o Personal' }
        '^O365Business' { return 'Microsoft 365 Empresa' }
        '^O365' { return 'Microsoft 365' }
        '^Visio' { return "Visio$year".Trim() }
        '^Project' { return "Project$year".Trim() }
        '^(Word|Excel|PowerPoint|Outlook|Access|Publisher)' { return ($matches[1] + $year).Trim() }
    }
    if ($year) { return "Office$year" }
    return "Office ($Id)"
}

function Get-StOfficeState {
    $ids = New-Object System.Collections.Generic.List[string]
    $platform = $null
    $version = $null
    foreach ($path in @('HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration')) {
        try {
            $cfg = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
            if ($cfg.PSObject.Properties['ProductReleaseIds'] -and $cfg.ProductReleaseIds) {
                foreach ($id in ($cfg.ProductReleaseIds -split ',')) {
                    $t = $id.Trim()
                    if ($t -and -not $ids.Contains($t)) { $ids.Add($t) }
                }
            }
            if ($cfg.PSObject.Properties['Platform']) { $platform = $cfg.Platform }
            if ($cfg.PSObject.Properties['VersionToReport']) { $version = $cfg.VersionToReport }
        } catch { }
    }

    $msi = New-Object System.Collections.Generic.List[string]
    foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        foreach ($k in @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            if ($k.PSChildName -notmatch '0FF1CE\}$') { continue }
            $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
            if (-not $p -or -not $p.PSObject.Properties['DisplayName']) { continue }
            if ($p.PSObject.Properties['SystemComponent'] -and $p.SystemComponent -eq 1) { continue }
            $name = [string]$p.DisplayName
            if ($name -match '^Microsoft (Office|Visio|Project)' -and -not $msi.Contains($name)) { $msi.Add($name) }
        }
    }

    $friendly = New-Object System.Collections.Generic.List[string]
    foreach ($id in $ids) { $f = Get-StFriendlyName $id; if (-not $friendly.Contains($f)) { $friendly.Add($f) } }
    foreach ($n in $msi) { if (-not $friendly.Contains($n)) { $friendly.Add($n) } }

    return [pscustomobject]@{
        C2RIds   = @($ids)
        MsiNames = @($msi)
        Platform = $platform
        Version  = $version
        HasAny   = (($ids.Count + $msi.Count) -gt 0)
        Has365   = [bool](@($ids | Where-Object { $_ -match '^O365' }).Count)
        Friendly = @($friendly)
    }
}

function Test-StProductInstalled {
    param([string]$ProductId, $State)
    if (-not $State) { $State = Get-StOfficeState }
    $prod = Get-StProduct $ProductId
    if ($ProductId -eq '365') { return $State.Has365 }
    return [bool](@($State.C2RIds | Where-Object { $_ -like "$($prod.Family)*" }).Count)
}

# ----- Requisitos -----------------------------------------------------------------
function Get-StFreeSpaceGB {
    $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
    if (-not $disk) { return $null }
    return [Math]::Round($disk.FreeSpace / 1GB, 1)
}

function Test-StRequirements {
    $os = [Environment]::OSVersion.Version
    Write-StLog ("Windows {0}; 64 bits: {1}; PowerShell {2}" -f $os, [Environment]::Is64BitOperatingSystem, $PSVersionTable.PSVersion)
    if ($os.Major -lt 10) {
        throw 'Este Office necesita Windows 10 u 11. Tu computador tiene una versión anterior de Windows.'
    }
    $free = Get-StFreeSpaceGB
    Write-StLog "Espacio libre en ${env:SystemDrive}: $free GB"
    if ($null -ne $free -and $free -lt 10) {
        throw "Tu disco $env:SystemDrive tiene solo $free GB libres. Office necesita al menos 10 GB. Libera espacio y vuelve a intentarlo."
    }
    try {
        $req = [Net.WebRequest]::Create('https://officecdn.microsoft.com/pr/wsus/setup.exe')
        $req.Method = 'HEAD'
        $req.Timeout = 20000
        $resp = $req.GetResponse()
        $resp.Close()
    } catch {
        throw ("No pudimos conectarnos con los servidores de Microsoft. Revisa tu conexión a Internet; si usas VPN, proxy o un antivirus con filtro web, desactívalo un momento.`n`nDetalle: " + $_.Exception.Message)
    }
}

function New-StRestorePoint {
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    try {
        New-ItemProperty -Path $key -Name 'SystemRestorePointCreationFrequency' -Value 0 -PropertyType DWord -Force | Out-Null
        try {
            Checkpoint-Computer -Description 'SafeTech - antes de instalar Office' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
        } catch {
            Write-StLog "Primer intento de punto de restauracion fallo: $($_.Exception.Message). Se intentara activar Restaurar sistema."
            Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction Stop
            Checkpoint-Computer -Description 'SafeTech - antes de instalar Office' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
        }
        Write-StLog 'Punto de restauracion creado.'
        return $true
    } catch {
        Write-StLog "No se pudo crear el punto de restauracion: $($_.Exception.Message)"
        return $false
    } finally {
        Remove-ItemProperty -Path $key -Name 'SystemRestorePointCreationFrequency' -ErrorAction SilentlyContinue
    }
}

# ----- ODT -------------------------------------------------------------------------
function Get-StOdtSetup {
    param([Parameter(Mandatory)][string]$WorkDir)
    New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
    $dest = Join-Path $WorkDir 'setup.exe'
    $last = ''
    for ($i = 1; $i -le 3; $i++) {
        try {
            if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
            Invoke-WebRequest -Uri $StConfig.OdtUrl -OutFile $dest -UseBasicParsing -TimeoutSec 180
            $sig = Get-AuthenticodeSignature -LiteralPath $dest
            if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
                throw "La firma digital del instalador descargado no es válida ($($sig.Status))."
            }
            Write-StLog "ODT descargado y verificado: $dest"
            return $dest
        } catch {
            $last = $_.Exception.Message
            Write-StLog "Intento $i de descarga del ODT fallo: $last"
            Start-Sleep -Seconds (4 * $i)
        }
    }
    throw "No pudimos descargar el instalador oficial de Microsoft. Revisa tu conexión a Internet.`n`nDetalle: $last"
}

function New-StOdtXml {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$ProductId,
        [switch]$Remove,
        [string]$Edition = '64'
    )
    if ($Remove) {
        $xml = @'
<Configuration>
  <Remove All="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
'@
    } else {
        $prod = Get-StProduct $ProductId
        $excl = (@($prod.Exclude) | ForEach-Object { "      <ExcludeApp ID=""$_"" />" }) -join "`r`n"
        $xml = @"
<Configuration>
  <Add OfficeClientEdition="$Edition" Channel="$($prod.Channel)" AllowCdnFallback="TRUE">
    <Product ID="$($prod.OdtId)">
      <Language ID="es-es" />
$excl
    </Product>
  </Add>
  <RemoveMSI />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Updates Enabled="TRUE" Channel="$($prod.Channel)" />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
    }
    [IO.File]::WriteAllText($Path, $xml, (New-Object Text.UTF8Encoding($false)))
    return $Path
}

function Invoke-StOdt {
    param([string]$Setup, [string]$Xml, [string]$Label)
    Write-StLog "ODT: $Label ($Xml)"
    # Sin ventana de consola: la ventana de avance de Microsoft se sigue mostrando.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Setup
    $psi.Arguments = "/configure `"$Xml`""
    $psi.WorkingDirectory = (Split-Path -Parent $Setup)
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.WaitForExit()
    Write-StLog "ODT: $Label termino con codigo $($proc.ExitCode)"
    return $proc.ExitCode
}

function Get-StOdtMessage {
    param([int]$Code)
    switch ($Code) {
        17002 { return 'La instalación se canceló o se cerró la ventana de Microsoft antes de terminar.' }
        17004 { return 'Microsoft no encontró los archivos de instalación para este producto.' }
        30015 { return 'Se cortó la descarga de Office. Revisa tu conexión a Internet.' }
        30088 { return 'Microsoft no pudo descargar Office. Revisa tu conexión a Internet.' }
        30182 { return 'Se cortó la descarga de Office. Revisa tu conexión a Internet.' }
        30183 { return 'Se cortó la descarga de Office. Revisa tu conexión a Internet.' }
    }
    return 'La herramienta de Microsoft informó un error.'
}

function Stop-StOfficeApps {
    $names = 'WINWORD', 'EXCEL', 'POWERPNT', 'OUTLOOK', 'MSACCESS', 'MSPUB', 'ONENOTE', 'VISIO', 'WINPROJ', 'lync', 'MSOSYNC'
    foreach ($n in $names) {
        foreach ($p in @(Get-Process -Name $n -ErrorAction SilentlyContinue)) {
            Write-StLog "Cerrando $($p.ProcessName) (PID $($p.Id))"
            try { $p.Kill() } catch { }
        }
    }
}

function Start-StClickToRunService {
    $svc = Get-Service -Name 'ClickToRunSvc' -ErrorAction SilentlyContinue
    if (-not $svc) { return }
    try {
        if ((Get-CimInstance Win32_Service -Filter "Name='ClickToRunSvc'").StartMode -eq 'Disabled') {
            Set-Service -Name 'ClickToRunSvc' -StartupType Automatic
        }
        if ($svc.Status -ne 'Running') { Start-Service -Name 'ClickToRunSvc' }
    } catch {
        Write-StLog "No se pudo iniciar ClickToRunSvc: $($_.Exception.Message)"
    }
}

# ----- Licencias (OSPP.VBS) ---------------------------------------------------------
function Get-StOsppPath {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\OSPP.VBS",
        "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OSPP.VBS",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return $c }
    }
    return $null
}

function Invoke-StOspp {
    param([Parameter(Mandatory)][string[]]$Arguments, [string]$OsppPath)
    if (-not $OsppPath) { $OsppPath = Get-StOsppPath }
    if (-not $OsppPath) {
        throw 'No encontramos los componentes de licencia de Office (OSPP.VBS). Verifica que Office esté instalado.'
    }
    $cscript = Join-Path $env:SystemRoot 'System32\cscript.exe'
    $ErrorActionPreference = 'Continue'
    $out = & $cscript //Nologo $OsppPath @Arguments 2>&1 | ForEach-Object { "$_" }
    $ErrorActionPreference = 'Stop'
    $text = (@($out) -join "`n")
    Write-StLog ("OSPP {0}`n{1}" -f (Hide-StSecret ($Arguments -join ' ')), (Hide-StSecret $text))
    return $text
}

function Test-StOsppSuccess {
    param([string]$Text)
    return ($Text -match '(?i)successful|correctamente|exitos')
}

function Get-StErrorCode {
    param([string]$Text)
    $m = [regex]::Match($Text, '0x[0-9A-Fa-f]{8}')
    if ($m.Success) { return $m.Value.ToUpperInvariant().Replace('0X', '0x') }
    return $null
}

function Test-StLicensedStatus {
    param([string]$Status)
    if (-not $Status) { return $false }
    if ($Status -match '(?i)unlicensed|sin licencia|notification|grace') { return $false }
    return ($Status -match '(?i)licensed|con licencia')
}

function ConvertFrom-StOsppStatus {
    param([string]$Text)
    $items = New-Object System.Collections.Generic.List[object]
    $new = { @{ Name = $null; Description = $null; Status = $null; Last5 = $null; ErrorCode = $null; SkuId = $null } }
    $cur = & $new
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\s*-{5,}') {
            if ($cur.Name) { $items.Add([pscustomobject]$cur) }
            $cur = & $new
            continue
        }
        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)$') {
            $k = $matches[1].Trim().ToUpperInvariant()
            $v = $matches[2].Trim()
            if ($k -match '^(LICENSE NAME|NOMBRE DE LA LICENCIA)') { $cur.Name = $v }
            elseif ($k -match '^(LICENSE DESCRIPTION|DESCRIPCI)') { $cur.Description = $v }
            elseif ($k -match '^(LICENSE STATUS|ESTADO DE LA LICENCIA)') { $cur.Status = ($v -replace '-', '').Trim() }
            elseif ($k -match 'LAST 5|LTIMOS 5') { $cur.Last5 = $v.ToUpperInvariant() }
            elseif ($k -match '^(ERROR CODE|C.DIGO DE ERROR)') { $cur.ErrorCode = $v }
            elseif ($k -match '^SKU ID') { $cur.SkuId = $v }
        }
    }
    if ($cur.Name) { $items.Add([pscustomobject]$cur) }
    return $items.ToArray()
}

function ConvertFrom-StOsppIid {
    param([string]$Text)
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($m in [regex]::Matches($Text, '(?im)^\s*Installation ID for:\s*(?<name>.+):\s*(?<iid>\d{54,63})\s*$')) {
        $items.Add([pscustomobject]@{ Name = $m.Groups['name'].Value.Trim(); Iid = $m.Groups['iid'].Value })
    }
    if ($items.Count -eq 0) {
        foreach ($m in [regex]::Matches($Text, '(?<!\d)(\d{63}|\d{54})(?!\d)')) {
            $items.Add([pscustomobject]@{ Name = ''; Iid = $m.Value })
        }
    }
    return $items.ToArray()
}

function Get-StOsppMessage {
    param([string]$Code)
    switch ($Code) {
        '0xC004F069' { return 'La clave no corresponde a la versión de Office instalada. Revisa que sea la clave de Office correcta.' }
        '0xC004F050' { return 'Microsoft indica que la clave de producto no es válida. Revisa que esté bien escrita.' }
        '0xC004F04D' { return 'Microsoft no aceptó el ID de confirmación. Puede que el ID de instalación haya cambiado.' }
        '0xC004F035' { return 'Esta clave no se puede usar en este equipo.' }
        '0xC004C060' { return 'Microsoft informa que esta clave está bloqueada. Escríbenos para revisar tu licencia.' }
        '0xC004C003' { return 'Microsoft informa que esta clave está bloqueada o alcanzó su límite. Escríbenos para revisar tu licencia.' }
        '0xC004C004' { return 'Microsoft informa que esta clave no es válida. Escríbenos para revisar tu licencia.' }
        '0xC004C008' { return 'La clave alcanzó su límite de activaciones en línea; se activará por teléfono.' }
        '0xC004C020' { return 'La clave alcanzó su límite de activaciones en línea; se activará por teléfono.' }
    }
    if ($Code) { return "Windows informó el error $Code." }
    return 'No recibimos una respuesta clara del servicio de licencias de Windows.'
}

function Get-StMyLicense {
    param([string]$Key)
    $last5 = Get-StKeyLast5 $Key
    $status = ConvertFrom-StOsppStatus (Invoke-StOspp -Arguments @('/dstatus'))
    return (Get-StFirst @($status | Where-Object { $_.Last5 -eq $last5 }))
}

function Install-StLicenseFiles {
    param([string]$Family)
    $ospp = Get-StOsppPath
    if (-not $ospp) { return 0 }
    # Click-to-Run: OSPP.VBS en ...\Microsoft Office\Office16 y licencias en ...\Microsoft Office\root\Licenses16
    $officeDir = Split-Path -Parent $ospp
    $baseDir = Split-Path -Parent $officeDir
    $licDir = $null
    foreach ($c in @((Join-Path $baseDir 'root\Licenses16'), (Join-Path $baseDir 'Licenses16'),
                     (Join-Path (Split-Path -Parent $baseDir) 'root\Licenses16'))) {
        if (Test-Path -LiteralPath $c) { $licDir = $c; break }
    }
    if (-not $licDir) { Write-StLog "No se encontró la carpeta Licenses16 junto a $ospp"; return 0 }
    $files = @(Get-ChildItem -LiteralPath $licDir -Filter "$Family*.xrm-ms" -File -ErrorAction SilentlyContinue)
    foreach ($f in $files) {
        [void](Invoke-StOspp -Arguments @("/inslic:$($f.FullName)") -OsppPath $ospp)
    }
    Write-StLog "Licencias instaladas para ${Family}: $($files.Count)"
    return $files.Count
}

function Clear-StOfficeLicenses {
    $ospp = Get-StOsppPath
    if (-not $ospp) { Write-StLog 'Sin OSPP.VBS: no hay licencias que limpiar.'; return }
    $text = Invoke-StOspp -Arguments @('/dstatusall') -OsppPath $ospp
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($lic in (ConvertFrom-StOsppStatus $text)) {
        if ($lic.Last5 -and $lic.Last5 -match '^[A-Z0-9]{5}$' -and -not $keys.Contains($lic.Last5)) { $keys.Add($lic.Last5) }
    }
    foreach ($k in $keys) {
        Write-StLog "Retirando clave terminada en $k"
        [void](Invoke-StOspp -Arguments @("/unpkey:$k") -OsppPath $ospp)
    }
    [void](Invoke-StOspp -Arguments @('/remhst') -OsppPath $ospp)
}

# ----- Instalacion completa (segundo plano) ------------------------------------------------
function Invoke-StInstall {
    param([Parameter(Mandatory)][string]$ProductId, [Parameter(Mandatory)][string]$WorkDir)
    $prod = Get-StProduct $ProductId
    $edition = if ([Environment]::Is64BitOperatingSystem) { '64' } else { '32' }
    Write-StLog "Inicio de instalacion de $($prod.FullName) ($edition bits) - instalador v$($StConfig.Version)"

    Set-StStep 'check' 'running' 'Revisando Windows, el espacio libre y la conexión con Microsoft...'
    Test-StRequirements
    Set-StStep 'check' 'done'

    Set-StStep 'restore' 'running' 'Creando un punto de restauración de Windows. Puede tardar unos minutos...'
    if (New-StRestorePoint) { Set-StStep 'restore' 'done' } else { Set-StStep 'restore' 'skip' }

    Set-StStep 'download' 'running' 'Descargando la herramienta oficial de Microsoft...'
    $setup = Get-StOdtSetup -WorkDir $WorkDir
    $removeXml = New-StOdtXml -Path (Join-Path $WorkDir 'quitar-office.xml') -Remove
    $installXml = New-StOdtXml -Path (Join-Path $WorkDir "instalar-$($prod.Id).xml") -ProductId $ProductId -Edition $edition
    Set-StStep 'download' 'done'

    $state = Get-StOfficeState
    $foundText = (@($state.Friendly) -join ', ')
    if (-not $foundText) { $foundText = 'ninguno' }
    Write-StLog "Office encontrado: $foundText"
    Set-StStep 'clean' 'running' 'Cerrando Word, Excel y los demás programas de Office...'
    Stop-StOfficeApps
    if ($state.HasAny) {
        Set-StStep 'clean' 'running' 'Retirando claves y configuraciones de activación anteriores...'
        Clear-StOfficeLicenses
    }
    Set-StStep 'clean' 'done'

    if ($state.C2RIds.Count -gt 0) {
        Set-StStep 'remove' 'running' ("Desinstalando {0}. No apagues el computador..." -f (@($state.Friendly) -join ', '))
        Start-StClickToRunService
        $code = Invoke-StOdt -Setup $setup -Xml $removeXml -Label 'Desinstalacion de Office anterior'
        if ($code -ne 0) {
            Start-Sleep -Seconds 8
            Start-StClickToRunService
            $code = Invoke-StOdt -Setup $setup -Xml $removeXml -Label 'Segundo intento de desinstalacion'
        }
        Start-Sleep -Seconds 10
        $after = Get-StOfficeState
        if ($after.C2RIds.Count -gt 0) {
            throw ("No pudimos desinstalar por completo la versión anterior de Office ({0}). Reinicia el computador y vuelve a abrir el instalador." -f (@($after.Friendly) -join ', '))
        }
        Stop-StOfficeApps
        Clear-StOfficeLicenses
        Set-StStep 'remove' 'done'
    } elseif ($state.MsiNames.Count -gt 0) {
        Set-StStep 'remove' 'done' 'La versión anterior se quitará durante la instalación.'
    } else {
        Set-StStep 'remove' 'skip'
    }

    Set-StStep 'install' 'running' "Instalando $($prod.Name). Se abrió la ventana de Microsoft con el avance; déjala trabajar, se cierra sola."
    Start-StClickToRunService
    $code = Invoke-StOdt -Setup $setup -Xml $installXml -Label "Instalacion de $($prod.FullName)"
    if ($code -ne 0) {
        throw ("{0} (código {1})`n`nRevisa tu conexión a Internet y vuelve a intentarlo. Si se repite, crea el ZIP de diagnóstico y envíalo a SafeTech." -f (Get-StOdtMessage $code), $code)
    }
    Set-StStep 'install' 'done'

    Set-StStep 'verify' 'running' 'Comprobando que Office quedó instalado...'
    $ok = $false
    for ($i = 0; $i -lt 18; $i++) {
        if (Test-StProductInstalled -ProductId $ProductId) { $ok = $true; break }
        Start-Sleep -Seconds 5
    }
    if (-not $ok) {
        throw "La herramienta de Microsoft terminó, pero no encontramos $($prod.Name) instalado. Reinicia el computador y vuelve a abrir el instalador."
    }
    Set-StStep 'verify' 'done' "$($prod.Name) quedó instalado."
    Write-StLog 'Instalacion verificada.'
    return [pscustomobject]@{ Ok = $true; Edition = $edition }
}

# ----- Activacion telefonica asistida ----------------------------------------------------
function Invoke-StActivateKey {
    param([Parameter(Mandatory)][string]$ProductId, [Parameter(Mandatory)][string]$Key)
    $prod = Get-StProduct $ProductId
    $mine = Get-StMyLicense -Key $Key
    if ($mine -and (Test-StLicensedStatus $mine.Status)) {
        Write-StLog 'La clave ya estaba activada.'
        return [pscustomobject]@{ Result = 'licensed'; Message = ''; ErrorCode = $null; LicenseName = $mine.Name }
    }

    if (-not $mine) {
        Set-StStep 'act_key' 'running' 'Registrando tu clave de producto en Office...'
        $out = Invoke-StOspp -Arguments @("/inpkey:$Key")
        if (-not (Test-StOsppSuccess $out)) {
            $code = Get-StErrorCode $out
            if ($code -in @('0xC004F069', '0xC004F050')) {
                Set-StStep 'act_key' 'running' 'Preparando las licencias de Office para tu clave...'
                [void](Install-StLicenseFiles -Family $prod.Family)
                $out = Invoke-StOspp -Arguments @("/inpkey:$Key")
            }
        }
        if (-not (Test-StOsppSuccess $out)) {
            $code = Get-StErrorCode $out
            return [pscustomobject]@{ Result = 'error'; Message = (Get-StOsppMessage $code); ErrorCode = $code; LicenseName = $null }
        }
        $mine = Get-StMyLicense -Key $Key
    }

    Set-StStep 'act_key' 'running' 'Probando la activación en línea...'
    $out = Invoke-StOspp -Arguments @('/act')
    $mine2 = Get-StMyLicense -Key $Key
    if ($mine2 -and (Test-StLicensedStatus $mine2.Status)) {
        return [pscustomobject]@{ Result = 'licensed'; Message = ''; ErrorCode = $null; LicenseName = $mine2.Name }
    }
    $code = Get-StErrorCode $out
    if ($code -in @('0xC004C060', '0xC004C003', '0xC004C004')) {
        return [pscustomobject]@{ Result = 'error'; Message = (Get-StOsppMessage $code); ErrorCode = $code; LicenseName = $null }
    }
    $name = $null
    if ($mine) { $name = $mine.Name }
    return [pscustomobject]@{ Result = 'phone'; Message = ''; ErrorCode = $code; LicenseName = $name }
}

function Get-StInstallationId {
    param([Parameter(Mandatory)][string]$ProductId, [Parameter(Mandatory)][string]$Key)
    $prod = Get-StProduct $ProductId
    $mine = Get-StMyLicense -Key $Key
    $entries = @(ConvertFrom-StOsppIid (Invoke-StOspp -Arguments @('/dinstid')))
    if ($entries.Count -eq 0) {
        throw 'Office no entregó el ID de instalación. Cierra Word y Excel y vuelve a intentarlo.'
    }
    $pick = $null
    if ($mine -and $mine.Name) { $pick = Get-StFirst @($entries | Where-Object { $_.Name -eq $mine.Name }) }
    if (-not $pick) {
        $yy = $prod.Year.Substring(2)
        $pick = Get-StFirst @($entries | Where-Object { $_.Name -match "($($prod.Year)|Office $yy)" -and $_.Name -notmatch 'KMS' })
    }
    if (-not $pick) { $pick = Get-StFirst $entries }
    Write-StLog "ID de instalacion obtenido para: $($pick.Name)"
    return $pick.Iid
}

function Request-StCid {
    param([Parameter(Mandatory)][string]$Code, [Parameter(Mandatory)][string]$Iid)
    $body = [Text.Encoding]::UTF8.GetBytes((@{ code = $Code; iid = $Iid } | ConvertTo-Json -Compress))
    $last = ''
    for ($i = 1; $i -le 3; $i++) {
        $resp = $null
        try {
            $resp = Invoke-RestMethod -Uri "$($StConfig.ApiBase)/generate" -Method Post -Body $body `
                -ContentType 'application/json; charset=utf-8' -UseBasicParsing -TimeoutSec 120
        } catch {
            $webResp = $null
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                try { $resp = $_.ErrorDetails.Message | ConvertFrom-Json } catch { $resp = $null }
            }
            if (-not $resp -and $_.Exception.PSObject.Properties['Response']) { $webResp = $_.Exception.Response }
            if ($webResp) {
                try {
                    $reader = New-Object IO.StreamReader($webResp.GetResponseStream(), [Text.Encoding]::UTF8)
                    $resp = $reader.ReadToEnd() | ConvertFrom-Json
                    $reader.Close()
                } catch { $resp = $null }
            }
            if (-not $resp) {
                $last = $_.Exception.Message
                Write-StLog "Intento $i de conexion con getcid.safe-tech.cl fallo: $last"
                # Solo se reintenta si la solicitud no llego al servidor; asi nunca se cobra dos veces el mismo crédito.
                $status = ''
                if ($_.Exception.PSObject.Properties['Status']) { $status = [string]$_.Exception.Status }
                if ($status -notin @('ConnectFailure', 'NameResolutionFailure', 'ProxyNameResolutionFailure', 'SecureChannelFailure', 'TrustFailure')) { break }
                Start-Sleep -Seconds (3 * $i)
                continue
            }
        }
        if ($resp.success -and $resp.cid) {
            $cid = ([string]$resp.cid) -replace '\D', ''
            Write-StLog "ID de confirmacion recibido ($($cid.Length) digitos)."
            return [pscustomobject]@{ Ok = $true; Cid = $cid; Kind = ''; Message = '' }
        }
        $msg = 'El servicio de activación no pudo generar el ID de confirmación.'
        if ($resp.PSObject.Properties['error'] -and $resp.error) { $msg = [string]$resp.error }
        $kind = 'other'
        if ($msg -match '(?i)c.digo|cr.dito|credito|saldo|agotad|expir|usad') { $kind = 'code' }
        Write-StLog "getcid.safe-tech.cl respondio: $msg"
        return [pscustomobject]@{ Ok = $false; Cid = $null; Kind = $kind; Message = $msg }
    }
    return [pscustomobject]@{ Ok = $false; Cid = $null; Kind = 'network'; Message = "No pudimos conectarnos con el servicio de activación de SafeTech. Detalle: $last" }
}

function Test-StCodeOnline {
    param([Parameter(Mandatory)][string]$Code)
    try {
        $r = Invoke-RestMethod -Uri ("$($StConfig.ApiBase)/check-code?code=" + [Uri]::EscapeDataString($Code)) -UseBasicParsing -TimeoutSec 20
    } catch {
        # Si el servidor responde con error HTTP y un JSON {valid:false}, se usa esa respuesta.
        $r = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { try { $r = $_.ErrorDetails.Message | ConvertFrom-Json } catch { $r = $null } }
        if (-not $r -or -not $r.PSObject.Properties['valid']) { throw }
    }
    $remaining = $null
    if ($r.PSObject.Properties['remaining']) { $remaining = $r.remaining }
    return [pscustomobject]@{ Valid = [bool]$r.valid; Remaining = $remaining }
}

function Confirm-StActivation {
    param([Parameter(Mandatory)][string]$ProductId, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Cid)
    $digits = $Cid -replace '\D', ''
    if ($digits.Length -ne 48) {
        return [pscustomobject]@{ Ok = $false; Message = 'El ID de confirmación debe tener 48 números.'; ErrorCode = $null }
    }
    $out = Invoke-StOspp -Arguments @("/actcid:$digits")
    $mine = Get-StMyLicense -Key $Key
    if (($mine -and (Test-StLicensedStatus $mine.Status)) -or (Test-StOsppSuccess $out)) {
        Write-StLog 'Activacion telefonica confirmada.'
        return [pscustomobject]@{ Ok = $true; Message = ''; ErrorCode = $null }
    }
    $code = Get-StErrorCode $out
    return [pscustomobject]@{ Ok = $false; Message = (Get-StOsppMessage $code); ErrorCode = $code }
}

# ----- Microsoft 365 -----------------------------------------------------------------------
function Test-StM365License {
    $dir = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\Licenses'
    $files = @(Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue)
    return ($files.Count -gt 0)
}

function Get-StSignInIssues {
    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($name in @('TokenBroker', 'wlidsvc')) {
        $svc = Get-CimInstance Win32_Service -Filter "Name='$name'" -ErrorAction SilentlyContinue
        if (-not $svc) {
            if ($name -eq 'TokenBroker') { $issues.Add('Falta el servicio de Windows "Administrador de cuentas web" (TokenBroker).') }
        } elseif ($svc.StartMode -eq 'Disabled') {
            $issues.Add("El servicio de Windows $name está deshabilitado.")
        }
    }
    foreach ($pkg in @('Microsoft.AAD.BrokerPlugin', 'Microsoft.Windows.CloudExperienceHost')) {
        if (-not (Get-AppxPackage -Name $pkg -ErrorAction SilentlyContinue)) {
            $issues.Add("Falta el componente de Windows $pkg.")
        }
    }
    return $issues.ToArray()
}

# ----- Diagnostico ----------------------------------------------------------------------------
function New-StDiagnosticZip {
    param([Parameter(Mandatory)][string]$LogDir)
    $info = New-Object System.Collections.Generic.List[string]
    $info.Add("Instalador SafeTech v$($StConfig.Version)")
    $info.Add("Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $info.Add("Windows: $([Environment]::OSVersion.VersionString); 64 bits: $([Environment]::Is64BitOperatingSystem)")
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $info.Add("Edicion: $($os.Caption) $($os.Version)")
    } catch { }
    $info.Add("Espacio libre: $(Get-StFreeSpaceGB) GB")
    try {
        $st = Get-StOfficeState
        $info.Add("Office (Click-to-Run): $(@($st.C2RIds) -join ', ')")
        $info.Add("Office (MSI): $(@($st.MsiNames) -join ', ')")
        $info.Add("Plataforma: $($st.Platform)  Version: $($st.Version)")
    } catch { }
    try {
        if (Get-StOsppPath) { $info.Add(''); $info.Add('Estado de licencias:'); $info.Add((Invoke-StOspp -Arguments @('/dstatus'))) }
    } catch { }
    [IO.File]::WriteAllText((Join-Path $LogDir 'informacion-del-equipo.txt'), (Hide-StSecret ($info -join "`r`n")), [Text.Encoding]::UTF8)

    $since = (Get-Date).AddHours(-12)
    $odtDir = Join-Path $LogDir 'registros-microsoft'
    foreach ($dir in @($env:TEMP, (Join-Path $env:SystemRoot 'Temp'))) {
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter "$env:COMPUTERNAME-*.log" -File -ErrorAction SilentlyContinue)) {
            if ($f.LastWriteTime -gt $since) {
                New-Item -ItemType Directory -Force -Path $odtDir | Out-Null
                Copy-Item -LiteralPath $f.FullName -Destination $odtDir -Force -ErrorAction SilentlyContinue
            }
        }
    }
    $desktop = [Environment]::GetFolderPath('Desktop')
    $zip = Join-Path $desktop ("SafeTech-Diagnostico-Office-{0}.zip" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($LogDir, $zip)
    return $zip
}

function Save-StAssets {
    param([string[]]$Paths, [string]$CacheDir, [int]$TimeoutSec = 30)
    New-Item -ItemType Directory -Force -Path $CacheDir | Out-Null
    foreach ($p in $Paths) {
        $dest = Join-Path $CacheDir ($p -replace '[\\/]', '_')
        if (Test-Path -LiteralPath $dest) { continue }
        try {
            Invoke-WebRequest -Uri "$($StConfig.RawBase)/$p" -OutFile "$dest.part" -UseBasicParsing -TimeoutSec $TimeoutSec
            Move-Item -LiteralPath "$dest.part" -Destination $dest -Force
        } catch { Remove-Item -LiteralPath "$dest.part" -ErrorAction SilentlyContinue }
    }
}
}

# Carga el motor en esta sesion.
. $StEngine
$ErrorActionPreference = 'Continue'

if ($StMode -eq 'lib') { return }

# ==============================================================================
#  INTERFAZ (WPF)
# ==============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

$St = @{
    EngineText  = $StEngine.ToString()
    UI          = @{}
    Jobs        = New-Object System.Collections.ArrayList
    Sync        = [hashtable]::Synchronized(@{ Steps = [hashtable]::Synchronized(@{}); Detail = ''; LogFile = $null })
    StepRows    = @{}
    StepOrder   = @()
    Pages       = @('Home', 'Data', 'Confirm', 'Progress', 'Done', 'Error', 'Tutorial')
    Product     = $null
    Mode        = 'install'
    Busy        = $false
    InstallDone = $false
    DialogOpen  = $false
    NoAnim      = $false
    StartTime   = $null
    Act         = @{}
    Tut         = @{ Index = 0; Started = $false; Auth = $null }
    CodeState   = 'empty'
    KeyFormatting = $false
    CacheDir    = (Join-Path $env:TEMP 'SafeTechOffice\cache')
    LogDir      = $null
}
$sync = $St.Sync
if (-not [System.Windows.SystemParameters]::ClientAreaAnimation) { $St.NoAnim = $true }

# ----- Estilos compartidos (ventana principal y ventanas emergentes) -------------------
$StResXaml = @'
<SolidColorBrush x:Key="Bg" Color="#09111F"/>
<SolidColorBrush x:Key="Surface" Color="#111C2F"/>
<SolidColorBrush x:Key="Surface2" Color="#17243A"/>
<SolidColorBrush x:Key="Surface3" Color="#20304A"/>
<SolidColorBrush x:Key="Line" Color="#2A3C57"/>
<SolidColorBrush x:Key="LineStrong" Color="#405776"/>
<SolidColorBrush x:Key="Text" Color="#F8FAFC"/>
<SolidColorBrush x:Key="TextSoft" Color="#DBE5F1"/>
<SolidColorBrush x:Key="Muted" Color="#AEBDD0"/>
<SolidColorBrush x:Key="Red" Color="#DC2626"/>
<SolidColorBrush x:Key="RedText" Color="#FCA5A5"/>
<LinearGradientBrush x:Key="RedGrad" StartPoint="0,0" EndPoint="1,1">
  <GradientStop Color="#DC2626" Offset="0"/><GradientStop Color="#991B1B" Offset="1"/>
</LinearGradientBrush>
<LinearGradientBrush x:Key="RedGradHover" StartPoint="0,0" EndPoint="1,1">
  <GradientStop Color="#EF4444" Offset="0"/><GradientStop Color="#B91C1C" Offset="1"/>
</LinearGradientBrush>
<LinearGradientBrush x:Key="TopLine" StartPoint="0,0" EndPoint="1,0">
  <GradientStop Color="#00DC2626" Offset="0"/><GradientStop Color="#DC2626" Offset="0.5"/><GradientStop Color="#00DC2626" Offset="1"/>
</LinearGradientBrush>
<FontFamily x:Key="Icons">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>

<Style x:Key="Kicker" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#FCA5A5"/><Setter Property="FontSize" Value="11.5"/>
  <Setter Property="FontWeight" Value="Black"/><Setter Property="Margin" Value="0,0,0,8"/>
</Style>
<Style x:Key="H1" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="32"/>
  <Setter Property="FontWeight" Value="Bold"/><Setter Property="TextWrapping" Value="Wrap"/>
  <Setter Property="FontFamily" Value="Segoe UI Variable Display, Segoe UI"/>
</Style>
<Style x:Key="H2" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="19"/>
  <Setter Property="FontWeight" Value="Bold"/><Setter Property="TextWrapping" Value="Wrap"/>
</Style>
<Style x:Key="Lead" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#AEBDD0"/><Setter Property="FontSize" Value="15.5"/>
  <Setter Property="TextWrapping" Value="Wrap"/><Setter Property="LineHeight" Value="24"/>
  <Setter Property="Margin" Value="0,10,0,0"/>
</Style>
<Style x:Key="Body" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#DBE5F1"/><Setter Property="FontSize" Value="14.5"/>
  <Setter Property="TextWrapping" Value="Wrap"/><Setter Property="LineHeight" Value="22"/>
</Style>
<Style x:Key="Small" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#AEBDD0"/><Setter Property="FontSize" Value="12.5"/>
  <Setter Property="TextWrapping" Value="Wrap"/>
</Style>
<Style x:Key="Label" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="14"/>
  <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Margin" Value="0,0,0,8"/>
</Style>
<Style x:Key="Panel" TargetType="Border">
  <Setter Property="Background" Value="#111C2F"/><Setter Property="BorderBrush" Value="#2A3C57"/>
  <Setter Property="BorderThickness" Value="1"/><Setter Property="CornerRadius" Value="20"/>
  <Setter Property="Padding" Value="22"/>
</Style>

<Style x:Key="BtnPrimary" TargetType="Button">
  <Setter Property="Foreground" Value="White"/><Setter Property="FontWeight" Value="Bold"/>
  <Setter Property="FontSize" Value="14"/><Setter Property="MinHeight" Value="46"/>
  <Setter Property="Padding" Value="24,10"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="RenderTransformOrigin" Value="0.5,0.5"/>
  <Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="1" ScaleY="1"/></Setter.Value></Setter>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="12" Padding="{TemplateBinding Padding}" Background="{StaticResource RedGrad}" BorderBrush="#00FFFFFF" BorderThickness="2">
        <Border.Effect><DropShadowEffect Color="#DC2626" BlurRadius="20" ShadowDepth="5" Direction="270" Opacity="0.32"/></Border.Effect>
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.Foreground="White"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource RedGradHover}"/></Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="0.97" ScaleY="0.97"/></Setter.Value></Setter></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#99FCA5A5"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.38"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnSoft" TargetType="Button">
  <Setter Property="Foreground" Value="#DBE5F1"/><Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="14"/><Setter Property="MinHeight" Value="46"/>
  <Setter Property="Padding" Value="20,10"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="RenderTransformOrigin" Value="0.5,0.5"/>
  <Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="1" ScaleY="1"/></Setter.Value></Setter>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="12" Padding="{TemplateBinding Padding}" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="1.5">
        <ContentPresenter x:Name="Cp" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="Bd" Property="BorderBrush" Value="#405776"/><Setter TargetName="Bd" Property="Background" Value="#20304A"/>
          <Setter Property="Foreground" Value="#FCA5A5"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="0.97" ScaleY="0.97"/></Setter.Value></Setter></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#99FCA5A5"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.38"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnLink" TargetType="Button">
  <Setter Property="Foreground" Value="#FCA5A5"/><Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="13.5"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Padding" Value="8,6"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="8" Padding="{TemplateBinding Padding}" Background="#00000000">
        <ContentPresenter VerticalAlignment="Center"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#1FDC2626"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="Background" Value="#1FDC2626"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnChrome" TargetType="Button">
  <Setter Property="Foreground" Value="#AEBDD0"/><Setter Property="Width" Value="46"/><Setter Property="Height" Value="34"/>
  <Setter Property="FontFamily" Value="{StaticResource Icons}"/><Setter Property="FontSize" Value="10"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="8" Background="#00000000"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#20304A"/><Setter Property="Foreground" Value="#F8FAFC"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="Card" TargetType="Button">
  <Setter Property="Cursor" Value="Hand"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="HorizontalContentAlignment" Value="Stretch"/><Setter Property="VerticalContentAlignment" Value="Stretch"/>
  <Setter Property="RenderTransform"><Setter.Value><TranslateTransform Y="0"/></Setter.Value></Setter>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="20" Background="#111C2F" BorderBrush="#2A3C57" BorderThickness="1.5">
        <Border.Effect><DropShadowEffect Color="#000000" BlurRadius="24" ShadowDepth="6" Direction="270" Opacity="0.25"/></Border.Effect>
        <Grid>
          <Border Height="1.5" VerticalAlignment="Top" Margin="24,0" Background="{StaticResource TopLine}"/>
          <ContentPresenter Margin="24,24,24,22"/>
        </Grid>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="Bd" Property="BorderBrush" Value="#B4DC2626"/><Setter TargetName="Bd" Property="Background" Value="#152238"/>
          <Setter Property="RenderTransform"><Setter.Value><TranslateTransform Y="-3"/></Setter.Value></Setter>
        </Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#DC2626"/></Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter Property="RenderTransform"><Setter.Value><TranslateTransform Y="0"/></Setter.Value></Setter></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="Field" TargetType="TextBox">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="CaretBrush" Value="#F8FAFC"/>
  <Setter Property="SelectionBrush" Value="#DC2626"/><Setter Property="FontSize" Value="18"/>
  <Setter Property="FontFamily" Value="Consolas, Segoe UI"/><Setter Property="Height" Value="54"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="TextBox">
      <Border x:Name="Bd" CornerRadius="12" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="1.5">
        <Grid Margin="16,0">
          <TextBlock x:Name="Ph" Text="{TemplateBinding Tag}" Foreground="#5D7391" VerticalAlignment="Center" Visibility="Collapsed" FontFamily="{TemplateBinding FontFamily}" FontSize="{TemplateBinding FontSize}"/>
          <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
        </Grid>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="Text" Value=""><Setter TargetName="Ph" Property="Visibility" Value="Visible"/></Trigger>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#405776"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#DC2626"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="Check" TargetType="CheckBox">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="14.5"/>
  <Setter Property="Cursor" Value="Hand"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="CheckBox">
      <Border x:Name="Row" CornerRadius="14" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="1.5" Padding="16,14">
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border x:Name="Box" Width="24" Height="24" CornerRadius="7" BorderBrush="#405776" BorderThickness="2" Background="#00000000" VerticalAlignment="Top">
            <TextBlock x:Name="Mark" Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="13" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
          </Border>
          <ContentPresenter Grid.Column="1" Margin="14,1,0,0" VerticalAlignment="Center"/>
        </Grid>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsChecked" Value="True">
          <Setter TargetName="Box" Property="Background" Value="#DC2626"/><Setter TargetName="Box" Property="BorderBrush" Value="#DC2626"/>
          <Setter TargetName="Mark" Property="Visibility" Value="Visible"/><Setter TargetName="Row" Property="BorderBrush" Value="#99DC2626"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Row" Property="Background" Value="#20304A"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Row" Property="BorderBrush" Value="#DC2626"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style TargetType="ScrollBar">
  <Setter Property="Width" Value="10"/><Setter Property="MinWidth" Value="10"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="ScrollBar">
      <Track x:Name="PART_Track" Orientation="Vertical" IsDirectionReversed="True">
        <Track.Thumb><Thumb><Thumb.Template><ControlTemplate TargetType="Thumb"><Border CornerRadius="4" Background="#405776" Margin="2,2"/></ControlTemplate></Thumb.Template></Thumb></Track.Thumb>
      </Track>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>
<Style TargetType="ToolTip">
  <Setter Property="Background" Value="#20304A"/><Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="BorderBrush" Value="#405776"/>
</Style>
'@

$StWindowXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Instalador de Office | SafeTech" Width="1080" Height="740" MinWidth="960" MinHeight="620"
        WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="CanResize"
        Background="#09111F" Foreground="#F8FAFC" FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14"
        UseLayoutRounding="True" SnapsToDevicePixels="True" TextOptions.TextFormattingMode="Ideal">
  <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="58" ResizeBorderThickness="7" CornerRadius="0" GlassFrameThickness="0" UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>
  <Window.Resources>
    __RESOURCES__
  </Window.Resources>
  <Border x:Name="RootBorder" BorderBrush="#2A3C57" BorderThickness="1" Background="#09111F">
  <Grid x:Name="Root">
    <Grid.RowDefinitions>
      <RowDefinition Height="58"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="46"/>
    </Grid.RowDefinitions>

    <!-- Barra superior -->
    <Border Grid.Row="0" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="0,0,0,1">
      <Grid Margin="22,0,10,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Grid>
            <Image x:Name="ImgLogo" Height="30" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
            <StackPanel x:Name="LogoFallback" Orientation="Horizontal" Visibility="Collapsed">
              <TextBlock Text="Safe" Foreground="#DC2626" FontSize="22" FontWeight="Bold" FontStyle="Italic"/>
              <TextBlock Text="Tech" Foreground="#F8FAFC" FontSize="22" FontWeight="Black"/>
            </StackPanel>
          </Grid>
          <Border Width="1" Height="22" Background="#2A3C57" Margin="16,0"/>
          <TextBlock Text="Instalador de Office" Foreground="#DBE5F1" FontSize="14.5" FontWeight="SemiBold" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
          <Button x:Name="BtnAbout" Style="{StaticResource BtnLink}" Margin="0,0,10,0" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="¿Cómo funciona este instalador?">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE946;" FontFamily="{StaticResource Icons}" FontSize="14" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock Text="¿Cómo funciona?" FontSize="13"/>
            </StackPanel>
          </Button>
          <Border CornerRadius="999" Background="#1F059669" BorderBrush="#66059669" BorderThickness="1" Padding="10,3" Margin="0,0,12,0" VerticalAlignment="Center">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE72E;" FontFamily="{StaticResource Icons}" FontSize="11" Foreground="#6EE7B7" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBlock Text="Descarga oficial de Microsoft" Foreground="#6EE7B7" FontSize="11.5" FontWeight="SemiBold"/>
            </StackPanel>
          </Border>
          <Button x:Name="BtnMin" Style="{StaticResource BtnChrome}" Content="&#xE921;" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Minimizar"/>
          <Button x:Name="BtnMax" Style="{StaticResource BtnChrome}" Content="&#xE922;" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Maximizar"/>
          <Button x:Name="BtnClose" Style="{StaticResource BtnChrome}" Content="&#xE8BB;" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Cerrar"/>
        </StackPanel>
      </Grid>
    </Border>

    <Grid Grid.Row="1" x:Name="PageHost">

      <!-- INICIO -->
      <Grid x:Name="PageHome">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <StackPanel Margin="48,24,48,20" MaxWidth="1240">
            <Border HorizontalAlignment="Left" CornerRadius="999" Background="#26DC2626" Padding="12,4" Margin="0,0,0,12">
              <TextBlock Text="INSTALADOR OFICIAL SAFETECH" Foreground="#FCA5A5" FontSize="11" FontWeight="Black"/>
            </Border>
            <TextBlock Style="{StaticResource H1}" FontSize="38" Text="Instala Office en pocos pasos"/>
            <TextBlock Style="{StaticResource Lead}" MaxWidth="760" HorizontalAlignment="Left"
                       Text="Elige el producto que compraste. Descargamos Office directamente desde Microsoft, lo instalamos y te ayudamos a activarlo. No necesitas escribir comandos."/>
            <Button x:Name="BtnAboutHome" Style="{StaticResource BtnLink}" HorizontalAlignment="Left" Margin="-8,8,0,0">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE72E;" FontFamily="{StaticResource Icons}" FontSize="13" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBlock Text="¿Cómo funciona este instalador? Es seguro y usa los métodos oficiales de Microsoft"/>
              </StackPanel>
            </Button>
            <UniformGrid Columns="3" Margin="0,24,0,0">
              <Button x:Name="BtnPick2024" Style="{StaticResource Card}" Margin="0,0,9,0">
                <StackPanel>
                  <Grid>
                    <Border Width="58" Height="58" CornerRadius="16" Background="{StaticResource RedGrad}" HorizontalAlignment="Left">
                      <TextBlock Text="24" Foreground="White" FontSize="23" FontWeight="Black" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <Border CornerRadius="999" Background="#1F059669" BorderBrush="#66059669" BorderThickness="1" Padding="9,3" HorizontalAlignment="Right" VerticalAlignment="Top">
                      <TextBlock Text="MÁS NUEVO" Foreground="#6EE7B7" FontSize="10.5" FontWeight="Black"/>
                    </Border>
                  </Grid>
                  <TextBlock Text="Office 2024" Style="{StaticResource H2}" FontSize="22" Margin="0,18,0,2"/>
                  <TextBlock Text="Office LTSC Profesional Plus 2024" Style="{StaticResource Small}"/>
                  <Border Height="1" Background="#2A3C57" Margin="0,16"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Licencia permanente para este PC"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Activación automática con tu código" Margin="0,4,0,0"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Word, Excel, PowerPoint, Outlook y más" Margin="0,4,0,0"/>
                  <TextBlock Text="Elegir Office 2024  &#x2192;" Foreground="#FCA5A5" FontWeight="Bold" Margin="0,18,0,0"/>
                </StackPanel>
              </Button>
              <Button x:Name="BtnPick2021" Style="{StaticResource Card}" Margin="9,0,9,0">
                <StackPanel>
                  <Border Width="58" Height="58" CornerRadius="16" Background="#20304A" BorderBrush="#405776" BorderThickness="1" HorizontalAlignment="Left">
                    <TextBlock Text="21" Foreground="#F8FAFC" FontSize="23" FontWeight="Black" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <TextBlock Text="Office 2021" Style="{StaticResource H2}" FontSize="22" Margin="0,18,0,2"/>
                  <TextBlock Text="Office Profesional Plus 2021" Style="{StaticResource Small}"/>
                  <Border Height="1" Background="#2A3C57" Margin="0,16"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Licencia permanente para este PC"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Activación automática con tu código" Margin="0,4,0,0"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Word, Excel, PowerPoint, Outlook y más" Margin="0,4,0,0"/>
                  <TextBlock Text="Elegir Office 2021  &#x2192;" Foreground="#FCA5A5" FontWeight="Bold" Margin="0,18,0,0"/>
                </StackPanel>
              </Button>
              <Button x:Name="BtnPick365" Style="{StaticResource Card}" Margin="9,0,0,0">
                <StackPanel>
                  <Grid>
                    <Border Width="58" Height="58" CornerRadius="16" Background="#20304A" BorderBrush="#405776" BorderThickness="1" HorizontalAlignment="Left">
                      <TextBlock Text="365" Foreground="#F8FAFC" FontSize="19" FontWeight="Black" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <Border CornerRadius="999" Background="#262563EB" BorderBrush="#662563EB" BorderThickness="1" Padding="9,3" HorizontalAlignment="Right" VerticalAlignment="Top">
                      <TextBlock Text="CON TUTORIAL" Foreground="#93C5FD" FontSize="10.5" FontWeight="Black"/>
                    </Border>
                  </Grid>
                  <TextBlock Text="Microsoft 365" Style="{StaticResource H2}" FontSize="22" Margin="0,18,0,2"/>
                  <TextBlock Text="Aplicaciones de Microsoft 365" Style="{StaticResource Small}"/>
                  <Border Height="1" Background="#2A3C57" Margin="0,16"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Se activa con la cuenta que te entregamos"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Tutorial guiado, paso a paso" Margin="0,4,0,0"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="&#x2713;  Siempre actualizado" Margin="0,4,0,0"/>
                  <TextBlock Text="Elegir Microsoft 365  &#x2192;" Foreground="#FCA5A5" FontWeight="Bold" Margin="0,18,0,0"/>
                </StackPanel>
              </Button>
            </UniformGrid>
            <Border Style="{StaticResource Panel}" Margin="0,18,0,0" Padding="20,14" BorderBrush="#66DC2626" Background="#14DC2626">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center">
                  <TextBlock Text="¿Ya tienes Office instalado?" FontWeight="Bold" FontSize="15"/>
                  <TextBlock Style="{StaticResource Small}" Text="Puedes activarlo o ver el tutorial de la cuenta sin volver a instalar." Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                  <Button x:Name="BtnOnlyActivate" Style="{StaticResource BtnPrimary}" Margin="0,0,10,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE8D7;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,9,0"/>
                      <TextBlock Text="Solo activar Office 2021 o 2024"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnOnlyTutorial" Style="{StaticResource BtnPrimary}">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE77B;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,9,0"/>
                      <TextBlock Text="Tutorial de cuenta 365"/>
                    </StackPanel>
                  </Button>
                </StackPanel>
              </Grid>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </Grid>

      <!-- DATOS DE LA COMPRA -->
      <Grid x:Name="PageData" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid Margin="48,30,48,28" MaxWidth="1240">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="330"/></Grid.ColumnDefinitions>
            <StackPanel Margin="0,0,34,0">
              <TextBlock x:Name="TxtDataKicker" Style="{StaticResource Kicker}" Text="PASO 1 DE 3 · OFFICE 2024"/>
              <TextBlock Style="{StaticResource H1}" Text="Ingresa los datos de tu compra"/>
              <TextBlock Style="{StaticResource Lead}" Text="Los encontrarás en el correo que te envió SafeTech junto a tu licencia."/>
              <TextBlock Style="{StaticResource Label}" Text="Clave de producto" Margin="0,28,0,8"/>
              <TextBox x:Name="InKey" Style="{StaticResource Field}" Tag="XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" MaxLength="40"/>
              <TextBlock x:Name="TxtKeyHint" Style="{StaticResource Small}" Margin="2,8,0,0" Text="25 letras y números. Los guiones se agregan solos."/>
              <TextBlock Style="{StaticResource Label}" Text="Código de activación SafeTech" Margin="0,24,0,8"/>
              <Grid>
                <TextBox x:Name="InCode" Style="{StaticResource Field}" Tag="ST-XXXXXX" MaxLength="12"/>
                <Border x:Name="ChipCode" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,12,0" CornerRadius="999" Padding="11,5" Visibility="Collapsed" Background="#1F059669" BorderBrush="#66059669" BorderThickness="1">
                  <TextBlock x:Name="TxtCodeStatus" FontSize="12" FontWeight="Bold" Foreground="#6EE7B7" Text=""/>
                </Border>
              </Grid>
              <TextBlock Style="{StaticResource Small}" Margin="2,8,0,0" Text="Empieza con ST-. Con él generamos tu activación automáticamente."/>
              <StackPanel Orientation="Horizontal" Margin="0,30,0,0">
                <Button x:Name="BtnDataBack" Style="{StaticResource BtnSoft}" Content="Atrás" Margin="0,0,10,0"/>
                <Button x:Name="BtnDataNext" Style="{StaticResource BtnPrimary}" IsEnabled="False">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="TxtDataNext" Text="Continuar"/>
                    <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="10,1,0,0" FontSize="12"/>
                  </StackPanel>
                </Button>
              </StackPanel>
            </StackPanel>
            <StackPanel Grid.Column="1">
              <Border Style="{StaticResource Panel}" BorderBrush="#4D0891B2" Background="#140891B2">
                <StackPanel>
                  <TextBlock Text="&#xE946;" FontFamily="{StaticResource Icons}" FontSize="22" Foreground="#67E8F9"/>
                  <TextBlock Text="¿Para qué es el código?" FontWeight="Bold" FontSize="16" Margin="0,12,0,6"/>
                  <TextBlock Style="{StaticResource Body}" FontSize="13.5" Text="Office 2021 y 2024 se activan por teléfono. El instalador obtiene el ID de instalación de tu equipo y, con tu código, consigue el ID de confirmación de Microsoft. Tú no copias ningún número."/>
                  <TextBlock Style="{StaticResource Small}" Margin="0,12,0,0" Text="Cada activación usa 1 crédito de tu código."/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Panel}" Margin="0,14,0,0">
                <StackPanel>
                  <TextBlock Text="&#xE72E;" FontFamily="{StaticResource Icons}" FontSize="20" Foreground="#AEBDD0"/>
                  <TextBlock Text="Tus datos están seguros" FontWeight="Bold" FontSize="15" Margin="0,10,0,6"/>
                  <TextBlock Style="{StaticResource Small}" FontSize="13" Text="La clave se guarda solo en Office. Los registros del instalador muestran únicamente sus últimos 5 caracteres."/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </Grid>

      <!-- CONFIRMACION -->
      <Grid x:Name="PageConfirm" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid Margin="48,30,48,28" MaxWidth="1240">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="330"/></Grid.ColumnDefinitions>
            <StackPanel Margin="0,0,34,0">
              <TextBlock x:Name="TxtConfirmKicker" Style="{StaticResource Kicker}" Text="PASO 2 DE 3"/>
              <TextBlock x:Name="TxtConfirmTitle" Style="{StaticResource H1}" Text="Antes de comenzar"/>
              <TextBlock x:Name="TxtConfirmLead" Style="{StaticResource Lead}" Text="Lee esto con calma. Después el instalador trabaja solo."/>
              <Border Style="{StaticResource Panel}" Margin="0,24,0,0" Padding="20,18" x:Name="PanelDetected">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border x:Name="DetectedIconBg" Width="42" Height="42" CornerRadius="12" Background="#26F59E0B" VerticalAlignment="Top">
                    <TextBlock x:Name="DetectedIcon" Text="&#xE7BA;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="#FCD34D" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <StackPanel Grid.Column="1" Margin="16,0,0,0">
                    <TextBlock x:Name="TxtDetectedTitle" Text="Encontramos otra versión de Office" FontWeight="Bold" FontSize="15.5"/>
                    <TextBlock x:Name="TxtDetected" Style="{StaticResource Body}" FontSize="13.5" Margin="0,4,0,0" Text=""/>
                  </StackPanel>
                </Grid>
              </Border>
              <StackPanel x:Name="StackConfirmList" Margin="0,18,0,0"/>
              <CheckBox x:Name="ChkConfirm" Style="{StaticResource Check}" Margin="0,20,0,0" Content="Guardé mis documentos y leí estas indicaciones."/>
              <WrapPanel Margin="0,24,0,0">
                <Button x:Name="BtnConfirmBack" Style="{StaticResource BtnSoft}" Content="Atrás" Margin="0,0,10,10"/>
                <Button x:Name="BtnConfirmAlt" Style="{StaticResource BtnSoft}" Content="Solo activar" Margin="0,0,10,10" Visibility="Collapsed"/>
                <Button x:Name="BtnConfirmStart" Style="{StaticResource BtnPrimary}" IsEnabled="False" Margin="0,0,0,10">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="TxtConfirmStart" Text="Comenzar instalación"/>
                    <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="10,1,0,0" FontSize="12"/>
                  </StackPanel>
                </Button>
              </WrapPanel>
            </StackPanel>
            <StackPanel Grid.Column="1">
              <Border Style="{StaticResource Panel}">
                <StackPanel>
                  <TextBlock Text="Así funciona" FontWeight="Bold" FontSize="16" Margin="0,0,0,14"/>
                  <StackPanel x:Name="StackHowItWorks"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </Grid>

      <!-- PROGRESO -->
      <Grid x:Name="PageProgress" Visibility="Collapsed">
        <Grid Margin="48,30,48,24" MaxWidth="1240">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="330"/></Grid.ColumnDefinitions>
          <Grid Margin="0,0,34,0">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
            <StackPanel>
              <TextBlock x:Name="TxtProgKicker" Style="{StaticResource Kicker}" Text="PASO 3 DE 3 · INSTALANDO"/>
              <TextBlock x:Name="TxtProgTitle" Style="{StaticResource H1}" Text="Instalando Office 2024"/>
              <TextBlock x:Name="TxtProgDetail" Style="{StaticResource Lead}" Text="Preparando..." MinHeight="48"/>
              <Border x:Name="ProgTrack" Height="5" CornerRadius="3" Background="#20304A" Margin="0,16,0,18" ClipToBounds="True">
                <Border x:Name="ProgBar" Width="180" HorizontalAlignment="Left" CornerRadius="3" Background="{StaticResource RedGrad}">
                  <Border.RenderTransform><TranslateTransform x:Name="ProgBarMove" X="-180"/></Border.RenderTransform>
                  <Border.Triggers>
                    <EventTrigger RoutedEvent="FrameworkElement.Loaded">
                      <BeginStoryboard><Storyboard>
                        <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.X)" From="-180" To="760" Duration="0:0:1.8" RepeatBehavior="Forever">
                          <DoubleAnimation.EasingFunction><CubicEase EasingMode="EaseInOut"/></DoubleAnimation.EasingFunction>
                        </DoubleAnimation>
                      </Storyboard></BeginStoryboard>
                    </EventTrigger>
                  </Border.Triggers>
                </Border>
              </Border>
            </StackPanel>
            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
              <StackPanel x:Name="StackSteps"/>
            </ScrollViewer>
          </Grid>
          <StackPanel Grid.Column="1">
            <Border Style="{StaticResource Panel}">
              <StackPanel>
                <TextBlock Text="TIEMPO TRANSCURRIDO" Style="{StaticResource Kicker}" Foreground="#AEBDD0"/>
                <TextBlock x:Name="TxtElapsed" Text="00:00" FontSize="40" FontWeight="Bold" FontFamily="Segoe UI Variable Display, Segoe UI"/>
                <TextBlock Style="{StaticResource Small}" Text="La instalación suele tardar entre 10 y 40 minutos, según tu Internet."/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Panel}" Margin="0,14,0,0" BorderBrush="#4DF59E0B" Background="#12F59E0B">
              <StackPanel>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE7BA;" FontFamily="{StaticResource Icons}" FontSize="16" Foreground="#FCD34D" VerticalAlignment="Center"/>
                  <TextBlock Text="Mientras tanto" FontWeight="Bold" FontSize="15" Margin="10,0,0,0"/>
                </StackPanel>
                <TextBlock Style="{StaticResource Body}" FontSize="13.5" Margin="0,10,0,0" Text="• No apagues ni reinicies el computador."/>
                <TextBlock Style="{StaticResource Body}" FontSize="13.5" Margin="0,4,0,0" Text="• Si aparece una ventana de Microsoft, déjala trabajar: se cierra sola."/>
                <TextBlock Style="{StaticResource Body}" FontSize="13.5" Margin="0,4,0,0" Text="• Mantén el computador conectado a Internet y a la corriente."/>
              </StackPanel>
            </Border>
            <Border x:Name="PanelTutorialPromo" Style="{StaticResource Panel}" Margin="0,14,0,0" BorderBrush="#662563EB" Background="#142563EB" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="&#xE77B;" FontFamily="{StaticResource Icons}" FontSize="20" Foreground="#93C5FD"/>
                <TextBlock Text="Aprovecha la espera" FontWeight="Bold" FontSize="15.5" Margin="0,10,0,4"/>
                <TextBlock Style="{StaticResource Small}" FontSize="13" Text="Mientras Office se instala, puedes preparar la cuenta que te entregamos. Te guiamos paso a paso."/>
                <Button x:Name="BtnTutorialEarly" Style="{StaticResource BtnPrimary}" Margin="0,14,0,0" HorizontalAlignment="Stretch" Content="Comenzar el tutorial"/>
              </StackPanel>
            </Border>
          </StackPanel>
        </Grid>
      </Grid>

      <!-- LISTO -->
      <Grid x:Name="PageDone" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <StackPanel MaxWidth="660" Margin="40,44,40,28" HorizontalAlignment="Center">
            <Grid Width="96" Height="96" HorizontalAlignment="Center">
              <Ellipse Fill="#1F059669" Stroke="#66059669" StrokeThickness="1.5"/>
              <Ellipse Width="64" Height="64" Fill="#059669"/>
              <TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="30" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <TextBlock x:Name="TxtDoneTitle" Style="{StaticResource H1}" TextAlignment="Center" Margin="0,24,0,0" Text="¡Office quedó listo!"/>
            <TextBlock x:Name="TxtDoneText" Style="{StaticResource Lead}" TextAlignment="Center" Text=""/>
            <Border Style="{StaticResource Panel}" Margin="0,26,0,0">
              <StackPanel x:Name="StackDoneList"/>
            </Border>
            <Border x:Name="DoneNote" CornerRadius="14" Background="#12F59E0B" BorderBrush="#4DF59E0B" BorderThickness="1" Padding="16,12" Margin="0,14,0,0">
              <TextBlock x:Name="TxtDoneNote" Style="{StaticResource Body}" FontSize="13" Foreground="#FDE68A" Text=""/>
            </Border>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,26,0,0">
              <Button x:Name="BtnDoneWord" Style="{StaticResource BtnPrimary}" Margin="0,0,10,0">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE8A5;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                  <TextBlock Text="Abrir Word"/>
                </StackPanel>
              </Button>
              <Button x:Name="BtnDoneClose" Style="{StaticResource BtnSoft}" Content="Cerrar el instalador"/>
            </StackPanel>
          </StackPanel>
        </ScrollViewer>
      </Grid>

      <!-- ERROR -->
      <Grid x:Name="PageError" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <StackPanel MaxWidth="720" Margin="40,40,40,28" HorizontalAlignment="Center">
            <Grid Width="88" Height="88" HorizontalAlignment="Center">
              <Ellipse Fill="#1FDC2626" Stroke="#66DC2626" StrokeThickness="1.5"/>
              <Ellipse Width="58" Height="58" Fill="#DC2626"/>
              <TextBlock Text="&#xE711;" FontFamily="{StaticResource Icons}" FontSize="24" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <TextBlock x:Name="TxtErrTitle" Style="{StaticResource H1}" TextAlignment="Center" Margin="0,22,0,0" Text="No pudimos completar la instalación"/>
            <TextBlock Style="{StaticResource Lead}" TextAlignment="Center" Text="No te preocupes: puedes reintentar. Si se repite, crea el ZIP de diagnóstico y envíanoslo."/>
            <Border Style="{StaticResource Panel}" Margin="0,22,0,0" Padding="18">
              <TextBox x:Name="TxtErrDetail" IsReadOnly="True" TextWrapping="Wrap" Background="#00000000" BorderThickness="0"
                       Foreground="#DBE5F1" FontSize="13.5" MaxHeight="180" VerticalScrollBarVisibility="Auto" Text=""/>
            </Border>
            <WrapPanel HorizontalAlignment="Center" Margin="0,22,0,0">
              <Button x:Name="BtnErrRetry" Style="{StaticResource BtnPrimary}" Content="Reintentar" Margin="0,0,10,10"/>
              <Button x:Name="BtnErrZip" Style="{StaticResource BtnSoft}" Margin="0,0,10,10">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE8B7;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,9,0"/>
                  <TextBlock Text="Crear ZIP para soporte"/>
                </StackPanel>
              </Button>
              <Button x:Name="BtnErrHome" Style="{StaticResource BtnSoft}" Content="Volver al inicio" Margin="0,0,0,10"/>
            </WrapPanel>
            <TextBlock x:Name="TxtErrHelp" Style="{StaticResource Small}" TextAlignment="Center" Margin="0,8,0,0"
                       Text="Envía el ZIP por WhatsApp al +56 9 5101 7168 o a contacto@safe-tech.cl."/>
          </StackPanel>
        </ScrollViewer>
      </Grid>

      <!-- TUTORIAL MICROSOFT 365 -->
      <Grid x:Name="PageTutorial" Visibility="Collapsed">
        <Grid Margin="32,24,40,22" MaxWidth="1380">
          <Grid.ColumnDefinitions><ColumnDefinition Width="268"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border Style="{StaticResource Panel}" Padding="14,18" Margin="0,0,28,0">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <StackPanel Margin="8,0,8,10">
                <TextBlock Text="TUTORIAL" Style="{StaticResource Kicker}"/>
                <TextBlock Text="Tu cuenta de Microsoft 365" FontWeight="Bold" FontSize="16" TextWrapping="Wrap"/>
                <Border x:Name="ChipInstall" CornerRadius="999" Padding="10,4" Margin="0,12,0,0" HorizontalAlignment="Left" Background="#26F59E0B" BorderBrush="#66F59E0B" BorderThickness="1" Visibility="Collapsed">
                  <TextBlock x:Name="TxtChipInstall" FontSize="11.5" FontWeight="Bold" Foreground="#FCD34D" Text="Office instalándose"/>
                </Border>
                <Button x:Name="BtnTutProgress" Style="{StaticResource BtnLink}" FontSize="12.5" Margin="-8,6,0,0" HorizontalAlignment="Left" Visibility="Collapsed" Content="Ver el avance de la instalación"/>
                <Button x:Name="BtnTutHome" Style="{StaticResource BtnLink}" FontSize="12.5" Margin="-8,6,0,0" HorizontalAlignment="Left">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE72B;" FontFamily="{StaticResource Icons}" FontSize="11" VerticalAlignment="Center" Margin="0,1,8,0"/>
                    <TextBlock Text="Volver al inicio"/>
                  </StackPanel>
                </Button>
              </StackPanel>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel x:Name="StackTutSteps"/>
              </ScrollViewer>
            </Grid>
          </Border>
          <Grid Grid.Column="1">
            <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <ScrollViewer x:Name="TutScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
              <StackPanel x:Name="TutContent" Margin="0,6,10,10">
                <TextBlock x:Name="TxtTutKicker" Style="{StaticResource Kicker}" Text="PASO 1 DE 10"/>
                <TextBlock x:Name="TxtTutTitle" Style="{StaticResource H1}" FontSize="29" Text=""/>
                <StackPanel x:Name="StackTutBody" Margin="0,16,0,0"/>
                <Border x:Name="TutNote" CornerRadius="14" BorderThickness="1" Padding="16,12" Margin="0,16,0,0">
                  <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <TextBlock x:Name="TutNoteIcon" FontFamily="{StaticResource Icons}" FontSize="18" Margin="0,1,12,0" Text="&#xE946;"/>
                    <StackPanel Grid.Column="1">
                      <TextBlock x:Name="TxtTutNoteLabel" FontSize="11" FontWeight="Black" Margin="0,2,0,4" Text="CONSEJO"/>
                      <TextBlock x:Name="TxtTutNote" Style="{StaticResource Body}" FontSize="14" Text=""/>
                    </StackPanel>
                  </Grid>
                </Border>
                <WrapPanel x:Name="StackTutActions" Margin="0,16,0,0"/>
                <Border x:Name="TutImageFrame" CornerRadius="16" BorderBrush="#2A3C57" BorderThickness="1" Background="#0C1626" Margin="0,18,0,0" Padding="8" HorizontalAlignment="Left">
                  <Image x:Name="ImgTut" MaxHeight="300" MaxWidth="600" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <Border Grid.Row="1" BorderBrush="#2A3C57" BorderThickness="0,1,0,0" Padding="0,16,0,0">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center" Margin="0,0,24,0">
                  <Border Height="6" CornerRadius="3" Background="#20304A">
                    <Border x:Name="TutProgFill" HorizontalAlignment="Left" Width="0" CornerRadius="3" Background="{StaticResource RedGrad}"/>
                  </Border>
                  <TextBlock x:Name="TxtTutProgress" Style="{StaticResource Small}" Margin="0,6,0,0" Text=""/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                  <Button x:Name="BtnTutBack" Style="{StaticResource BtnSoft}" Margin="0,0,10,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE72B;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,1,10,0" FontSize="12"/>
                      <TextBlock Text="Atrás"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnTutNext" Style="{StaticResource BtnPrimary}" MinWidth="150">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock x:Name="TxtTutNext" Text="Siguiente"/>
                      <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="10,1,0,0" FontSize="12"/>
                    </StackPanel>
                  </Button>
                </StackPanel>
              </Grid>
            </Border>
          </Grid>
        </Grid>
      </Grid>
    </Grid>

    <!-- Pie -->
    <Border Grid.Row="2" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="0,1,0,0">
      <Grid Margin="22,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="¿Necesitas ayuda?" Foreground="#AEBDD0" FontSize="12.5" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <Button x:Name="BtnHelpCenter" Style="{StaticResource BtnLink}" FontSize="12.5" Content="Centro de Ayuda"/>
          <TextBlock Text="·" Foreground="#5D7391" VerticalAlignment="Center"/>
          <Button x:Name="BtnHelpWa" Style="{StaticResource BtnLink}" FontSize="12.5" Content="WhatsApp +56 9 5101 7168"/>
        </StackPanel>
        <TextBlock x:Name="TxtVersion" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="#5D7391" FontSize="12" Text="SafeTech SpA"/>
      </Grid>
    </Border>
  </Grid>
  </Border>
</Window>
'@

$StDialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SafeTech" SizeToContent="WidthAndHeight" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" ShowInTaskbar="False" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI Variable Text, Segoe UI"
        FontSize="14" Foreground="#F8FAFC" UseLayoutRounding="True" TextOptions.TextFormattingMode="Ideal">
  <Window.Resources>
    __RESOURCES__
  </Window.Resources>
  <Border Margin="24" CornerRadius="22" Background="#111C2F" BorderBrush="#405776" BorderThickness="1" MinWidth="440" MaxWidth="600">
    <Border.Effect><DropShadowEffect Color="#000000" BlurRadius="40" ShadowDepth="10" Direction="270" Opacity="0.55"/></Border.Effect>
    <Grid>
      <Border Height="2" VerticalAlignment="Top" Margin="40,0" Background="{StaticResource TopLine}"/>
      <StackPanel Margin="30,28,30,26">
        <Grid x:Name="DlgDrag" Background="#00000000">
          <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border x:Name="DlgIconBg" Width="46" Height="46" CornerRadius="14" Background="#262563EB" VerticalAlignment="Top">
            <TextBlock x:Name="DlgIcon" FontFamily="{StaticResource Icons}" FontSize="20" Foreground="#93C5FD" HorizontalAlignment="Center" VerticalAlignment="Center" Text="&#xE946;"/>
          </Border>
          <StackPanel Grid.Column="1" Margin="16,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="DlgKicker" Style="{StaticResource Kicker}" Margin="0,0,0,4" Text="SAFETECH"/>
            <TextBlock x:Name="DlgTitle" Style="{StaticResource H2}" FontSize="20" Text=""/>
          </StackPanel>
        </Grid>
        <TextBlock x:Name="DlgMessage" Style="{StaticResource Body}" Margin="0,16,0,0" Text=""/>
        <ContentControl x:Name="DlgExtra" Margin="0,16,0,0"/>
        <WrapPanel x:Name="DlgButtons" HorizontalAlignment="Right" Margin="0,22,0,0"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
'@

# ----- Utilidades de interfaz -------------------------------------------------------------
function New-StXamlElement {
    param([string]$Xaml)
    $full = $Xaml -replace '^\s*<(\w+)', '<$1 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
    return [Windows.Markup.XamlReader]::Parse($full)
}

function Get-StImage {
    param([string]$RelPath)
    try {
        $file = $null
        if ($env:SAFETECH_OFFICE_ASSETS) {
            $file = Join-Path $env:SAFETECH_OFFICE_ASSETS $RelPath
        } else {
            $file = Join-Path $St.CacheDir ($RelPath -replace '[\\/]', '_')
            if (-not (Test-Path -LiteralPath $file)) { Save-StAssets -Paths @($RelPath) -CacheDir $St.CacheDir -TimeoutSec 8 }
        }
        if (-not (Test-Path -LiteralPath $file)) { return $null }
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit()
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.UriSource = New-Object Uri($file)
        $bmp.EndInit()
        $bmp.Freeze()
        return $bmp
    } catch { return $null }
}

function Open-StUnelevated {
    param([string]$Target)
    # explorer.exe abre la ruta o el enlace con los permisos normales del usuario (sin administrador).
    try { Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList "`"$Target`"" }
    catch { try { Start-Process $Target } catch { } }
}

function Get-StWordPath {
    foreach ($p in @("$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
                     "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE")) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    try {
        $reg = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe' -ErrorAction Stop
        if ($reg.'(default)' -and (Test-Path -LiteralPath $reg.'(default)')) { return $reg.'(default)' }
    } catch { }
    return $null
}

function Open-StWord {
    $word = Get-StWordPath
    if ($word) { Open-StUnelevated $word; return $true }
    [void](Show-StDialog -Kind 'warn' -Title 'No encontramos Word' -Message 'Word todavía no aparece en este computador. Si la instalación terminó hace poco, espera un minuto y vuelve a intentarlo, o búscalo en el menú Inicio.')
    return $false
}

function Initialize-StLogs {
    if ($St.LogDir) { return }
    $dir = Join-Path $StConfig.DataDir ('registros\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $St.LogDir = $dir
    $St.Sync.LogFile = Join-Path $dir 'proceso.log'
    Write-StLog "Instalador SafeTech v$($StConfig.Version) iniciado. Windows $([Environment]::OSVersion.Version)."
}

function Show-StPage {
    param([string]$Name)
    foreach ($p in $St.Pages) { $St.UI["Page$p"].Visibility = 'Collapsed' }
    $el = $St.UI["Page$Name"]
    $el.Visibility = 'Visible'
    $St.CurrentPage = $Name
    if ($St.NoAnim) { $el.Opacity = 1; return }
    $tt = New-Object System.Windows.Media.TranslateTransform 0, 18
    $el.RenderTransform = $tt
    $ease = New-Object System.Windows.Media.Animation.QuinticEase
    $ease.EasingMode = 'EaseOut'
    $dur = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(520))
    $fade = [System.Windows.Media.Animation.DoubleAnimation]::new([double]0, [double]1, $dur)
    $fade.EasingFunction = $ease
    $move = [System.Windows.Media.Animation.DoubleAnimation]::new([double]18, [double]0, $dur)
    $move.EasingFunction = $ease
    $el.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fade)
    $tt.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $move)
}

# ----- Ventanas emergentes -------------------------------------------------------------------
function Show-StDialog {
    param(
        [string]$Title,
        [string]$Message,
        [ValidateSet('info', 'success', 'warn', 'error', 'question', 'key')][string]$Kind = 'info',
        [string[]]$Buttons = @('Aceptar'),
        [string]$Primary,
        [string]$Kicker = 'SAFETECH',
        $Extra
    )
    $xaml = $StDialogXaml.Replace('__RESOURCES__', $StResXaml)
    $dlg = [Windows.Markup.XamlReader]::Parse($xaml)
    $look = @{
        info     = @('#262563EB', '#93C5FD', [char]0xE946)
        success  = @('#26059669', '#6EE7B7', [char]0xE73E)
        warn     = @('#26F59E0B', '#FCD34D', [char]0xE7BA)
        error    = @('#26DC2626', '#FCA5A5', [char]0xEA39)
        question = @('#262563EB', '#93C5FD', [char]0xE9CE)
        key      = @('#26DC2626', '#FCA5A5', [char]0xE8D7)
    }[$Kind]
    $conv = New-Object System.Windows.Media.BrushConverter
    $dlg.FindName('DlgIconBg').Background = $conv.ConvertFromString($look[0])
    $icon = $dlg.FindName('DlgIcon')
    $icon.Foreground = $conv.ConvertFromString($look[1])
    $icon.Text = [string]$look[2]
    $dlg.FindName('DlgKicker').Text = $Kicker.ToUpper()
    $dlg.FindName('DlgTitle').Text = $Title
    $msg = $dlg.FindName('DlgMessage')
    $msg.Text = $Message
    if (-not $Message) { $msg.Visibility = 'Collapsed' }
    $extraHost = $dlg.FindName('DlgExtra')
    if ($Extra) { $extraHost.Content = $Extra } else { $extraHost.Visibility = 'Collapsed' }
    if (-not $Primary) { $Primary = $Buttons[0] }
    $panel = $dlg.FindName('DlgButtons')
    foreach ($label in $Buttons) {
        $b = New-Object System.Windows.Controls.Button
        $b.Content = $label
        $b.Tag = $label
        $b.Margin = '10,0,0,0'
        $b.MinWidth = 120
        $styleKey = if ($label -eq $Primary) { 'BtnPrimary' } else { 'BtnSoft' }
        $b.Style = $dlg.FindResource($styleKey)
        if ($label -eq $Primary) { $b.IsDefault = $true }
        $b.Add_Click({
            param($s, $e)
            $w = [System.Windows.Window]::GetWindow($s)
            $w.Tag = $s.Tag
            $w.Close()
        })
        [void]$panel.Children.Add($b)
    }
    $dlg.FindName('DlgDrag').Add_MouseLeftButtonDown({ param($s, $e) try { [System.Windows.Window]::GetWindow($s).DragMove() } catch { } })
    if ($St.Window -and $St.Window.IsVisible) { $dlg.Owner = $St.Window } else { $dlg.WindowStartupLocation = 'CenterScreen' }
    if ($St.SelfTest) { return $dlg }
    $St.DialogOpen = $true
    try { [void]$dlg.ShowDialog() } finally { $St.DialogOpen = $false }
    return [string]$dlg.Tag
}

function New-StDialogField {
    param([string]$Label, [string]$Placeholder, [string]$Value = '', [int]$MaxLength = 80, [string]$Hint)
    $sp = New-Object System.Windows.Controls.StackPanel
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Label
    $tb.FontWeight = 'SemiBold'
    $tb.Margin = '0,0,0,8'
    [void]$sp.Children.Add($tb)
    $box = New-Object System.Windows.Controls.TextBox
    $box.Style = $St.Window.FindResource('Field')
    $box.Tag = $Placeholder
    $box.Text = $Value
    $box.MaxLength = $MaxLength
    [void]$sp.Children.Add($box)
    if ($Hint) {
        $h = New-Object System.Windows.Controls.TextBlock
        $h.Text = $Hint
        $h.Foreground = '#AEBDD0'
        $h.FontSize = 12.5
        $h.TextWrapping = 'Wrap'
        $h.Margin = '2,8,0,0'
        [void]$sp.Children.Add($h)
    }
    return @{ Panel = $sp; Box = $box }
}

# ----- Tareas en segundo plano -------------------------------------------------------------------
function Start-StAsync {
    param([Parameter(Mandatory)][scriptblock]$Work, [hashtable]$Params = @{}, [scriptblock]$OnDone)
    $script = "param(`$P, `$sync, `$StConfig)`n" + $St.EngineText + "`n" + @'
try {
    $__out = & {
'@ + "`n" + $Work.ToString() + "`n" + @'
    }
    @{ Ok = $true; Result = $__out; Error = $null }
} catch {
    Write-StLog ("ERROR: " + $_.Exception.Message)
    @{ Ok = $false; Result = $null; Error = $_.Exception.Message }
}
'@
    $ps = [powershell]::Create()
    $ps.RunspacePool = $St.Pool
    [void]$ps.AddScript($script).AddArgument($Params).AddArgument($St.Sync).AddArgument($StConfig)
    $handle = $ps.BeginInvoke()
    [void]$St.Jobs.Add(@{ PS = $ps; Handle = $handle; OnDone = $OnDone })
}

function Invoke-StJobPump {
    foreach ($job in @($St.Jobs)) {
        if (-not $job.Handle.IsCompleted) { continue }
        [void]$St.Jobs.Remove($job)
        $res = $null
        try {
            $out = $job.PS.EndInvoke($job.Handle)
            if ($out -and $out.Count -gt 0) { $res = $out[$out.Count - 1] }
            if ($res -is [System.Management.Automation.PSObject]) { $res = $res.BaseObject }
        } catch {
            $res = @{ Ok = $false; Result = $null; Error = $_.Exception.Message }
        } finally {
            $job.PS.Dispose()
        }
        if (-not ($res -is [hashtable])) { $res = @{ Ok = $false; Result = $null; Error = 'La tarea en segundo plano se detuvo sin informar el motivo. Revisa el registro o crea el ZIP de diagnóstico.' } }
        if ($job.OnDone) {
            try { & $job.OnDone $res.Ok $res.Result $res.Error }
            catch { Write-StLog "Error de interfaz: $($_.Exception.Message)"; Show-StError -Title 'Ocurrió un error inesperado' -Detail $_.Exception.Message }
        }
    }
}

# ----- Lista de pasos (pantalla de avance) ----------------------------------------------------
$StStepDefs = @{
    check       = @('Revisar el computador', 'Windows, espacio libre y conexión con Microsoft')
    restore     = @('Crear un punto de restauración', 'Un respaldo de Windows por si necesitas volver atrás')
    download    = @('Descargar el instalador de Microsoft', 'Herramienta oficial de implementación de Office')
    clean       = @('Cerrar Office y limpiar licencias anteriores', 'Claves y configuraciones que podrían causar errores')
    remove      = @('Desinstalar versiones anteriores', 'Solo si encontramos otro Office en el equipo')
    install     = @('Instalar Office', 'Se descarga desde Microsoft; verás su ventana de avance')
    verify      = @('Verificar la instalación', 'Confirmamos que Office quedó instalado')
    act_key     = @('Registrar tu clave de producto', 'Primero se intenta la activación en línea')
    act_iid     = @('Obtener el ID de instalación', 'Lo genera Office en este computador')
    act_cid     = @('Generar el ID de confirmación', 'Con tu código SafeTech, sin copiar números')
    act_confirm = @('Confirmar la activación', 'Microsoft valida el ID y activa Office')
    account     = @('Conectar tu cuenta', 'Tutorial guiado, paso a paso')
}

function Initialize-StSteps {
    param([string[]]$Ids)
    $St.UI.StackSteps.Children.Clear()
    $St.StepRows = @{}
    $St.StepOrder = $Ids
    $St.Sync.Steps.Clear()
    $n = 0
    foreach ($id in $Ids) {
        $n++
        $def = $StStepDefs[$id]
        $title = $def[0]
        if ($id -eq 'install' -and $St.Product) { $title = "Instalar $($St.Product.Name)" }
        $row = New-StXamlElement @"
<Border CornerRadius="14" Padding="12,10" Margin="0,0,0,6" Background="#00000000">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="40"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Grid Width="30" Height="30" VerticalAlignment="Top" Margin="0,1,0,0">
      <Ellipse x:Name="Ring" Stroke="#2A3C57" StrokeThickness="1.5" Fill="#00000000"/>
      <TextBlock x:Name="Num" Text="$n" Foreground="#AEBDD0" FontSize="12.5" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      <TextBlock x:Name="Glyph" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
      <Path x:Name="Spin" Width="30" Height="30" Data="M 15,1.5 A 13.5,13.5 0 0 1 28.5,15" Stroke="#DC2626" StrokeThickness="3" StrokeStartLineCap="Round" StrokeEndLineCap="Round" RenderTransformOrigin="0.5,0.5" Visibility="Collapsed">
        <Path.RenderTransform><RotateTransform Angle="0"/></Path.RenderTransform>
        <Path.Triggers>
          <EventTrigger RoutedEvent="FrameworkElement.Loaded">
            <BeginStoryboard><Storyboard>
              <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(RotateTransform.Angle)" From="0" To="360" Duration="0:0:0.9" RepeatBehavior="Forever"/>
            </Storyboard></BeginStoryboard>
          </EventTrigger>
        </Path.Triggers>
      </Path>
    </Grid>
    <StackPanel Grid.Column="1" VerticalAlignment="Center">
      <TextBlock x:Name="Title" Foreground="#AEBDD0" FontSize="14.5" FontWeight="SemiBold" TextWrapping="Wrap"/>
      <TextBlock x:Name="Sub" Foreground="#7F93AE" FontSize="12.5" TextWrapping="Wrap" Margin="0,2,0,0"/>
    </StackPanel>
  </Grid>
</Border>
"@
        $row.FindName('Title').Text = $title
        $row.FindName('Sub').Text = $def[1]
        [void]$St.UI.StackSteps.Children.Add($row)
        $St.StepRows[$id] = @{ Row = $row; State = 'pending'; Sub = $def[1] }
        $St.Sync.Steps[$id] = 'pending'
    }
}

function Update-StStepRow {
    param([string]$Id, [string]$State)
    $r = $St.StepRows[$Id]
    if (-not $r -or $r.State -eq $State) { return }
    $r.State = $State
    $row = $r.Row
    $conv = New-Object System.Windows.Media.BrushConverter
    $ring = $row.FindName('Ring'); $num = $row.FindName('Num'); $glyph = $row.FindName('Glyph')
    $spin = $row.FindName('Spin'); $title = $row.FindName('Title'); $sub = $row.FindName('Sub')
    $spin.Visibility = 'Collapsed'; $glyph.Visibility = 'Collapsed'; $num.Visibility = 'Visible'
    $row.Background = $conv.ConvertFromString('#00000000')
    $ring.Fill = $conv.ConvertFromString('#00000000'); $ring.Stroke = $conv.ConvertFromString('#2A3C57')
    $title.Foreground = $conv.ConvertFromString('#AEBDD0'); $sub.Text = $r.Sub
    switch ($State) {
        'running' {
            $spin.Visibility = 'Visible'; $num.Foreground = $conv.ConvertFromString('#FCA5A5')
            $row.Background = $conv.ConvertFromString('#17243A'); $title.Foreground = $conv.ConvertFromString('#F8FAFC')
        }
        'done' {
            $num.Visibility = 'Collapsed'; $glyph.Visibility = 'Visible'; $glyph.Text = [string][char]0xE73E
            $glyph.Foreground = $conv.ConvertFromString('#FFFFFF')
            $ring.Fill = $conv.ConvertFromString('#059669'); $ring.Stroke = $conv.ConvertFromString('#059669')
            $title.Foreground = $conv.ConvertFromString('#DBE5F1')
        }
        'skip' {
            $num.Visibility = 'Collapsed'; $glyph.Visibility = 'Visible'; $glyph.Text = [string][char]0xE738
            $glyph.Foreground = $conv.ConvertFromString('#7F93AE'); $sub.Text = 'No fue necesario'
        }
        'error' {
            $num.Visibility = 'Collapsed'; $glyph.Visibility = 'Visible'; $glyph.Text = [string][char]0xE711
            $glyph.Foreground = $conv.ConvertFromString('#FFFFFF')
            $ring.Fill = $conv.ConvertFromString('#DC2626'); $ring.Stroke = $conv.ConvertFromString('#DC2626')
            $title.Foreground = $conv.ConvertFromString('#FCA5A5')
        }
        default { $num.Foreground = $conv.ConvertFromString('#AEBDD0') }
    }
}

function Set-StUiStep {
    param([string]$Id, [string]$State, [string]$Detail)
    $St.Sync.Steps[$Id] = $State
    if ($Detail) { $St.Sync.Detail = $Detail }
    Update-StStepRow $Id $State
    if ($Detail) { $St.UI.TxtProgDetail.Text = $Detail }
}

function Update-StProgressUi {
    foreach ($id in $St.StepOrder) {
        $state = $St.Sync.Steps[$id]
        if ($state) { Update-StStepRow $id $state }
    }
    if ($St.Sync.Detail -and $St.UI.TxtProgDetail.Text -ne $St.Sync.Detail) { $St.UI.TxtProgDetail.Text = $St.Sync.Detail }
    if ($St.StartTime) {
        $span = (Get-Date) - $St.StartTime
        $txt = '{0:00}:{1:00}' -f [Math]::Floor($span.TotalMinutes), $span.Seconds
        if ($span.TotalHours -ge 1) { $txt = '{0}:{1:00}:{2:00}' -f [Math]::Floor($span.TotalHours), $span.Minutes, $span.Seconds }
        $St.UI.TxtElapsed.Text = $txt
        if ($St.Mode -eq 'install' -and $St.Product -and $St.Product.Id -eq '365') {
            if ($St.Busy) { $St.UI.TxtChipInstall.Text = "Office instalándose · $txt" }
        }
    }
}

function Stop-StProgressAnimation {
    $St.UI.ProgTrack.Visibility = 'Collapsed'
}

# ----- Errores ---------------------------------------------------------------------------------
function Show-StError {
    param([string]$Title = 'No pudimos completar la instalación', [string]$Detail, [string]$RetryLabel = 'Reintentar', [scriptblock]$OnRetry)
    $St.Busy = $false
    Stop-StProgressAnimation
    $St.UI.TxtErrTitle.Text = $Title
    $St.UI.TxtErrDetail.Text = $Detail
    $St.UI.BtnErrRetry.Content = $RetryLabel
    $St.ErrRetry = $OnRetry
    $St.UI.BtnErrRetry.Visibility = $(if ($OnRetry) { 'Visible' } else { 'Collapsed' })
    Show-StPage 'Error'
}

# ----- Flujo: instalacion ------------------------------------------------------------------------
function Select-StProduct {
    param([string]$Id)
    $St.Product = Get-StProduct $Id
    $St.Mode = 'install'
    if ($St.Product.NeedsKey) {
        Show-StDataPage
    } else {
        Show-StConfirmPage
    }
}

function Show-StDataPage {
    $p = $St.Product
    $total = if ($St.Mode -eq 'activate') { '2' } else { '3' }
    $St.UI.TxtDataKicker.Text = "PASO 1 DE $total · $($p.Name.ToUpper())"
    $St.UI.TxtDataNext.Text = $(if ($St.Mode -eq 'activate') { 'Activar Office' } else { 'Continuar' })
    Update-StDataValidity
    Show-StPage 'Data'
    [void]$St.UI.InKey.Focus()
}

function Update-StDataValidity {
    $keyOk = Test-StProductKey $St.UI.InKey.Text
    $code = ConvertTo-StAccessCode $St.UI.InCode.Text
    $codeOk = ($null -ne $code) -and ($St.CodeState -ne 'invalid')
    $St.UI.BtnDataNext.IsEnabled = ($keyOk -and $codeOk)
    $conv = New-Object System.Windows.Media.BrushConverter
    if ($St.UI.InKey.Text.Length -eq 0) {
        $St.UI.TxtKeyHint.Text = '25 letras y números. Los guiones se agregan solos.'
        $St.UI.TxtKeyHint.Foreground = $conv.ConvertFromString('#AEBDD0')
    } elseif ($keyOk) {
        $St.UI.TxtKeyHint.Text = 'Clave completa.'
        $St.UI.TxtKeyHint.Foreground = $conv.ConvertFromString('#6EE7B7')
    } else {
        $n = ($St.UI.InKey.Text -replace '[^A-Za-z0-9]', '').Length
        $St.UI.TxtKeyHint.Text = "Llevas $n de 25 caracteres."
        $St.UI.TxtKeyHint.Foreground = $conv.ConvertFromString('#AEBDD0')
    }
}

function Set-StCodeChip {
    param([string]$State, [string]$Text)
    $St.CodeState = $State
    $chip = $St.UI.ChipCode
    $conv = New-Object System.Windows.Media.BrushConverter
    if ($State -eq 'empty') { $chip.Visibility = 'Collapsed'; Update-StDataValidity; return }
    $colors = @{
        valid    = @('#1F059669', '#66059669', '#6EE7B7')
        invalid  = @('#26DC2626', '#66DC2626', '#FCA5A5')
        unknown  = @('#26F59E0B', '#66F59E0B', '#FCD34D')
        checking = @('#20304A', '#405776', '#DBE5F1')
    }[$State]
    $chip.Background = $conv.ConvertFromString($colors[0])
    $chip.BorderBrush = $conv.ConvertFromString($colors[1])
    $St.UI.TxtCodeStatus.Foreground = $conv.ConvertFromString($colors[2])
    $St.UI.TxtCodeStatus.Text = $Text
    $chip.Visibility = 'Visible'
    Update-StDataValidity
}

function Start-StCodeCheck {
    $code = ConvertTo-StAccessCode $St.UI.InCode.Text
    if (-not $code) {
        if ($St.UI.InCode.Text.Length -eq 0) { Set-StCodeChip 'empty' '' } else { Set-StCodeChip 'checking' 'Formato ST-XXXXXX' ; $St.CodeState = 'format' ; Update-StDataValidity }
        return
    }
    Set-StCodeChip 'checking' 'Verificando...'
    $St.CodeChecking = $code
    Start-StAsync -Params @{ Code = $code } -Work { Test-StCodeOnline -Code $P.Code } -OnDone {
        param($ok, $r, $err)
        if ($St.CodeChecking -ne (ConvertTo-StAccessCode $St.UI.InCode.Text)) { return }
        if (-not $ok) { Set-StCodeChip 'unknown' 'No se pudo verificar'; return }
        if ($r.Valid) {
            $label = 'Código válido'
            if ($null -ne $r.Remaining -and "$($r.Remaining)" -ne '') {
                $n = [int]$r.Remaining
                if ($n -le 0) { Set-StCodeChip 'invalid' 'Sin créditos'; return }
                $label = "Válido · $n " + $(if ($n -eq 1) { 'crédito' } else { 'créditos' })
            }
            Set-StCodeChip 'valid' $label
        } else {
            Set-StCodeChip 'invalid' 'Código no válido'
        }
    }
}

function Show-StConfirmPage {
    $p = $St.Product
    $state = Get-StOfficeState
    $St.OfficeState = $state
    $conv = New-Object System.Windows.Media.BrushConverter
    $ui = $St.UI
    $ui.ChkConfirm.IsChecked = $false
    $ui.ChkConfirm.Visibility = 'Visible'
    $ui.BtnConfirmAlt.Visibility = 'Collapsed'
    $ui.BtnConfirmStart.IsEnabled = $false
    $ui.TxtConfirmStart.Text = 'Comenzar instalación'
    $ui.TxtConfirmTitle.Text = "Antes de instalar $($p.Name)"
    $ui.TxtConfirmLead.Text = 'Lee esto con calma. Después el instalador trabaja solo.'
    $ui.TxtConfirmKicker.Text = $(if ($p.NeedsKey) { "PASO 2 DE 3 · $($p.Name.ToUpper())" } else { "PASO 1 DE 2 · $($p.Name.ToUpper())" })
    $St.ConfirmAction = 'install'

    $same = Test-StProductInstalled -ProductId $p.Id -State $state
    $others = @($state.Friendly)
    $setDetected = {
        param($kind, $title, $text)
        $map = @{ ok = @('#26059669', '#6EE7B7', [char]0xE73E); warn = @('#26F59E0B', '#FCD34D', [char]0xE7BA); info = @('#262563EB', '#93C5FD', [char]0xE946) }[$kind]
        $ui.DetectedIconBg.Background = $conv.ConvertFromString($map[0])
        $ui.DetectedIcon.Foreground = $conv.ConvertFromString($map[1])
        $ui.DetectedIcon.Text = [string]$map[2]
        $ui.TxtDetectedTitle.Text = $title
        $ui.TxtDetected.Text = $text
    }

    $items = New-Object System.Collections.Generic.List[string]
    if ($p.Id -eq '365' -and $same) {
        & $setDetected 'ok' 'Microsoft 365 ya está instalado' 'No necesitas reinstalarlo. Solo falta conectar la cuenta que te entregó SafeTech.'
        $ui.TxtConfirmTitle.Text = 'Microsoft 365 ya está en tu equipo'
        $ui.TxtConfirmLead.Text = 'Puede ser una edición Personal, Familia, Empresa o Educación. Funciona con la cuenta entregada por SafeTech.'
        $items.Add('No se desinstalará ni se cambiará nada en tu Office.')
        $items.Add('No es necesario cerrar tu cuenta personal de OneDrive o Teams.')
        $items.Add('Te guiaremos para agregar la cuenta de SafeTech en Word.')
        $ui.ChkConfirm.Visibility = 'Collapsed'
        $ui.BtnConfirmStart.IsEnabled = $true
        $ui.TxtConfirmStart.Text = 'Ir al tutorial de la cuenta'
        $St.ConfirmAction = 'tutorial'
    } else {
        if ($same) {
            & $setDetected 'info' "$($p.Name) ya está instalado" "Puedes solo activarlo (más rápido) o reinstalarlo desde cero si tienes problemas con él."
            if ($p.NeedsKey) { $ui.BtnConfirmAlt.Visibility = 'Visible'; $ui.BtnConfirmAlt.Content = 'Solo activar (recomendado)' }
            $ui.TxtConfirmStart.Text = 'Reinstalar desde cero'
        } elseif ($state.HasAny) {
            & $setDetected 'warn' 'Encontramos otra versión de Office' ("En este computador está instalado: {0}. Se desinstalará por completo antes de instalar {1}." -f ($others -join ', '), $p.Name)
        } else {
            & $setDetected 'ok' 'No encontramos otra versión de Office' "El computador está listo para instalar $($p.Name)."
        }
        $items.Add('Se cerrarán Word, Excel, PowerPoint y Outlook si están abiertos. Guarda tus documentos ahora.')
        $items.Add('Tus documentos y archivos personales no se borran.')
        $items.Add('Office se descarga desde Microsoft (unos 3 GB). Mantén el equipo conectado a Internet y a la corriente.')
        $items.Add('Verás una ventana de Microsoft con el avance. Déjala trabajar: se cierra sola.')
        if ($p.NeedsKey) { $items.Add('Al terminar, activaremos Office automáticamente con tu código SafeTech.') }
        else { $items.Add('Mientras se instala, podrás preparar tu cuenta con el tutorial guiado.') }
        if ($p.Id -eq '365' -and $state.HasAny) {
            $items.Add('Si usabas otra cuenta en Office, puede que tengas que volver a elegirla después.')
        }
    }

    $ui.StackConfirmList.Children.Clear()
    foreach ($t in $items) {
        $row = New-StXamlElement @"
<Grid Margin="0,0,0,10">
  <Grid.ColumnDefinitions><ColumnDefinition Width="26"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Ellipse Width="7" Height="7" Fill="#DC2626" VerticalAlignment="Top" Margin="0,8,0,0" HorizontalAlignment="Left"/>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="14.5" TextWrapping="Wrap" LineHeight="22"/>
</Grid>
"@
        $row.FindName('T').Text = $t
        [void]$ui.StackConfirmList.Children.Add($row)
    }

    $how = if ($p.NeedsKey) {
        @('Revisamos tu equipo y creamos un punto de restauración.', 'Quitamos versiones anteriores de Office.', "Instalamos $($p.Name) desde Microsoft.", 'Activamos Office con tu clave y tu código.')
    } else {
        @('Revisamos tu equipo y creamos un punto de restauración.', 'Quitamos versiones anteriores de Office.', 'Instalamos Microsoft 365 desde Microsoft.', 'Te guiamos para conectar tu cuenta.')
    }
    $ui.StackHowItWorks.Children.Clear()
    $i = 0
    foreach ($t in $how) {
        $i++
        $row = New-StXamlElement @"
<Grid Margin="0,0,0,12">
  <Grid.ColumnDefinitions><ColumnDefinition Width="36"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Width="26" Height="26" CornerRadius="13" Background="#26DC2626" VerticalAlignment="Top">
    <TextBlock Text="$i" Foreground="#FCA5A5" FontWeight="Bold" FontSize="12.5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
  </Border>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="13.5" TextWrapping="Wrap" Margin="0,3,0,0"/>
</Grid>
"@
        $row.FindName('T').Text = $t
        [void]$ui.StackHowItWorks.Children.Add($row)
    }
    Show-StPage 'Confirm'
}

function Start-StInstallFlow {
    $p = $St.Product
    Initialize-StLogs
    Write-StLog "Producto elegido: $($p.FullName)"
    $ids = @('check', 'restore', 'download', 'clean', 'remove', 'install', 'verify')
    if ($p.NeedsKey) { $ids += @('act_key', 'act_iid', 'act_cid', 'act_confirm') } else { $ids += 'account' }
    Initialize-StSteps $ids
    $St.UI.TxtProgKicker.Text = $(if ($p.NeedsKey) { "PASO 3 DE 3 · $($p.Name.ToUpper())" } else { "PASO 2 DE 2 · $($p.Name.ToUpper())" })
    $St.UI.TxtProgTitle.Text = "Instalando $($p.Name)"
    $St.UI.TxtProgDetail.Text = 'Preparando...'
    $St.Sync.Detail = 'Preparando...'
    $St.UI.ProgTrack.Visibility = 'Visible'
    $St.UI.PanelTutorialPromo.Visibility = $(if ($p.Id -eq '365') { 'Visible' } else { 'Collapsed' })
    $St.UI.ChipInstall.Visibility = $(if ($p.Id -eq '365') { 'Visible' } else { 'Collapsed' })
    $St.StartTime = Get-Date
    $St.Busy = $true
    $St.InstallDone = $false
    Show-StPage 'Progress'

    $work = Join-Path $StConfig.DataDir 'trabajo'
    Start-StAsync -Params @{ ProductId = $p.Id; WorkDir = $work } -Work {
        Invoke-StInstall -ProductId $P.ProductId -WorkDir $P.WorkDir
    } -OnDone {
        param($ok, $r, $err)
        $St.Busy = $false
        Update-StProgressUi
        if (-not $ok) {
            foreach ($id in $St.StepOrder) { if ($St.Sync.Steps[$id] -eq 'running') { Set-StUiStep $id 'error' } }
            Show-StError -Title "No pudimos instalar $($St.Product.Name)" -Detail $err -OnRetry { Show-StConfirmPage }
            return
        }
        $St.InstallDone = $true
        if ($St.Product.NeedsKey) {
            $choice = Show-StDialog -Kind 'success' -Kicker 'Instalación lista' -Title "$($St.Product.Name) quedó instalado" `
                -Message 'Ahora lo activaremos automáticamente con tu clave y tu código SafeTech. No necesitas abrir Word ni escribir números.' `
                -Buttons @('Activar ahora')
            Start-StActivationFlow -KeepSteps
        } else {
            Stop-StProgressAnimation
            Set-StUiStep 'account' 'running' 'Microsoft 365 quedó instalado. Ahora conecta tu cuenta con el tutorial.'
            $St.UI.TxtChipInstall.Text = 'Office instalado'
            $St.UI.BtnTutProgress.Visibility = 'Collapsed'
            $St.UI.BtnTutHome.Visibility = 'Visible'
            $conv = New-Object System.Windows.Media.BrushConverter
            $St.UI.ChipInstall.Background = $conv.ConvertFromString('#1F059669')
            $St.UI.ChipInstall.BorderBrush = $conv.ConvertFromString('#66059669')
            $St.UI.TxtChipInstall.Foreground = $conv.ConvertFromString('#6EE7B7')
            $msg = 'Ahora conecta la cuenta que te entregó SafeTech. Te mostraremos qué hacer en cada paso.'
            if ($St.Tut.Started) { $msg = 'Ya puedes continuar el tutorial: abre Word e inicia sesión con tu cuenta.' }
            [void](Show-StDialog -Kind 'success' -Kicker 'Instalación lista' -Title 'Microsoft 365 quedó instalado' -Message $msg -Buttons @('Continuar con el tutorial'))
            $target = $St.Tut.Index
            if ($St.Tut.Started -and ($St.Tut.Index -ge 6 -or ($St.Tut.Auth -eq $false -and $St.Tut.Index -ge 3))) { $target = 7 }
            Show-StTutorial -Index $target
        }
    }
}

# ----- Flujo: activacion asistida -----------------------------------------------------------------
function Start-StActivationFlow {
    param([switch]$KeepSteps)
    Initialize-StLogs
    $p = $St.Product
    if (-not $St.Act.Key) { $St.Act.Key = $St.UI.InKey.Text }
    if (-not $St.Act.Code) { $St.Act.Code = ConvertTo-StAccessCode $St.UI.InCode.Text }
    $St.Act.Iid = $null
    $St.Act.Cid = $null
    if (-not $KeepSteps) {
        Initialize-StSteps @('act_key', 'act_iid', 'act_cid', 'act_confirm')
        $St.StartTime = Get-Date
        $St.UI.TxtProgKicker.Text = "PASO 2 DE 2 · $($p.Name.ToUpper())"
        $St.UI.PanelTutorialPromo.Visibility = 'Collapsed'
    }
    $St.UI.TxtProgTitle.Text = "Activando $($p.Name)"
    $St.UI.ProgTrack.Visibility = 'Visible'
    $St.Busy = $true
    Show-StPage 'Progress'
    Set-StUiStep 'act_key' 'running' 'Registrando tu clave de producto en Office...'
    Start-StAsync -Params @{ ProductId = $p.Id; Key = $St.Act.Key } -Work {
        Invoke-StActivateKey -ProductId $P.ProductId -Key $P.Key
    } -OnDone {
        param($ok, $r, $err)
        if (-not $ok) { Stop-StActivation 'act_key' 'No pudimos registrar tu clave' $err; return }
        switch ($r.Result) {
            'licensed' {
                foreach ($id in @('act_key', 'act_iid', 'act_cid', 'act_confirm')) { Set-StUiStep $id 'done' }
                Complete-StActivation -Online
            }
            'error' {
                Set-StUiStep 'act_key' 'error' $r.Message
                $choice = Show-StDialog -Kind 'error' -Title 'No pudimos registrar tu clave' -Message $r.Message `
                    -Buttons @('Corregir la clave', 'Cerrar') -Primary 'Corregir la clave'
                $St.Busy = $false
                if ($choice -eq 'Corregir la clave') {
                    $St.Mode = 'activate'
                    $St.Act.Key = $null
                    Show-StDataPage
                } else {
                    Show-StError -Title 'Office quedó instalado, pero no activado' -Detail ($r.Message + "`n`nPuedes volver a abrir el instalador y elegir «Solo activar Office 2021 o 2024» cuando tengas la clave correcta.") -RetryLabel 'Corregir la clave' -OnRetry { $St.Mode = 'activate'; $St.Act.Key = $null; Show-StDataPage }
                }
            }
            default {
                Set-StUiStep 'act_key' 'done'
                Set-StUiStep 'act_iid' 'running' 'Obteniendo el ID de instalación de Office...'
                Start-StAsync -Params @{ ProductId = $St.Product.Id; Key = $St.Act.Key } -Work {
                    Get-StInstallationId -ProductId $P.ProductId -Key $P.Key
                } -OnDone {
                    param($ok, $iid, $err)
                    if (-not $ok -or -not $iid) { Stop-StActivation 'act_iid' 'No pudimos obtener el ID de instalación' $err; return }
                    $St.Act.Iid = [string]$iid
                    Set-StUiStep 'act_iid' 'done'
                    Request-StCidFlow
                }
            }
        }
    }
}

function Request-StCidFlow {
    if (-not $St.Act.Code) { Show-StManualActivation; return }
    $masked = Hide-StSecret $St.Act.Code
    Set-StUiStep 'act_cid' 'running' "Generando el ID de confirmación con tu código $masked..."
    Start-StAsync -Params @{ Code = $St.Act.Code; Iid = $St.Act.Iid } -Work {
        Request-StCid -Code $P.Code -Iid $P.Iid
    } -OnDone {
        param($ok, $r, $err)
        if ($ok -and $r.Ok) {
            $St.Act.Cid = $r.Cid
            Set-StUiStep 'act_cid' 'done'
            Confirm-StFlow
            return
        }
        $message = if ($ok) { $r.Message } else { $err }
        $kind = if ($ok) { $r.Kind } else { 'network' }
        Set-StUiStep 'act_cid' 'error' $message
        $buttons = if ($kind -eq 'code') { @('Usar otro código', 'Activar manualmente', 'Cancelar') } else { @('Reintentar', 'Activar manualmente', 'Cancelar') }
        $choice = Show-StDialog -Kind 'warn' -Title 'No pudimos generar el ID de confirmación' -Message $message -Buttons $buttons
        switch ($choice) {
            'Usar otro código' {
                $newCode = $null
                $hint = 'Lo encuentras en el correo de tu compra.'
                while (-not $newCode) {
                    $f = New-StDialogField -Label 'Código de activación SafeTech' -Placeholder 'ST-XXXXXX' -MaxLength 12 -Hint $hint
                    $again = Show-StDialog -Kind 'key' -Title 'Ingresa otro código' -Message '' -Extra $f.Panel -Buttons @('Usar este código', 'Cancelar')
                    if ($again -ne 'Usar este código') { break }
                    $newCode = ConvertTo-StAccessCode $f.Box.Text
                    $hint = 'Ese código está incompleto. Debe tener el formato ST-XXXXXX.'
                }
                if ($newCode) { $St.Act.Code = $newCode; Request-StCidFlow } else { Stop-StActivationPending }
            }
            'Reintentar' { Request-StCidFlow }
            'Activar manualmente' { Show-StManualActivation }
            default { Stop-StActivationPending }
        }
    }
}

function Confirm-StFlow {
    Set-StUiStep 'act_confirm' 'running' 'Confirmando la activación con Microsoft...'
    Start-StAsync -Params @{ ProductId = $St.Product.Id; Key = $St.Act.Key; Cid = $St.Act.Cid } -Work {
        Confirm-StActivation -ProductId $P.ProductId -Key $P.Key -Cid $P.Cid
    } -OnDone {
        param($ok, $r, $err)
        if ($ok -and $r.Ok) { Set-StUiStep 'act_confirm' 'done'; Complete-StActivation; return }
        $message = if ($ok) { $r.Message } else { $err }
        Set-StUiStep 'act_confirm' 'error' $message
        $choice = Show-StDialog -Kind 'error' -Title 'Microsoft no confirmó la activación' `
            -Message ($message + "`n`nPuedes intentarlo de nuevo (se generará un ID nuevo y puede usar otro crédito) o activar manualmente.") `
            -Buttons @('Intentar de nuevo', 'Activar manualmente', 'Cancelar')
        switch ($choice) {
            'Intentar de nuevo' { Start-StActivationFlow -KeepSteps }
            'Activar manualmente' { Show-StManualActivation }
            default { Stop-StActivationPending }
        }
    }
}

$StAboutSections = @(
    @{ Icon = [char]0xE72E; Color = '#6EE7B7'; Title = 'No es un instalador pirata'
       Text = 'Este instalador **no usa cracks, activadores KMS ni programas que modifiquen Windows u Office**. Solo instala y activa Office con tu licencia, igual que lo haría un técnico a mano. Su código es público y cualquiera puede revisarlo en GitHub.' },
    @{ Icon = [char]0xE896; Color = '#93C5FD'; Title = 'Office se descarga directo desde Microsoft'
       Text = 'Usamos la **Herramienta de implementación de Office (ODT)**, la misma que usan empresas y técnicos para instalar Office. Se descarga desde los servidores de Microsoft y el instalador **comprueba su firma digital de Microsoft** antes de usarla. Office llega **original, sin archivos modificados**.' },
    @{ Icon = [char]0xE8D7; Color = '#FCA5A5'; Title = 'Se activa con los métodos oficiales de Microsoft'
       Text = '**Office 2021 y 2024:** se registra tu clave de producto con la herramienta de licencias que trae Office. Primero se intenta la activación en línea y, si la clave lo requiere, la **activación por teléfono de Microsoft**: Office genera un ID de instalación y se ingresa un ID de confirmación. En vez de que tú llames y dictes números, el instalador obtiene ese ID con tu código SafeTech y lo ingresa por ti. Microsoft valida el ID antes de activar.' + "`n`n" + '**Microsoft 365:** inicias sesión en los servidores de Microsoft con la cuenta que te entregamos; el tutorial te guía paso a paso.' },
    @{ Icon = [char]0xE8B7; Color = '#FCD34D'; Title = 'Qué cambia en tu computador'
       Text = 'Antes de empezar se crea un **punto de restauración de Windows**. Se desinstalan las versiones anteriores de Office para evitar conflictos. **Tus documentos y archivos personales no se tocan.** Al terminar no queda ningún programa de SafeTech funcionando en segundo plano; solo un registro del proceso en C:\ProgramData\SafeTech.' },
    @{ Icon = [char]0xE7EF; Color = '#AEBDD0'; Title = '¿Por qué pide permisos de administrador?'
       Text = 'Instalar o desinstalar Office y registrar su licencia siempre requiere permisos de administrador, también cuando lo haces a mano. Por eso Windows te pregunta si permites los cambios.' }
)

function Show-StAbout {
    $panel = New-StXamlElement @'
<ScrollViewer MaxHeight="430" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,6,0">
  <StackPanel x:Name="List"/>
</ScrollViewer>
'@
    $list = $panel.FindName('List')
    foreach ($sec in $StAboutSections) {
        $row = New-StXamlElement @"
<Border CornerRadius="14" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="1" Padding="14,12" Margin="0,0,0,10">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="40"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="30" Height="30" CornerRadius="9" Background="#20304A" VerticalAlignment="Top">
      <TextBlock x:Name="Ico" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="14" Foreground="$($sec.Color)" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <StackPanel Grid.Column="1">
      <TextBlock x:Name="Ttl" Foreground="#F8FAFC" FontSize="15" FontWeight="Bold" TextWrapping="Wrap" Margin="0,4,0,6"/>
      <TextBlock x:Name="Txt" Foreground="#DBE5F1" FontSize="13.5" TextWrapping="Wrap" LineHeight="21"/>
    </StackPanel>
  </Grid>
</Border>
"@
        $row.FindName('Ico').Text = [string]$sec.Icon
        $row.FindName('Ttl').Text = $sec.Title
        Set-StRichText -Target $row.FindName('Txt') -Text $sec.Text
        [void]$list.Children.Add($row)
    }
    $note = New-Object System.Windows.Controls.TextBlock
    $note.TextWrapping = 'Wrap'
    $note.FontSize = 12
    $note.Foreground = '#7F93AE'
    $note.Margin = '2,4,0,0'
    $note.Text = 'SafeTech SpA es un vendedor independiente de licencias y servicio técnico en Antofagasta, Chile; no es distribuidor oficial de Microsoft. Microsoft, Office y Microsoft 365 son marcas de Microsoft Corporation.'
    [void]$list.Children.Add($note)
    $r = Show-StDialog -Kind 'info' -Kicker 'Transparencia' -Title '¿Cómo funciona este instalador?' -Message '' -Extra $panel -Buttons @('Entendido', 'Ver el código en GitHub')
    if ($St.SelfTest) { return $r }
    if ($r -eq 'Ver el código en GitHub') { Open-StUnelevated "https://github.com/$($StConfig.Repo)" }
}

function Show-StManualActivation {
    param([string]$ErrorText)
    $panel = New-Object System.Windows.Controls.StackPanel
    $groups = Format-StDigitGroups -Digits $St.Act.Iid -Size 7
    $iidBox = New-StXamlElement @"
<Border CornerRadius="14" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="1" Padding="16,12">
  <StackPanel>
    <TextBlock Text="1. TU ID DE INSTALACI&#xD3;N" Foreground="#FCA5A5" FontSize="11" FontWeight="Black" Margin="0,0,0,6"/>
    <TextBox x:Name="Iid" IsReadOnly="True" Background="#00000000" BorderThickness="0" Foreground="#F8FAFC" FontFamily="Consolas" FontSize="15.5" TextWrapping="Wrap"/>
  </StackPanel>
</Border>
"@
    $iidBox.FindName('Iid').Text = $groups
    [void]$panel.Children.Add($iidBox)
    $btnRow = New-Object System.Windows.Controls.WrapPanel
    $btnRow.Margin = '0,10,0,0'
    $bCopy = New-Object System.Windows.Controls.Button
    $bCopy.Content = 'Copiar ID'
    $bCopy.Style = $St.Window.FindResource('BtnSoft')
    $bCopy.Margin = '0,0,8,0'
    $bCopy.Tag = $St.Act.Iid
    $bCopy.Add_Click({ param($s, $e) try { [System.Windows.Clipboard]::SetText([string]$s.Tag); $s.Content = 'Copiado' } catch { } })
    $bOpen = New-Object System.Windows.Controls.Button
    $bOpen.Content = 'Abrir getcid.safe-tech.cl'
    $bOpen.Style = $St.Window.FindResource('BtnSoft')
    $bOpen.Add_Click({ Open-StUnelevated $StConfig.ApiBase })
    [void]$btnRow.Children.Add($bCopy)
    [void]$btnRow.Children.Add($bOpen)
    [void]$panel.Children.Add($btnRow)
    $f = New-StDialogField -Label '2. Pega aquí el ID de confirmación (48 números)' -Placeholder '000000 000000 000000 000000 000000 000000 000000 000000' -MaxLength 80 `
        -Hint 'En getcid.safe-tech.cl ingresa tu código y el ID de instalación; usa «Copiar código completo» y pégalo aquí.'
    $f.Box.Style = $St.Window.FindResource('Field')
    $f.Box.FontSize = 15
    $f.Panel.Margin = '0,18,0,0'
    [void]$panel.Children.Add($f.Panel)
    $msg = 'Si el código automático no funcionó, puedes obtener el ID de confirmación tú mismo en el portal de SafeTech.'
    if ($ErrorText) { $msg = $ErrorText }
    $choice = Show-StDialog -Kind 'key' -Kicker 'Activación manual' -Title 'Activar con el portal de SafeTech' -Message $msg -Extra $panel -Buttons @('Activar', 'Cancelar')
    if ($choice -ne 'Activar') { Stop-StActivationPending; return }
    $digits = $f.Box.Text -replace '\D', ''
    if ($digits.Length -ne 48) {
        Show-StManualActivation -ErrorText "El ID de confirmación debe tener 48 números y escribiste $($digits.Length). Revisa que lo hayas copiado completo."
        return
    }
    $St.Act.Cid = $digits
    Set-StUiStep 'act_cid' 'done'
    Confirm-StFlow
}

function Stop-StActivation {
    param([string]$StepId, [string]$Title, [string]$Detail)
    Set-StUiStep $StepId 'error'
    Show-StError -Title $Title -Detail $Detail -RetryLabel 'Reintentar la activación' -OnRetry { Start-StActivationFlow }
}

function Stop-StActivationPending {
    Show-StError -Title "$($St.Product.Name) quedó instalado, pero falta activarlo" `
        -Detail "Puedes reintentar ahora o más tarde: abre el instalador y elige «Solo activar Office 2021 o 2024».`n`nSi necesitas ayuda, escríbenos por WhatsApp al $($StConfig.WhatsApp)." `
        -RetryLabel 'Reintentar la activación' -OnRetry { Start-StActivationFlow }
}

function Complete-StActivation {
    param([switch]$Online)
    $St.Busy = $false
    Stop-StProgressAnimation
    $p = $St.Product
    Write-StLog "$($p.Name) activado correctamente."
    $sub = if ($Online) { 'La activación se completó en línea.' } else { 'La activación por teléfono se completó automáticamente.' }
    [void](Show-StDialog -Kind 'success' -Kicker 'Activación completa' -Title "¡$($p.Name) quedó activado!" -Message "$sub Ya puedes usar Word, Excel, PowerPoint y los demás programas." -Buttons @('Continuar'))
    $St.UI.TxtDoneTitle.Text = "¡$($p.Name) quedó listo!"
    $St.UI.TxtDoneText.Text = "Office está instalado y activado en este computador."
    Set-StDoneList @("$($p.FullName) instalado", 'Licencia activada con tu clave', 'Actualizaciones automáticas de Microsoft activadas')
    $St.UI.TxtDoneNote.Text = 'La activación queda asociada a este computador. Si formateas o reinstalas Windows, se pierde. Guarda el correo de tu compra.'
    $St.UI.DoneNote.Visibility = 'Visible'
    Show-StPage 'Done'
}

function Set-StDoneList {
    param([string[]]$Items)
    $St.UI.StackDoneList.Children.Clear()
    foreach ($t in $Items) {
        $row = New-StXamlElement @"
<Grid Margin="0,4">
  <Grid.ColumnDefinitions><ColumnDefinition Width="32"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Width="22" Height="22" CornerRadius="11" Background="#26059669">
    <TextBlock Text="&#xE73E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="11" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/>
  </Border>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="14.5" VerticalAlignment="Center" TextWrapping="Wrap"/>
</Grid>
"@
        $row.FindName('T').Text = $t
        [void]$St.UI.StackDoneList.Children.Add($row)
    }
}

function Start-StActivateOnly {
    $state = Get-StOfficeState
    $found = @()
    foreach ($id in @('2024', '2021')) { if (Test-StProductInstalled -ProductId $id -State $state) { $found += $id } }
    if ($found.Count -eq 0) {
        $choice = Show-StDialog -Kind 'question' -Title 'No encontramos Office 2021 ni 2024' `
            -Message 'Para activar, primero debe estar instalado Office 2021 u Office 2024. ¿Cuál quieres instalar?' `
            -Buttons @('Instalar Office 2024', 'Instalar Office 2021', 'Cancelar')
        if ($choice -eq 'Instalar Office 2024') { Select-StProduct '2024' }
        elseif ($choice -eq 'Instalar Office 2021') { Select-StProduct '2021' }
        return
    }
    $pick = $found[0]
    if ($found.Count -gt 1) {
        $choice = Show-StDialog -Kind 'question' -Title '¿Qué Office quieres activar?' -Message 'Encontramos Office 2024 y Office 2021 instalados.' -Buttons @('Office 2024', 'Office 2021')
        $pick = $(if ($choice -eq 'Office 2021') { '2021' } else { '2024' })
    }
    $St.Product = Get-StProduct $pick
    $St.Mode = 'activate'
    Show-StDataPage
}

# ----- Tutorial de Microsoft 365 -------------------------------------------------------------------
$StTutorial = @(
    # Marcado del texto: **negrita**; una línea que empieza con "!" es un paso que el cliente no se puede saltar.
    # NoteKind: info (consejo), warn (importante), critical (muy importante).
    @{ Short = 'Prepárate'; Title = 'Ten a mano tu cuenta y tu celular'
       Body = @('El **correo y la contraseña de Microsoft 365** que te envió SafeTech.',
                'Tu **celular con Internet**, por si Microsoft te pide instalar la app Microsoft Authenticator (algunas cuentas lo piden y otras no).',
                '!Papel o una nota para **guardar la contraseña nueva** que vas a crear.')
       Note = 'Si este computador es de tu **empresa, colegio o universidad**, consulta antes con quien lo administra: la cuenta puede entrar en conflicto con sus políticas.'; NoteKind = 'warn'
       Image = $null; Actions = @(); NeedsOffice = $false },
    @{ Short = 'Iniciar sesión'; Title = 'Inicia sesión en el portal de Microsoft 365'
       Body = @('Haz clic en **«Abrir el portal»**. Se abrirá tu navegador.',
                'Escribe el **correo entregado por SafeTech** y presiona **Siguiente**.',
                'Escribe la **contraseña temporal** que te enviamos.',
                '!**No uses tu correo personal** en este paso.')
       Note = 'Si en el navegador ya tienes abierta otra cuenta de Microsoft, abre una **ventana privada (Ctrl + Mayús + N)**, pega el enlace y entra desde ahí.'; NoteKind = 'info'
       Image = 'assets/tutorial/02-portal-inicio-sesion.jpg'; Actions = @('portal', 'copyPortal'); NeedsOffice = $false },
    @{ Short = 'Nueva contraseña'; Title = 'Cambia la contraseña temporal'
       Body = @('La primera vez, Microsoft te pedirá **crear una contraseña nueva**.',
                'En el primer campo escribe la **contraseña temporal**; en los otros dos, **tu contraseña nueva**.',
                'Usa mayúsculas, minúsculas, números y un símbolo. Ejemplo: **MiClaveSegura123!**')
       Note = '**Anota y guarda tu nueva contraseña ahora.** SafeTech no puede verla ni recuperarla después; si la pierdes, pierdes el acceso a la cuenta.'; NoteKind = 'critical'
       Image = 'assets/tutorial/03-cambiar-contrasena.jpg'; Actions = @(); NeedsOffice = $false },
    @{ Short = '¿Te pidió seguridad?'; Title = '¿Te apareció este mensaje?'
       Body = @('Después de cambiar la contraseña, algunas cuentas muestran **«Vamos a garantizar la seguridad de la cuenta»** y piden instalar Microsoft Authenticator.',
                'Otras cuentas no lo piden y entran directo. Depende de la cuenta que te entregamos.',
                '!**Elige lo que te pasó a ti** con los botones de abajo.')
       Note = 'Las dos situaciones son normales. Si no te apareció, **tu cuenta está bien** y te saltas esta parte.'; NoteKind = 'info'
       Image = 'assets/tutorial/04-garantizar-seguridad.jpg'; Actions = @('authYes', 'authNo'); NeedsOffice = $false },
    @{ Short = 'Authenticator'; Title = 'Instala Microsoft Authenticator en tu celular'
       Body = @('!En el mensaje «Vamos a garantizar la seguridad de la cuenta», haz clic en **Siguiente**. **No elijas «Usar otra cuenta».**',
                'En tu celular abre **Google Play** (Android) o **App Store** (iPhone).',
                'Busca **«Microsoft Authenticator»** e instálala. El desarrollador debe ser **Microsoft Corporation**.',
                'Vuelve al computador y haz clic en **Siguiente**.')
       Note = 'No uses la opción «Configurar una aplicación de autenticación diferente». Este tutorial es para **Microsoft Authenticator**.'; NoteKind = 'info'
       Image = 'assets/tutorial/05-instalar-authenticator.jpg'; Actions = @(); NeedsOffice = $false },
    @{ Short = 'Código QR'; Title = 'Agrega la cuenta y escanea el código QR'
       Body = @('Abre Microsoft Authenticator y **permite las notificaciones**.',
                '!Toca **«Agregar cuenta»** (o el botón +) y elige **«Cuenta profesional o educativa»**, no «Cuenta personal».',
                'Elige **«Escanear un código QR»** y apunta la cámara al código que aparece en el computador.',
                'Cuando la cuenta aparezca en tu celular, haz clic en **Siguiente** en el computador.')
       Note = '¿No puedes escanearlo? Haz clic en **«Can''t scan the QR code?»** y escribe a mano el código y la URL que aparecen.'; NoteKind = 'info'
       Image = 'assets/tutorial/06-codigo-qr.jpg'; Actions = @(); NeedsOffice = $false },
    @{ Short = 'Aprobar'; Title = 'Aprueba la notificación de prueba'
       Body = @('En el computador aparecerá **un número**.',
                '!En tu celular abre la notificación de Authenticator, **escribe ese número** y toca **«Sí»** o **«Aprobar»**.',
                'Cuando el computador diga **«Notificación aprobada»**, haz clic en **Siguiente** y luego en **Listo**.',
                'Si Microsoft pide otro método (por ejemplo, tu número de teléfono), complétalo con el código que te llegue por SMS.')
       Note = '**No borres la app ni la cuenta de tu celular**: Microsoft puede pedirte aprobar el ingreso otra vez. Si cambias de celular, **avísanos antes por WhatsApp**.'; NoteKind = 'critical'
       Image = $null; Actions = @(); NeedsOffice = $false },
    @{ Short = 'Abrir Word'; Title = 'Inicia sesión en Word'
       Body = @('Haz clic en **«Abrir Word»**.',
                '!Si Word te pide iniciar sesión, escribe el **correo entregado por SafeTech** y **tu contraseña nueva**.',
                'Si no lo pide, ve a **Archivo > Cuenta > Iniciar sesión** (o «Cambiar cuenta»).',
                'Si Microsoft lo solicita, aprueba el ingreso en Authenticator con el número que aparece.')
       Note = 'No es necesario cerrar tu cuenta personal de OneDrive o Teams.'; NoteKind = 'info'
       Image = 'assets/tutorial/08-iniciar-sesion-word.jpg'; Actions = @('word'); NeedsOffice = $true },
    @{ Short = 'Solo esta app'; Title = 'Elige «No, solo esta aplicación»'
       Body = @('Puede aparecer una ventana que pregunta si quieres iniciar sesión en **todas las aplicaciones del dispositivo**.',
                '!Haz clic en **«No, solo esta aplicación»**.')
       Note = '**No elijas «Sí».** La cuenta quedaría ligada al computador como cuenta de empresa y cambiaría su configuración. Esta cuenta solo debe usarse para activar Office.'; NoteKind = 'critical'
       Image = 'assets/tutorial/09-solo-esta-aplicacion.jpg'; Actions = @(); NeedsOffice = $true },
    @{ Short = 'Comprobar'; Title = 'Comprueba que Office quedó activado'
       Body = @('En Word ve a **Archivo > Cuenta**.',
                '!Debe decir **«Producto de suscripción»** y **«Aplicaciones de Microsoft 365»**.',
                'Si aparece «Producto sin licencia», **cierra todos los programas de Office**, vuelve a abrir Word y revisa otra vez.')
       Note = '**No guardes archivos importantes en el OneDrive de esta cuenta**: Microsoft puede reducir o eliminar ese espacio sin aviso. Usa tu OneDrive personal.'; NoteKind = 'warn'
       Image = 'assets/tutorial/10-producto-activado.jpg'; Actions = @('check', 'word'); NeedsOffice = $true }
)

function Initialize-StTutorialNav {
    $St.UI.StackTutSteps.Children.Clear()
    $St.TutRows = @()
    $i = 0
    foreach ($s in $StTutorial) {
        $i++
        $row = New-StXamlElement @"
<Button Cursor="Hand" Margin="0,0,0,3" FocusVisualStyle="{x:Null}" Background="#00000000" BorderThickness="0" HorizontalContentAlignment="Stretch">
  <Button.Template>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="12" Padding="10,8" Background="{TemplateBinding Background}"><ContentPresenter/></Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#17243A"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Button.Template>
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="34"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Grid Width="24" Height="24">
      <Ellipse x:Name="Dot" Stroke="#2A3C57" StrokeThickness="1.5"/>
      <TextBlock x:Name="Num" Text="$i" FontSize="11.5" FontWeight="Bold" Foreground="#AEBDD0" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Grid>
    <TextBlock Grid.Column="1" x:Name="Lbl" Foreground="#AEBDD0" FontSize="13.5" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
  </Grid>
</Button>
"@
        $row.FindName('Lbl').Text = $s.Short
        $row.Tag = $i - 1
        $row.Add_Click({
            param($sender, $e)
            $target = [int]$sender.Tag
            if (Test-StAuthSkipped $target) { $St.Tut.Auth = $true }
            Set-StTutorialStep -Index $target
        })
        [void]$St.UI.StackTutSteps.Children.Add($row)
        $St.TutRows += $row
    }
}

function Show-StTutorial {
    param([int]$Index = 0)
    if (-not $St.TutRows) { Initialize-StTutorialNav }
    $St.Tut.Started = $true
    $installing = Test-StInstalling365
    $St.UI.BtnTutProgress.Visibility = $(if ($installing) { 'Visible' } else { 'Collapsed' })
    $St.UI.BtnTutHome.Visibility = $(if ($installing) { 'Collapsed' } else { 'Visible' })
    Show-StPage 'Tutorial'
    Set-StTutorialStep -Index $Index
}

function Test-StInstalling365 {
    return [bool]($St.Mode -eq 'install' -and $St.Product -and $St.Product.Id -eq '365' -and $St.Busy)
}

function Set-StRichText {
    # Escribe el texto en un TextBlock; lo que va entre **dobles asteriscos** queda en negrita y blanco.
    param($Target, [string]$Text)
    $Target.Inlines.Clear()
    if (-not $Text) { return }
    $parts = [regex]::Split($Text, '\*\*')
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($parts[$i] -eq '') { continue }
        $run = New-Object System.Windows.Documents.Run($parts[$i])
        if ($i % 2 -eq 1) {
            $run.FontWeight = [System.Windows.FontWeights]::Bold
            $run.Foreground = [System.Windows.Media.Brushes]::White
        }
        $Target.Inlines.Add($run)
    }
}

function Test-StAuthSkipped {
    param([int]$Index)
    return ($St.Tut.Auth -eq $false -and $Index -ge 4 -and $Index -le 6)
}

function Set-StTutorialAuth {
    param([bool]$Shown)
    $St.Tut.Auth = $Shown
    if ($Shown) { Set-StTutorialStep -Index 4; return }
    [void](Show-StDialog -Kind 'success' -Kicker 'Todo bien' -Title 'Tu cuenta no necesita Authenticator' `
        -Message 'Algunas cuentas piden Microsoft Authenticator y otras no; depende de la cuenta. Que no te haya aparecido es normal y no afecta la activación de Office. Nos saltamos esa parte.' `
        -Buttons @('Continuar'))
    if (Test-StTutorialBlocked 7) { Set-StTutorialStep -Index 3 } else { Set-StTutorialStep -Index 7 }
}

function Test-StTutorialBlocked {
    param([int]$Index)
    return ($StTutorial[$Index].NeedsOffice -and $St.Mode -eq 'install' -and $St.Product -and $St.Product.Id -eq '365' -and -not $St.InstallDone)
}

function Set-StTutorialStep {
    param([int]$Index)
    if ($Index -lt 0) { $Index = 0 }
    if ($Index -ge $StTutorial.Count) { $Index = $StTutorial.Count - 1 }
    if (Test-StTutorialBlocked $Index) {
        [void](Show-StDialog -Kind 'info' -Kicker 'UN MOMENTO' -Title 'Office todavía se está instalando' `
            -Message 'Este paso se hace en Word. Cuando la instalación termine te avisaremos y podrás continuar desde aquí.' -Buttons @('Entendido'))
        return
    }
    $St.Tut.Index = $Index
    $s = $StTutorial[$Index]
    $ui = $St.UI
    $conv = New-Object System.Windows.Media.BrushConverter
    $n = $StTutorial.Count
    $ui.TxtTutKicker.Text = "PASO $($Index + 1) DE $n"
    $ui.TxtTutTitle.Text = $s.Title
    $ui.StackTutBody.Children.Clear()
    $k = 0
    foreach ($line in $s.Body) {
        $k++
        $must = $line.StartsWith('!')
        $text = $line.TrimStart('!')
        if ($must) {
            $row = New-StXamlElement @"
<Border CornerRadius="12" Background="#1FDC2626" BorderBrush="#66DC2626" BorderThickness="1" Padding="10,9" Margin="-10,0,0,10">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="34"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="24" Height="24" CornerRadius="8" Background="#DC2626" VerticalAlignment="Top">
      <TextBlock Text="$k" FontSize="12" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <TextBlock Grid.Column="1" x:Name="T" Foreground="#F8FAFC" FontSize="15" TextWrapping="Wrap" LineHeight="23" Margin="0,1,0,0"/>
  </Grid>
</Border>
"@
        } else {
            $row = New-StXamlElement @"
<Grid Margin="0,0,0,10">
  <Grid.ColumnDefinitions><ColumnDefinition Width="34"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Width="24" Height="24" CornerRadius="8" Background="#20304A" VerticalAlignment="Top">
    <TextBlock Text="$k" FontSize="12" FontWeight="Bold" Foreground="#FCA5A5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
  </Border>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="15" TextWrapping="Wrap" LineHeight="23" Margin="0,1,0,0"/>
</Grid>
"@
        }
        Set-StRichText -Target $row.FindName('T') -Text $text
        [void]$ui.StackTutBody.Children.Add($row)
    }
    $noteText = $s.Note
    $noteKind = $s.NoteKind
    if ($Index -eq 0 -and $St.SignInIssues -and $St.SignInIssues.Count -gt 0) {
        $noteText = $s.Note + "`n`n" + 'Revisamos tu Windows y encontramos algo que podría impedir el inicio de sesión: ' + ($St.SignInIssues -join ' ') + ' Si Word no te deja iniciar sesión, escríbenos.'
    } elseif ($Index -eq 3 -and $St.Tut.Auth -eq $false) {
        $noteText = 'Respondiste que no te apareció: **está todo bien**. Presiona «Siguiente» para continuar en Word.'
    }
    Set-StRichText -Target $ui.TxtTutNote -Text $noteText
    $noteMap = @{
        info     = @('#140891B2', '#4D0891B2', '#67E8F9', [char]0xE946, 'CONSEJO', '1')
        warn     = @('#1AF59E0B', '#80F59E0B', '#FCD34D', [char]0xE7BA, 'IMPORTANTE', '1.5')
        critical = @('#26DC2626', '#B3DC2626', '#FCA5A5', [char]0xE7BA, 'MUY IMPORTANTE · LÉELO', '2')
    }[$noteKind]
    $ui.TutNote.Background = $conv.ConvertFromString($noteMap[0])
    $ui.TutNote.BorderBrush = $conv.ConvertFromString($noteMap[1])
    $ui.TutNote.BorderThickness = $noteMap[5]
    $ui.TutNoteIcon.Foreground = $conv.ConvertFromString($noteMap[2])
    $ui.TutNoteIcon.Text = [string]$noteMap[3]
    $ui.TxtTutNoteLabel.Text = $noteMap[4]
    $ui.TxtTutNoteLabel.Foreground = $conv.ConvertFromString($noteMap[2])
    $ui.TxtTutNote.Foreground = $conv.ConvertFromString($(if ($noteKind -eq 'info') { '#DBE5F1' } else { '#F8FAFC' }))
    $ui.TutNote.Visibility = $(if ($s.Note) { 'Visible' } else { 'Collapsed' })

    $ui.StackTutActions.Children.Clear()
    foreach ($a in $s.Actions) {
        $b = New-Object System.Windows.Controls.Button
        $b.Margin = '0,0,10,10'
        $b.Tag = $a
        switch ($a) {
            'portal' { $b.Content = 'Abrir el portal de Microsoft 365'; $b.Style = $St.Window.FindResource('BtnPrimary') }
            'copyPortal' { $b.Content = 'Copiar el enlace'; $b.Style = $St.Window.FindResource('BtnSoft') }
            'word' { $b.Content = 'Abrir Word'; $b.Style = $St.Window.FindResource($(if ($s.Actions[0] -eq 'word') { 'BtnPrimary' } else { 'BtnSoft' })) }
            'check' { $b.Content = 'Comprobar la activación'; $b.Style = $St.Window.FindResource('BtnPrimary') }
            'authYes' { $b.Content = 'Sí, me apareció'; $b.Style = $St.Window.FindResource($(if ($St.Tut.Auth -eq $false) { 'BtnSoft' } else { 'BtnPrimary' })) }
            'authNo' { $b.Content = 'No me apareció'; $b.Style = $St.Window.FindResource($(if ($St.Tut.Auth -eq $false) { 'BtnPrimary' } else { 'BtnSoft' })) }
        }
        $b.Add_Click({ param($sender, $e) Invoke-StTutorialAction ([string]$sender.Tag) $sender })
        [void]$ui.StackTutActions.Children.Add($b)
    }
    $ui.StackTutActions.Visibility = $(if ($s.Actions.Count) { 'Visible' } else { 'Collapsed' })

    $img = $null
    if ($s.Image) { $img = Get-StImage $s.Image }
    $ui.ImgTut.Source = $img
    $ui.TutImageFrame.Visibility = $(if ($img) { 'Visible' } else { 'Collapsed' })

    for ($i = 0; $i -lt $St.TutRows.Count; $i++) {
        $r = $St.TutRows[$i]
        $dot = $r.FindName('Dot'); $num = $r.FindName('Num'); $lbl = $r.FindName('Lbl')
        if ($i -eq $Index) {
            $r.Background = $conv.ConvertFromString('#20304A')
            $dot.Fill = $St.Window.FindResource('RedGrad'); $dot.Stroke = $conv.ConvertFromString('#DC2626')
            $num.Foreground = $conv.ConvertFromString('#FFFFFF'); $num.Text = [string]($i + 1); $num.FontFamily = 'Segoe UI'
            $lbl.Foreground = $conv.ConvertFromString('#F8FAFC'); $lbl.FontWeight = 'Bold'
        } elseif ($i -lt $Index) {
            $r.Background = $conv.ConvertFromString('#00000000')
            $dot.Fill = $conv.ConvertFromString('#059669'); $dot.Stroke = $conv.ConvertFromString('#059669')
            $num.Foreground = $conv.ConvertFromString('#FFFFFF'); $num.Text = [string][char]0xE73E; $num.FontFamily = 'Segoe Fluent Icons, Segoe MDL2 Assets'
            $lbl.Foreground = $conv.ConvertFromString('#DBE5F1'); $lbl.FontWeight = 'Normal'
        } else {
            $r.Background = $conv.ConvertFromString('#00000000')
            $dot.Fill = $conv.ConvertFromString('#00000000'); $dot.Stroke = $conv.ConvertFromString('#2A3C57')
            $num.Foreground = $conv.ConvertFromString('#AEBDD0'); $num.Text = [string]($i + 1); $num.FontFamily = 'Segoe UI'
            $lbl.Foreground = $conv.ConvertFromString('#AEBDD0'); $lbl.FontWeight = 'Normal'
            if (Test-StTutorialBlocked $i) { $lbl.Foreground = $conv.ConvertFromString('#5D7391') }
        }
        if ($i -ne $Index -and (Test-StAuthSkipped $i)) {
            $dot.Fill = $conv.ConvertFromString('#00000000'); $dot.Stroke = $conv.ConvertFromString('#2A3C57')
            $num.Foreground = $conv.ConvertFromString('#5D7391'); $num.Text = [string][char]0xE738; $num.FontFamily = 'Segoe Fluent Icons, Segoe MDL2 Assets'
            $lbl.Foreground = $conv.ConvertFromString('#5D7391'); $lbl.FontWeight = 'Normal'
        }
    }
    # En el primer paso, «Atrás» vuelve a la pantalla principal (salvo mientras se instala Microsoft 365).
    $ui.BtnTutBack.IsEnabled = ($Index -gt 0 -or -not (Test-StInstalling365))
    $ui.TxtTutNext.Text = $(if ($Index -eq $n - 1) { 'Terminar' } else { 'Siguiente' })
    $ui.TxtTutProgress.Text = "Paso $($Index + 1) de $n · $([Math]::Round((($Index + 1) / $n) * 100))%"
    $trackWidth = $ui.TutProgFill.Parent.ActualWidth
    if ($trackWidth -le 0) { $trackWidth = 520 }
    $target = [double]($trackWidth * (($Index + 1) / $n))
    if ($St.NoAnim) { $ui.TutProgFill.Width = $target }
    else {
        $from = [double]$ui.TutProgFill.ActualWidth
        $anim = [System.Windows.Media.Animation.DoubleAnimation]::new($from, $target, [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(420)))
        $e = New-Object System.Windows.Media.Animation.QuinticEase; $e.EasingMode = 'EaseOut'; $anim.EasingFunction = $e
        $ui.TutProgFill.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $anim)
    }
    $ui.TutScroll.ScrollToTop()
    if (-not $St.NoAnim) {
        $ui.TutContent.Opacity = 0
        $fade = [System.Windows.Media.Animation.DoubleAnimation]::new([double]0, [double]1, [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(320)))
        $ui.TutContent.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fade)
    }
}

function Invoke-StTutorialAction {
    param([string]$Action, $Button)
    switch ($Action) {
        'portal' { Open-StUnelevated $StConfig.PortalUrl }
        'copyPortal' { try { [System.Windows.Clipboard]::SetText($StConfig.PortalUrl); $Button.Content = 'Enlace copiado' } catch { } }
        'word' { [void](Open-StWord) }
        'authYes' { Set-StTutorialAuth $true }
        'authNo' { Set-StTutorialAuth $false }
        'check' {
            if (-not (Test-StProductInstalled -ProductId '365')) {
                [void](Show-StDialog -Kind 'warn' -Title 'No encontramos Microsoft 365' -Message 'Microsoft 365 no aparece instalado en este computador.' -Buttons @('Entendido'))
            } elseif (Test-StM365License) {
                [void](Show-StDialog -Kind 'success' -Title 'Encontramos una licencia de Microsoft 365' -Message 'Word ya recibió una licencia en este equipo. Para confirmarlo, en Word ve a Archivo > Cuenta: debe decir «Producto de suscripción».' -Buttons @('Perfecto'))
            } else {
                [void](Show-StDialog -Kind 'warn' -Title 'Todavía no vemos la licencia' -Message 'Asegúrate de haber iniciado sesión en Word con la cuenta de SafeTech. Luego cierra Word, vuelve a abrirlo y revisa Archivo > Cuenta.' -Buttons @('Entendido'))
            }
        }
    }
}

function Complete-StTutorial {
    $St.UI.TxtDoneTitle.Text = '¡Microsoft 365 quedó listo!'
    $St.UI.TxtDoneText.Text = 'Ya puedes usar Word, Excel, PowerPoint y Outlook con la cuenta que te entregamos.'
    if ($St.Tut.Auth -eq $false) {
        Set-StDoneList @('Aplicaciones de Microsoft 365 instaladas', 'Cuenta de SafeTech conectada en Office', 'Contraseña nueva creada')
        $St.UI.TxtDoneNote.Text = 'Guarda tu nueva contraseña: SafeTech no puede recuperarla. La garantía dura 1 año desde la compra.'
    } else {
        Set-StDoneList @('Aplicaciones de Microsoft 365 instaladas', 'Cuenta de SafeTech conectada en Office', 'Microsoft Authenticator configurado en tu celular')
        $St.UI.TxtDoneNote.Text = 'Guarda tu nueva contraseña y no borres Microsoft Authenticator de tu celular. La garantía dura 1 año desde la compra.'
    }
    $St.UI.DoneNote.Visibility = 'Visible'
    if ($St.StepRows.ContainsKey('account')) { Set-StUiStep 'account' 'done' }
    Show-StPage 'Done'
}

# ==============================================================================
#  Construccion de la ventana
# ==============================================================================
$St.Window = [Windows.Markup.XamlReader]::Parse($StWindowXaml.Replace('__RESOURCES__', $StResXaml))
foreach ($m in [regex]::Matches($StWindowXaml, 'x:Name="([A-Za-z0-9_]+)"')) {
    $name = $m.Groups[1].Value
    $el = $St.Window.FindName($name)
    if ($el) { $St.UI[$name] = $el }
}
$St.UI.TxtVersion.Text = "SafeTech SpA · Instalador v$($StConfig.Version)"

$wa = [System.Windows.SystemParameters]::WorkArea
if ($wa.Height -lt 780) { $St.Window.Height = [Math]::Max(600, $wa.Height - 16) }
if ($wa.Width -lt 1100) { $St.Window.Width = [Math]::Max(940, $wa.Width - 16) }

$St.Pool = [runspacefactory]::CreateRunspacePool(1, 4)
$St.Pool.Open()

# Logo e icono
$logo = Get-StImage 'assets/logo-safetech-dark.png'
if ($logo) { $St.UI.ImgLogo.Source = $logo } else { $St.UI.ImgLogo.Visibility = 'Collapsed'; $St.UI.LogoFallback.Visibility = 'Visible' }
$icon = Get-StImage 'assets/icono.png'
if ($icon) { $St.Window.Icon = $icon }

# Descarga en segundo plano las imagenes del tutorial
if (-not $env:SAFETECH_OFFICE_ASSETS -and $StMode -eq 'app') {
    $assetList = @($StTutorial | Where-Object { $_.Image } | ForEach-Object { $_.Image })
    Start-StAsync -Params @{ Paths = $assetList; Cache = $St.CacheDir } -Work { Save-StAssets -Paths $P.Paths -CacheDir $P.Cache }
}
if ($StMode -eq 'app') {
    Start-StAsync -Work { Get-StSignInIssues } -OnDone { param($ok, $r, $err) if ($ok) { $St.SignInIssues = @($r) } }
}

# ----- Eventos ---------------------------------------------------------------------------------
$ui = $St.UI
$ui.BtnMin.Add_Click({ $St.Window.WindowState = 'Minimized' })
$ui.BtnMax.Add_Click({ $St.Window.WindowState = $(if ($St.Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }) })
$St.Window.Add_StateChanged({
    # Con WindowChrome la ventana maximizada sobresale unos pixeles de la pantalla: se compensa con un margen.
    if ($St.Window.WindowState -eq 'Maximized') {
        $St.UI.RootBorder.Margin = '7'
        $St.UI.RootBorder.BorderThickness = '0'
        $St.UI.BtnMax.Content = [string][char]0xE923
        $St.UI.BtnMax.ToolTip = 'Restaurar'
    } else {
        $St.UI.RootBorder.Margin = '0'
        $St.UI.RootBorder.BorderThickness = '1'
        $St.UI.BtnMax.Content = [string][char]0xE922
        $St.UI.BtnMax.ToolTip = 'Maximizar'
    }
})
$ui.BtnClose.Add_Click({ $St.Window.Close() })
$ui.BtnHelpCenter.Add_Click({ Open-StUnelevated $StConfig.HelpUrl })
$ui.BtnAbout.Add_Click({ Show-StAbout })
$ui.BtnAboutHome.Add_Click({ Show-StAbout })
$ui.BtnHelpWa.Add_Click({ Open-StUnelevated $StConfig.WhatsAppUrl })

$ui.BtnPick2024.Add_Click({ Select-StProduct '2024' })
$ui.BtnPick2021.Add_Click({ Select-StProduct '2021' })
$ui.BtnPick365.Add_Click({ Select-StProduct '365' })
$ui.BtnOnlyActivate.Add_Click({ Start-StActivateOnly })
$ui.BtnOnlyTutorial.Add_Click({
    $St.Product = Get-StProduct '365'
    $St.Mode = 'tutorial'
    if (-not (Test-StProductInstalled -ProductId '365')) {
        $choice = Show-StDialog -Kind 'question' -Title 'No encontramos Microsoft 365 instalado' `
            -Message 'Puedes instalarlo ahora (te guiaremos con la cuenta mientras se instala) o ver el tutorial de todas formas.' `
            -Buttons @('Instalar Microsoft 365', 'Ver el tutorial', 'Cancelar')
        if ($choice -eq 'Instalar Microsoft 365') { Select-StProduct '365'; return }
        if ($choice -ne 'Ver el tutorial') { return }
    }
    Show-StTutorial -Index 0
})

$ui.InKey.Add_TextChanged({
    if ($St.KeyFormatting) { return }
    $St.KeyFormatting = $true
    try {
        $fmt = Format-StProductKey $St.UI.InKey.Text
        if ($fmt -ne $St.UI.InKey.Text) {
            $St.UI.InKey.Text = $fmt
            $St.UI.InKey.CaretIndex = $fmt.Length
        }
    } finally { $St.KeyFormatting = $false }
    Update-StDataValidity
})
$ui.InCode.Add_TextChanged({
    $t = $St.UI.InCode.Text
    $up = $t.ToUpperInvariant()
    if ($up -ne $t) { $caret = $St.UI.InCode.CaretIndex; $St.UI.InCode.Text = $up; $St.UI.InCode.CaretIndex = $caret }
    $St.CodeState = 'typing'
    $St.CodeTimer.Stop()
    $St.CodeTimer.Start()
    Update-StDataValidity
})
$St.CodeTimer = New-Object System.Windows.Threading.DispatcherTimer
$St.CodeTimer.Interval = [TimeSpan]::FromMilliseconds(450)
$St.CodeTimer.Add_Tick({ $St.CodeTimer.Stop(); Start-StCodeCheck })

$ui.BtnDataBack.Add_Click({ Show-StPage 'Home' })
$ui.BtnDataNext.Add_Click({
    $St.Act = @{ Key = $St.UI.InKey.Text; Code = (ConvertTo-StAccessCode $St.UI.InCode.Text) }
    if ($St.Mode -eq 'activate') { Start-StActivationFlow } else { Show-StConfirmPage }
})

$ui.ChkConfirm.Add_Checked({ $St.UI.BtnConfirmStart.IsEnabled = $true })
$ui.ChkConfirm.Add_Unchecked({ if ($St.ConfirmAction -ne 'tutorial') { $St.UI.BtnConfirmStart.IsEnabled = $false } })
$ui.BtnConfirmBack.Add_Click({ if ($St.Product.NeedsKey) { Show-StPage 'Data' } else { Show-StPage 'Home' } })
$ui.BtnConfirmAlt.Add_Click({ $St.Mode = 'activate'; Start-StActivationFlow })
$ui.BtnConfirmStart.Add_Click({
    if ($St.ConfirmAction -eq 'tutorial') { $St.Mode = 'tutorial'; Show-StTutorial -Index 0 } else { Start-StInstallFlow }
})

$ui.BtnTutorialEarly.Add_Click({ Show-StTutorial -Index $St.Tut.Index })
$ui.BtnTutProgress.Add_Click({ Show-StPage 'Progress' })
$ui.BtnTutBack.Add_Click({
    if ($St.Tut.Index -le 0) { if (-not (Test-StInstalling365)) { Show-StPage 'Home' } }
    elseif ($St.Tut.Index -eq 7 -and $St.Tut.Auth -eq $false) { Set-StTutorialStep -Index 3 }
    else { Set-StTutorialStep -Index ($St.Tut.Index - 1) }
})
$ui.BtnTutHome.Add_Click({ if (-not (Test-StInstalling365)) { Show-StPage 'Home' } })
$ui.BtnTutNext.Add_Click({
    $i = $St.Tut.Index
    if ($i -ge $StTutorial.Count - 1) { Complete-StTutorial }
    elseif ($i -eq 3 -and $null -eq $St.Tut.Auth) {
        $c = Show-StDialog -Kind 'question' -Kicker 'Una pregunta' -Title '¿Te apareció «Vamos a garantizar la seguridad de la cuenta»?' `
            -Message 'Algunas cuentas lo muestran después de cambiar la contraseña y otras no. Las dos cosas son normales.' `
            -Buttons @('Sí, me apareció', 'No me apareció')
        if ($c -eq 'Sí, me apareció') { Set-StTutorialAuth $true } elseif ($c -eq 'No me apareció') { Set-StTutorialAuth $false }
    }
    elseif ($i -eq 3 -and $St.Tut.Auth -eq $false) { Set-StTutorialStep -Index 7 }
    else { Set-StTutorialStep -Index ($i + 1) }
})

$ui.BtnDoneWord.Add_Click({ [void](Open-StWord) })
$ui.BtnDoneClose.Add_Click({ $St.Window.Close() })
$ui.BtnErrHome.Add_Click({ $St.Busy = $false; Show-StPage 'Home' })
$ui.BtnErrRetry.Add_Click({ if ($St.ErrRetry) { & $St.ErrRetry } })
$ui.BtnErrZip.Add_Click({
    Initialize-StLogs
    $St.UI.BtnErrZip.IsEnabled = $false
    Start-StAsync -Params @{ Dir = $St.LogDir } -Work { New-StDiagnosticZip -LogDir $P.Dir } -OnDone {
        param($ok, $zip, $err)
        $St.UI.BtnErrZip.IsEnabled = $true
        if ($ok) {
            $c = Show-StDialog -Kind 'success' -Title 'ZIP de diagnóstico creado' -Message "Lo dejamos en tu Escritorio:`n$zip`n`nEnvíalo por WhatsApp al $($StConfig.WhatsApp) o a $($StConfig.Email)." -Buttons @('Mostrar el archivo', 'Cerrar')
            if ($c -eq 'Mostrar el archivo') { Start-Process (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList "/select,`"$zip`"" }
        } else {
            [void](Show-StDialog -Kind 'error' -Title 'No pudimos crear el ZIP' -Message "$err`n`nLos registros están en: $($St.LogDir)" -Buttons @('Cerrar'))
        }
    }
})

$St.Window.Add_Closing({
    param($s, $e)
    if ($St.Busy -and -not $St.SelfTest) {
        $c = Show-StDialog -Kind 'warn' -Title 'La instalación sigue en curso' `
            -Message 'Si cierras el instalador ahora, Office podría quedar a medio instalar o sin activar. ¿Quieres esperar a que termine?' `
            -Buttons @('Seguir esperando', 'Cerrar de todas formas')
        if ($c -ne 'Cerrar de todas formas') { $e.Cancel = $true }
    }
})
$St.Window.Add_Closed({ try { $St.Timer.Stop() } catch { } })

$St.Timer = New-Object System.Windows.Threading.DispatcherTimer
$St.Timer.Interval = [TimeSpan]::FromMilliseconds(200)
$St.Timer.Add_Tick({
    try { Update-StProgressUi } catch { }
    if (-not $St.DialogOpen) { Invoke-StJobPump }
})
$St.Timer.Start()

# ==============================================================================
#  Autoprueba (CI): dibuja cada pantalla en PNG sin instalar nada.
# ==============================================================================
if ($StMode -eq 'selftest') {
    $St.SelfTest = $true
    $St.NoAnim = $true
    $out = if ($env:SAFETECH_OFFICE_OUT) { $env:SAFETECH_OFFICE_OUT } else { Join-Path $env:TEMP 'safetech-selftest' }
    New-Item -ItemType Directory -Force -Path $out | Out-Null
    $St.Window.WindowStartupLocation = 'Manual'
    $St.Window.Left = 0; $St.Window.Top = 0
    $St.Window.Width = 1080; $St.Window.Height = 740
    $St.Window.ShowActivated = $false
    $St.Window.Show()
    function Save-StShot {
        param($Visual, [string]$Name)
        $Visual.UpdateLayout()
        $St.Window.Dispatcher.Invoke([action] { }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
        $w = [int][Math]::Ceiling($Visual.ActualWidth); $h = [int][Math]::Ceiling($Visual.ActualHeight)
        $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap $w, $h, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
        $rtb.Render($Visual)
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $fs = [IO.File]::Create((Join-Path $out "$Name.png"))
        $enc.Save($fs); $fs.Close()
        Write-Host "Captura: $Name ($w x $h)"
    }
    $root = $St.Window.Content
    Show-StPage 'Home'; Save-StShot $root '01-inicio'
    $St.Window.Width = 1500; $St.Window.Height = 940
    Save-StShot $root '01b-inicio-grande'
    $St.Window.Width = 1080; $St.Window.Height = 740
    $St.Product = Get-StProduct '2024'; $St.Mode = 'install'
    Show-StDataPage
    $St.UI.InKey.Text = 'ABCDE12345FGHIJ67890KLMNO'
    $St.UI.InCode.Text = 'ST-AB12CD'
    $St.CodeTimer.Stop()
    Set-StCodeChip 'valid' 'Válido · 3 créditos'
    Save-StShot $root '02-datos'
    Show-StConfirmPage
    $St.UI.ChkConfirm.IsChecked = $true
    Save-StShot $root '03-confirmar'
    Initialize-StSteps @('check', 'restore', 'download', 'clean', 'remove', 'install', 'verify', 'act_key', 'act_iid', 'act_cid', 'act_confirm')
    foreach ($id in @('check', 'restore', 'download', 'clean')) { Set-StUiStep $id 'done' }
    Set-StUiStep 'remove' 'skip'
    Set-StUiStep 'install' 'running' 'Instalando Office 2024. Se abrió la ventana de Microsoft con el avance; déjala trabajar, se cierra sola.'
    $St.UI.TxtProgTitle.Text = 'Instalando Office 2024'
    $St.UI.TxtElapsed.Text = '07:42'
    Show-StPage 'Progress'; Save-StShot $root '04-avance'
    $St.Product = Get-StProduct '365'
    $St.UI.PanelTutorialPromo.Visibility = 'Visible'
    Save-StShot $root '05-avance-365'
    $St.Mode = 'tutorial'
    Show-StTutorial -Index 1; Save-StShot $root '06-tutorial-paso2'
    Set-StTutorialStep -Index 2; Save-StShot $root '06a-tutorial-contrasena'
    Set-StTutorialStep -Index 3; Save-StShot $root '06b-tutorial-pregunta'
    $St.Tut.Auth = $false
    Set-StTutorialStep -Index 8; Save-StShot $root '07-tutorial-paso9'
    $St.Product = Get-StProduct '2024'
    $St.UI.TxtDoneTitle.Text = '¡Office 2024 quedó listo!'
    $St.UI.TxtDoneText.Text = 'Office está instalado y activado en este computador.'
    Set-StDoneList @('Office LTSC Profesional Plus 2024 instalado', 'Licencia activada con tu clave', 'Actualizaciones automáticas de Microsoft activadas')
    $St.UI.TxtDoneNote.Text = 'La activación queda asociada a este computador. Si formateas o reinstalas Windows, se pierde.'
    Show-StPage 'Done'; Save-StShot $root '08-listo'
    $St.ErrRetry = { }
    Show-StError -Title 'No pudimos instalar Office 2024' -Detail 'No pudimos conectarnos con los servidores de Microsoft. Revisa tu conexión a Internet.' -OnRetry { }
    Save-StShot $root '09-error'
    $St.Act = @{ Iid = '123456712345671234567123456712345671234567123456712345671234567' }
    $dlg = Show-StDialog -Kind 'success' -Kicker 'Instalación lista' -Title 'Office 2024 quedó instalado' -Message 'Ahora lo activaremos automáticamente con tu clave y tu código SafeTech. No necesitas abrir Word ni escribir números.' -Buttons @('Activar ahora')
    $dlg.Show(); Save-StShot $dlg '10-ventana-emergente'; $dlg.Close()
    $f = New-StDialogField -Label 'Código de activación SafeTech' -Placeholder 'ST-XXXXXX' -MaxLength 12 -Hint 'Lo encuentras en el correo de tu compra.'
    $f.Box.Style = $St.Window.FindResource('Field')
    $dlg = Show-StDialog -Kind 'warn' -Title 'No pudimos generar el ID de confirmación' -Message 'Código no válido. Verifica que esté escrito correctamente.' -Extra $f.Panel -Buttons @('Usar otro código', 'Activar manualmente', 'Cancelar')
    $dlg.Show(); Save-StShot $dlg '11-ventana-codigo'; $dlg.Close()
    $dlg = Show-StAbout
    $dlg.Show(); Save-StShot $dlg '12-como-funciona'; $dlg.Close()
    $St.Window.Close()
    Write-Host "Autoprueba completa: $out"
    return
}

# ==============================================================================
#  Inicio
# ==============================================================================
try {
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        Add-Type -Namespace StWin -Name Native -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -ErrorAction SilentlyContinue
        [void][StWin.Native]::ShowWindow($hwnd, 0)
    }
} catch { }

Show-StPage 'Home'
[void]$St.Window.ShowDialog()
if ($env:SAFETECH_OFFICE_CHILD -eq '1') { [Environment]::Exit(0) }
