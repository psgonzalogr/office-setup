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
    Version     = '1.2.0'
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
if ($StMode -eq 'probe') {
    # Usado por las pruebas del lanzador .cmd: confirma que el script se descargo y se ejecuto.
    Write-Output 'SAFETECH-PROBE-OK'
    return
}

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
            Write-Host '  Listo: el instalador se esta abriendo en una ventana nueva.' -ForegroundColor Green
            Write-Host ''
            # Esta consola ya no hace falta: se cierra sola a los 10 segundos para no confundir al cliente.
            for ($stLeft = 10; $stLeft -ge 1; $stLeft--) {
                Write-Host ("`r  Esta ventana se cerrara sola en {0,2} segundos..." -f $stLeft) -NoNewline -ForegroundColor Gray
                Start-Sleep -Seconds 1
            }
            [Environment]::Exit(0)
        } catch {
            Write-Host '  No se otorgaron los permisos de administrador.' -ForegroundColor Yellow
            Write-Host '  Vuelve a ejecutar el comando y responde "Si" cuando Windows lo pregunte.' -ForegroundColor Yellow
        }
        Write-Host ''
        if ($env:SAFETECH_OFFICE_LAUNCHER -eq 'cmd') { exit 1 }
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

function New-StAsyncScript {
    # Arma el texto que corre en segundo plano: motor + trabajo, y siempre devuelve @{ Ok; Result; Error }.
    param([string]$EngineText, [string]$WorkText)
    return ("param(`$P, `$sync, `$StConfig)`n" + $EngineText + "`n" + @'
try {
    $__out = & {
'@ + "`n" + $WorkText + "`n" + @'
    }
    Write-StLog 'Tarea terminada.'
    @{ Ok = $true; Result = $__out; Error = $null }
} catch {
    Write-StLog ("ERROR: " + $_.Exception.Message)
    @{ Ok = $false; Result = $null; Error = $_.Exception.Message }
}
'@)
}

function Get-StAsyncResult {
    # Busca el @{ Ok; Result; Error } entre lo que devolvio la tarea.
    # Ojo: no usar ".BaseObject" sobre el elemento: en una tabla hash PowerShell lo busca como clave y da $null
    # (ese era el fallo "La tarea termino sin respuesta" de las versiones 1.0.0 a 1.0.6).
    param($Output)
    $list = @($Output)
    for ($i = $list.Count - 1; $i -ge 0; $i--) {
        $item = $list[$i]
        if ($item -is [hashtable] -and $item.ContainsKey('Ok')) { return $item }
    }
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

function Get-StScale([double]$Width, [double]$Height) {
    # Diseñado para 1280x760 con letra de 17 px. En pantallas chicas se achica un poco;
    # en las grandes NO se agranda (a 1920x1080 la letra quedaba excesivamente grande).
    $scale = [Math]::Min($Width / 1280, $Height / 760)
    return [Math]::Max(0.8, [Math]::Min(1.0, $scale))
}

function Find-StPurchaseData {
    # Busca la clave y el código ST dentro de un texto pegado (por ejemplo, un trozo del correo).
    param([string]$Text)
    $r = @{ Key = $null; Code = $null }
    if (-not $Text) { return $r }
    $up = $Text.ToUpperInvariant()
    $mk = [regex]::Match($up, '(?<![A-Z0-9])[A-Z0-9]{5}(-[A-Z0-9]{5}){4}(?![A-Z0-9])')
    if ($mk.Success) { $r.Key = $mk.Value }
    $mc = [regex]::Match($up, '(?<![A-Z0-9])ST\s*-?\s*([A-Z0-9]{6})(?![A-Z0-9])')
    if ($mc.Success) { $r.Code = 'ST-' + $mc.Groups[1].Value }
    return $r
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

function Get-StInstallHint {
    # Lo que se le dice al cliente mientras corre la ventana de Microsoft.
    param([string]$ProductId)
    $prod = Get-StProduct $ProductId
    $t = 'Se abrió la ventana de Microsoft con el avance. Si al final dice «¡Hemos terminado!», presiona «Cerrar» (o espera: la cerramos por ti).'
    if ($prod.NeedsKey) { $t += " Aunque diga «Microsoft 365», se está instalando tu $($prod.Name)." }
    return $t
}

function Close-StOdtWindows {
    # Al terminar, Microsoft deja abierta su ventana «¡Hemos terminado!» aunque setup.exe ya salió,
    # y tapa la ventana del instalador. Se cierra igual que con su botón «Cerrar» (WM_CLOSE).
    # Solo toca ventanas de los programas de Click-to-Run, nunca Word, Excel, etc.
    param([int]$WaitSeconds = 20)
    $roots = @()
    foreach ($base in @($env:CommonProgramFiles, ${env:CommonProgramFiles(x86)}, $env:CommonProgramW6432)) {
        if ($base) { $roots += (Join-Path $base 'microsoft shared\ClickToRun') }
    }
    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    $closed = 0
    do {
        foreach ($p in @(Get-Process -ErrorAction SilentlyContinue)) {
            if ($p.MainWindowHandle -eq [IntPtr]::Zero) { continue }
            $path = $null
            try { $path = $p.Path } catch { }
            if (-not $path) { continue }
            $match = $false
            foreach ($r in $roots) { if ($path.StartsWith($r, [StringComparison]::OrdinalIgnoreCase)) { $match = $true } }
            if (-not $match) { continue }
            Write-StLog ("Cerrando la ventana final de Microsoft: {0} «{1}»" -f $p.ProcessName, $p.MainWindowTitle)
            try { if ($p.CloseMainWindow()) { $closed++ } } catch { }
        }
        if ($closed -gt 0) { break }
        Start-Sleep -Milliseconds 700
    } while ((Get-Date) -lt $deadline)
    if ($closed -eq 0) { Write-StLog 'No apareció una ventana final de Microsoft para cerrar.' }
    return $closed
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
    $exit = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    $text = (@($out) -join "`n")
    Write-StLog ("OSPP {0} (salida {1}, {2} lineas, {3} caracteres)`n{4}" -f (Hide-StSecret ($Arguments -join ' ')), $exit, @($out).Count, $text.Length, (Hide-StSecret $text))
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

    Set-StStep 'install' 'running' (Get-StInstallHint -ProductId $ProductId)
    Start-StClickToRunService
    $code = Invoke-StOdt -Setup $setup -Xml $installXml -Label "Instalacion de $($prod.FullName)"
    if ($code -ne 0) {
        throw ("{0} (código {1})`n`nRevisa tu conexión a Internet y vuelve a intentarlo. Si se repite, crea el ZIP de diagnóstico y envíalo a SafeTech." -f (Get-StOdtMessage $code), $code)
    }
    Set-StStep 'install' 'running' 'Microsoft terminó. Cerrando su ventana y comprobando...'
    [void](Close-StOdtWindows)
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

# ----- Maximizar sin tapar la barra de tareas ----------------------------------------------------
# La ventana no usa el marco de Windows (WindowStyle None + WindowChrome). Asi, al maximizar,
# Windows le da el monitor completo y la barra de tareas queda tapada. Este gancho responde a
# WM_GETMINMAXINFO con el area de trabajo del monitor donde esta la ventana.
$StMaximizeFixCode = @'
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

public static class SafeTechMaximizeFix {
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct MINMAXINFO {
        public POINT Reserved; public POINT MaxSize; public POINT MaxPosition; public POINT MinTrackSize; public POINT MaxTrackSize;
    }
    [StructLayout(LayoutKind.Sequential)] public struct MONITORINFO {
        public int Size; public RECT Monitor; public RECT Work; public int Flags;
    }
    [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr hwnd, int flags);
    [DllImport("user32.dll")] static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFO info);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

    public static RECT GetWorkArea(IntPtr hwnd) {
        MONITORINFO mi = new MONITORINFO();
        mi.Size = Marshal.SizeOf(typeof(MONITORINFO));
        GetMonitorInfo(MonitorFromWindow(hwnd, 2), ref mi);
        return mi.Work;
    }

    static IntPtr Hook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
        if (msg != 0x0024) { return IntPtr.Zero; }
        MONITORINFO mi = new MONITORINFO();
        mi.Size = Marshal.SizeOf(typeof(MONITORINFO));
        if (!GetMonitorInfo(MonitorFromWindow(hwnd, 2), ref mi)) { return IntPtr.Zero; }
        MINMAXINFO mmi = (MINMAXINFO)Marshal.PtrToStructure(lParam, typeof(MINMAXINFO));
        mmi.MaxPosition.X = mi.Work.Left - mi.Monitor.Left;
        mmi.MaxPosition.Y = mi.Work.Top - mi.Monitor.Top;
        mmi.MaxSize.X = mi.Work.Right - mi.Work.Left;
        mmi.MaxSize.Y = mi.Work.Bottom - mi.Work.Top;
        Marshal.StructureToPtr(mmi, lParam, true);
        handled = true;
        return IntPtr.Zero;
    }

    public static void Attach(Window window) {
        IntPtr hwnd = new WindowInteropHelper(window).EnsureHandle();
        HwndSource source = HwndSource.FromHwnd(hwnd);
        if (source != null) { source.AddHook(Hook); }
    }
}
'@

function Enable-StMaximizeFix {
    # Devuelve $true si el gancho quedo puesto. Si falla (por ejemplo, sin compilador), la app sigue igual.
    param([Parameter(Mandatory)]$Window)
    try {
        if (-not ('SafeTechMaximizeFix' -as [type])) {
            Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
            Add-Type -TypeDefinition $StMaximizeFixCode -ReferencedAssemblies PresentationFramework, PresentationCore, WindowsBase, System.Xaml -ErrorAction Stop
        }
        [SafeTechMaximizeFix]::Attach($Window)
        return $true
    } catch {
        Write-StLog "No se pudo ajustar el maximizado: $($_.Exception.Message)"
        return $false
    }
}

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
    Tut         = @{ Index = 0; Started = $false; Auth = $null; Checked = @{} }
    CodeState   = 'empty'
    KeyFormatting = $false
    CacheDir    = (Join-Path $env:TEMP 'SafeTechOffice\cache')
    LogDir      = $null
    SoundOn     = $true
    CuesDone    = @{}
    VoicePlayer = $null
    PendingVoice = $null
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
<SolidColorBrush x:Key="Focus" Color="#FCD34D"/>
<LinearGradientBrush x:Key="RedGrad" StartPoint="0,0" EndPoint="1,1">
  <GradientStop Color="#DC2626" Offset="0"/><GradientStop Color="#991B1B" Offset="1"/>
</LinearGradientBrush>
<LinearGradientBrush x:Key="RedGradHover" StartPoint="0,0" EndPoint="1,1">
  <GradientStop Color="#E23434" Offset="0"/><GradientStop Color="#A11D1D" Offset="1"/>
</LinearGradientBrush>
<LinearGradientBrush x:Key="StubGrad" StartPoint="0,0" EndPoint="0,1">
  <GradientStop Color="#DC2626" Offset="0"/><GradientStop Color="#8E1818" Offset="1"/>
</LinearGradientBrush>
<LinearGradientBrush x:Key="TopLine" StartPoint="0,0" EndPoint="1,0">
  <GradientStop Color="#00DC2626" Offset="0"/><GradientStop Color="#DC2626" Offset="0.5"/><GradientStop Color="#00DC2626" Offset="1"/>
</LinearGradientBrush>
<FontFamily x:Key="Icons">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>

<Style x:Key="H1" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="34"/>
  <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="TextWrapping" Value="Wrap"/>
  <Setter Property="FontFamily" Value="Segoe UI Variable Display, Segoe UI"/>
</Style>
<Style x:Key="H2" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="22"/>
  <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="TextWrapping" Value="Wrap"/>
  <Setter Property="FontFamily" Value="Segoe UI Variable Display, Segoe UI"/>
</Style>
<Style x:Key="Lead" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#C9D5E3"/><Setter Property="FontSize" Value="18"/>
  <Setter Property="TextWrapping" Value="Wrap"/><Setter Property="LineHeight" Value="27"/>
  <Setter Property="Margin" Value="0,8,0,0"/>
</Style>
<Style x:Key="Body" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#DBE5F1"/><Setter Property="FontSize" Value="17"/>
  <Setter Property="TextWrapping" Value="Wrap"/><Setter Property="LineHeight" Value="26"/>
</Style>
<Style x:Key="Small" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#AEBDD0"/><Setter Property="FontSize" Value="16"/>
  <Setter Property="TextWrapping" Value="Wrap"/><Setter Property="LineHeight" Value="22"/>
</Style>
<Style x:Key="Label" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="17"/>
  <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Margin" Value="0,0,0,8"/>
</Style>
<Style x:Key="Kicker" TargetType="TextBlock">
  <Setter Property="Foreground" Value="#AEBDD0"/><Setter Property="FontSize" Value="15"/>
  <Setter Property="Margin" Value="0,0,0,6"/>
</Style>
<Style x:Key="Panel" TargetType="Border">
  <Setter Property="Background" Value="#111C2F"/><Setter Property="BorderBrush" Value="#2A3C57"/>
  <Setter Property="BorderThickness" Value="1"/><Setter Property="CornerRadius" Value="20"/>
  <Setter Property="Padding" Value="24"/>
</Style>

<Style x:Key="BtnPrimary" TargetType="Button">
  <Setter Property="Foreground" Value="White"/><Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="17"/><Setter Property="MinHeight" Value="56"/>
  <Setter Property="Padding" Value="28,12"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="RenderTransformOrigin" Value="0.5,0.5"/>
  <Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="1" ScaleY="1"/></Setter.Value></Setter>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Grid>
        <Border x:Name="Ring" CornerRadius="17" Margin="-5" BorderBrush="#FCD34D" BorderThickness="3" Visibility="Hidden"/>
        <Border x:Name="Bd" CornerRadius="14" Padding="{TemplateBinding Padding}" Background="{StaticResource RedGrad}">
          <Border.Effect><DropShadowEffect Color="#7F1D1D" BlurRadius="22" ShadowDepth="6" Direction="270" Opacity="0.45"/></Border.Effect>
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.Foreground="White"/>
        </Border>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource RedGradHover}"/></Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="0.97" ScaleY="0.97"/></Setter.Value></Setter></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Ring" Property="Visibility" Value="Visible"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.38"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnSoft" TargetType="Button">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="17"/><Setter Property="MinHeight" Value="56"/>
  <Setter Property="Padding" Value="24,12"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="RenderTransformOrigin" Value="0.5,0.5"/>
  <Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="1" ScaleY="1"/></Setter.Value></Setter>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Grid>
        <Border x:Name="Ring" CornerRadius="17" Margin="-5" BorderBrush="#FCD34D" BorderThickness="3" Visibility="Hidden"/>
        <Border x:Name="Bd" CornerRadius="14" Padding="{TemplateBinding Padding}" Background="#17243A" BorderBrush="#6B82A3" BorderThickness="2">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="Bd" Property="BorderBrush" Value="#AEBDD0"/><Setter TargetName="Bd" Property="Background" Value="#20304A"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="0.97" ScaleY="0.97"/></Setter.Value></Setter></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Ring" Property="Visibility" Value="Visible"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.38"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnWa" TargetType="Button">
  <Setter Property="Foreground" Value="#052E16"/><Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="17"/><Setter Property="MinHeight" Value="56"/>
  <Setter Property="Padding" Value="24,12"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Grid>
        <Border x:Name="Ring" CornerRadius="17" Margin="-5" BorderBrush="#FCD34D" BorderThickness="3" Visibility="Hidden"/>
        <Border x:Name="Bd" CornerRadius="14" Padding="{TemplateBinding Padding}" Background="#25D366">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" TextBlock.Foreground="#052E16"/>
        </Border>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#3BE07A"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Ring" Property="Visibility" Value="Visible"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.38"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnLink" TargetType="Button">
  <Setter Property="Foreground" Value="#FCA5A5"/><Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="16"/><Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Padding" Value="10,8"/><Setter Property="MinHeight" Value="48"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="10" Padding="{TemplateBinding Padding}" Background="#00000000" BorderBrush="#00000000" BorderThickness="2">
        <ContentPresenter VerticalAlignment="Center"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#26DC2626"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#FCD34D"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnChrome" TargetType="Button">
  <Setter Property="Foreground" Value="#C9D5E3"/><Setter Property="Width" Value="48"/><Setter Property="Height" Value="40"/>
  <Setter Property="FontFamily" Value="{StaticResource Icons}"/><Setter Property="FontSize" Value="12"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="10" Background="#00000000" BorderBrush="#00000000" BorderThickness="2"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#20304A"/><Setter Property="Foreground" Value="#F8FAFC"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#FCD34D"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="BtnClose" TargetType="Button" BasedOn="{StaticResource BtnChrome}">
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="10" Background="#00000000" BorderBrush="#00000000" BorderThickness="2"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#C42B1C"/><Setter Property="Foreground" Value="White"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#FCD34D"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="Ticket" TargetType="Button">
  <Setter Property="Cursor" Value="Hand"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Foreground" Value="#F8FAFC"/>
  <Setter Property="HorizontalContentAlignment" Value="Stretch"/><Setter Property="VerticalContentAlignment" Value="Stretch"/>
  <Setter Property="RenderTransform"><Setter.Value><TranslateTransform Y="0"/></Setter.Value></Setter>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="Button">
      <Grid>
        <Border x:Name="Ring" CornerRadius="25" Margin="-6" BorderBrush="#FCD34D" BorderThickness="3" Visibility="Hidden"/>
        <Border x:Name="Bd" CornerRadius="20" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="2">
          <Border.Effect><DropShadowEffect Color="#000000" BlurRadius="30" ShadowDepth="10" Direction="270" Opacity="0.35"/></Border.Effect>
          <ContentPresenter/>
        </Border>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="Bd" Property="BorderBrush" Value="#DC2626"/>
          <Setter Property="RenderTransform"><Setter.Value><TranslateTransform Y="-4"/></Setter.Value></Setter>
        </Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Ring" Property="Visibility" Value="Visible"/></Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter Property="RenderTransform"><Setter.Value><TranslateTransform Y="0"/></Setter.Value></Setter></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="Field" TargetType="TextBox">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="CaretBrush" Value="#F8FAFC"/>
  <Setter Property="SelectionBrush" Value="#DC2626"/><Setter Property="FontSize" Value="21"/>
  <Setter Property="FontFamily" Value="Consolas, Segoe UI"/><Setter Property="Height" Value="58"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="TextBox">
      <Border x:Name="Bd" CornerRadius="14" Background="#0C1626" BorderBrush="#6B82A3" BorderThickness="2">
        <Grid Margin="18,0">
          <TextBlock x:Name="Ph" Text="{TemplateBinding Tag}" Foreground="#8FA3BE" VerticalAlignment="Center" Visibility="Collapsed" FontFamily="{TemplateBinding FontFamily}" FontSize="{TemplateBinding FontSize}"/>
          <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
        </Grid>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="Text" Value=""><Setter TargetName="Ph" Property="Visibility" Value="Visible"/></Trigger>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#AEBDD0"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#FCD34D"/><Setter TargetName="Bd" Property="Background" Value="#111C2F"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style x:Key="Check" TargetType="CheckBox">
  <Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="FontSize" Value="17"/>
  <Setter Property="Cursor" Value="Hand"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="CheckBox">
      <Border x:Name="Row" CornerRadius="14" Background="#17243A" BorderBrush="#405776" BorderThickness="2" Padding="18,14" MinHeight="58">
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border x:Name="Box" Width="28" Height="28" CornerRadius="8" BorderBrush="#8FA3BE" BorderThickness="2" Background="#00000000" VerticalAlignment="Center">
            <TextBlock x:Name="Mark" Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="15" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
          </Border>
          <ContentPresenter Grid.Column="1" Margin="16,0,0,0" VerticalAlignment="Center"/>
        </Grid>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsChecked" Value="True">
          <Setter TargetName="Box" Property="Background" Value="#059669"/><Setter TargetName="Box" Property="BorderBrush" Value="#059669"/>
          <Setter TargetName="Mark" Property="Visibility" Value="Visible"/><Setter TargetName="Row" Property="BorderBrush" Value="#059669"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Row" Property="Background" Value="#20304A"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Row" Property="BorderBrush" Value="#FCD34D"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>

<Style TargetType="ScrollBar">
  <Setter Property="Width" Value="12"/><Setter Property="MinWidth" Value="12"/>
  <Setter Property="Template"><Setter.Value>
    <ControlTemplate TargetType="ScrollBar">
      <Track x:Name="PART_Track" Orientation="Vertical" IsDirectionReversed="True">
        <Track.Thumb><Thumb><Thumb.Template><ControlTemplate TargetType="Thumb"><Border CornerRadius="5" Background="#6B82A3" Margin="2,2"/></ControlTemplate></Thumb.Template></Thumb></Track.Thumb>
      </Track>
    </ControlTemplate>
  </Setter.Value></Setter>
</Style>
<Style TargetType="ToolTip">
  <Setter Property="Background" Value="#20304A"/><Setter Property="Foreground" Value="#F8FAFC"/><Setter Property="BorderBrush" Value="#405776"/><Setter Property="FontSize" Value="15"/>
</Style>
'@

$StWindowXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Instalador de Office | SafeTech" Width="1280" Height="760" MinWidth="960" MinHeight="620"
        WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="CanResize"
        Background="#09111F" Foreground="#F8FAFC" FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="17"
        UseLayoutRounding="True" SnapsToDevicePixels="True" TextOptions.TextFormattingMode="Ideal">
  <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="64" ResizeBorderThickness="7" CornerRadius="0" GlassFrameThickness="0" UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>
  <Window.Resources>
    __RESOURCES__
  </Window.Resources>
  <Border x:Name="RootBorder" BorderBrush="#2A3C57" BorderThickness="1" Background="#09111F">
  <Grid x:Name="Root">
    <Grid.LayoutTransform><ScaleTransform x:Name="RootScale" ScaleX="1" ScaleY="1"/></Grid.LayoutTransform>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="64"/>
    </Grid.RowDefinitions>

    <!-- Barra superior -->
    <Border Grid.Row="0" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="0,0,0,1">
      <Grid Margin="24,0,10,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Grid>
            <Image x:Name="ImgLogo" Height="32" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
            <StackPanel x:Name="LogoFallback" Orientation="Horizontal" Visibility="Collapsed">
              <TextBlock Text="Safe" Foreground="#DC2626" FontSize="24" FontWeight="Bold" FontStyle="Italic"/>
              <TextBlock Text="Tech" Foreground="#F8FAFC" FontSize="24" FontWeight="Black"/>
            </StackPanel>
          </Grid>
          <Border Width="1" Height="24" Background="#2A3C57" Margin="18,0"/>
          <TextBlock Text="Instalador de Office" Foreground="#DBE5F1" FontSize="17" FontWeight="SemiBold" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
          <Button x:Name="BtnSound" Style="{StaticResource BtnSoft}" MinHeight="48" Padding="16,6" FontSize="16" Margin="0,0,10,0" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Activa o apaga los sonidos y la voz de SafeBOT">
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="TxtSoundIcon" Text="&#xE767;" FontFamily="{StaticResource Icons}" FontSize="16" VerticalAlignment="Center" Margin="0,0,9,0"/>
              <TextBlock x:Name="TxtSound" Text="Sonido activado" VerticalAlignment="Center"/>
            </StackPanel>
          </Button>
          <Button x:Name="BtnAbout" Style="{StaticResource BtnSoft}" MinHeight="48" Padding="16,6" FontSize="16" Margin="0,0,14,0" WindowChrome.IsHitTestVisibleInChrome="True">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE946;" FontFamily="{StaticResource Icons}" FontSize="16" VerticalAlignment="Center" Margin="0,0,9,0"/>
              <TextBlock Text="¿Cómo funciona?" VerticalAlignment="Center"/>
            </StackPanel>
          </Button>
          <Button x:Name="BtnMin" Style="{StaticResource BtnChrome}" Content="&#xE921;" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Minimizar"/>
          <Button x:Name="BtnMax" Style="{StaticResource BtnChrome}" Content="&#xE922;" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Maximizar"/>
          <Button x:Name="BtnClose" Style="{StaticResource BtnClose}" Content="&#xE8BB;" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="Cerrar"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Guía: SafeBOT + etapas -->
    <Grid x:Name="GuideBar" Grid.Row="1" Margin="40,12,40,4" MaxWidth="1240">
      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <ContentControl x:Name="BotHost" Height="72" Width="56" VerticalAlignment="Center" Focusable="False"/>
      <Grid Grid.Column="1" Margin="10,0,24,0" VerticalAlignment="Center" HorizontalAlignment="Left">
        <Path Data="M 12,0 L 0,10 L 12,20 Z" Fill="#17243A" Stroke="#2A3C57" StrokeThickness="1.5" VerticalAlignment="Center" HorizontalAlignment="Left"/>
        <Border Margin="10,0,0,0" CornerRadius="18" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="1.5" Padding="20,12" MaxWidth="660">
          <TextBlock x:Name="TxtGuide" Foreground="#F8FAFC" FontSize="18" TextWrapping="Wrap" LineHeight="26" Text=""/>
        </Border>
      </Grid>
      <StackPanel x:Name="Journey" Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
        <StackPanel HorizontalAlignment="Center">
          <Border x:Name="JDot1" Width="36" Height="36" CornerRadius="18" BorderThickness="2" HorizontalAlignment="Center">
            <TextBlock x:Name="JNum1" Text="1" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <TextBlock x:Name="JLbl1" Text="Elegir" FontSize="16" Margin="0,6,0,0" HorizontalAlignment="Center"/>
        </StackPanel>
        <Border x:Name="JLine1" Width="44" Height="2" Margin="8,0,8,24" VerticalAlignment="Center"/>
        <StackPanel HorizontalAlignment="Center">
          <Border x:Name="JDot2" Width="36" Height="36" CornerRadius="18" BorderThickness="2" HorizontalAlignment="Center">
            <TextBlock x:Name="JNum2" Text="2" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <TextBlock x:Name="JLbl2" Text="Tus datos" FontSize="16" Margin="0,6,0,0" HorizontalAlignment="Center"/>
        </StackPanel>
        <Border x:Name="JLine2" Width="44" Height="2" Margin="8,0,8,24" VerticalAlignment="Center"/>
        <StackPanel HorizontalAlignment="Center">
          <Border x:Name="JDot3" Width="36" Height="36" CornerRadius="18" BorderThickness="2" HorizontalAlignment="Center">
            <TextBlock x:Name="JNum3" Text="3" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <TextBlock x:Name="JLbl3" Text="Instalar y activar" FontSize="16" Margin="0,6,0,0" HorizontalAlignment="Center"/>
        </StackPanel>
      </StackPanel>
    </Grid>

    <Grid Grid.Row="2" x:Name="PageHost">

      <!-- INICIO -->
      <Grid x:Name="PageHome">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid MaxWidth="1240" Margin="40,10,40,18">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <StackPanel>
              <TextBlock Style="{StaticResource H1}" Text="¿Qué compraste?"/>
              <TextBlock Style="{StaticResource Lead}" Text="Míralo en el correo de tu compra de SafeTech y elige ese mismo producto."/>
            </StackPanel>
            <UniformGrid Grid.Row="1" Columns="3" Margin="0,18,0,0">
              <Button x:Name="BtnPick2024" Style="{StaticResource Ticket}" Margin="0,0,12,0" Height="256" AutomationProperties.Name="Instalar Office 2024">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="108"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border CornerRadius="18,0,0,18" Background="{StaticResource StubGrad}">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                      <TextBlock Text="OFFICE" Foreground="#FECACA" FontSize="13" FontWeight="Bold" HorizontalAlignment="Center"/>
                      <TextBlock Text="2024" Foreground="White" FontSize="32" FontWeight="Black" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Border>
                  <Line Grid.Column="1" X1="0" Y1="16" X2="0" Y2="238" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
                  <Ellipse Grid.ColumnSpan="2" Width="24" Height="24" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="96,-12,0,0"/>
                  <Ellipse Grid.ColumnSpan="2" Width="24" Height="24" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="96,0,0,-12"/>
                  <StackPanel Grid.Column="1" Margin="24,20,18,0">
                    <TextBlock Text="Office 2024" FontSize="26" FontWeight="SemiBold" FontFamily="Segoe UI Variable Display, Segoe UI"/>
                    <TextBlock Text="Word, Excel, PowerPoint y más" Style="{StaticResource Small}" FontSize="16" Margin="0,2,0,14"/>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="24" Height="24" CornerRadius="12" Background="#26059669" VerticalAlignment="Top"><TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                      <TextBlock Grid.Column="1" Text="Licencia para siempre" FontSize="16" Foreground="#DBE5F1" Margin="10,1,0,0" TextWrapping="Wrap"/>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="24" Height="24" CornerRadius="12" Background="#26059669" VerticalAlignment="Top"><TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                      <TextBlock Grid.Column="1" Text="Se activa con tu código" FontSize="16" Foreground="#DBE5F1" Margin="10,1,0,0" TextWrapping="Wrap"/>
                    </Grid>
                  </StackPanel>
                  <Border Grid.Column="1" VerticalAlignment="Bottom" Margin="24,0,18,18" CornerRadius="12" Background="{StaticResource RedGrad}" Height="54">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                      <TextBlock Text="Instalar Office 2024" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                      <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" FontSize="13" Foreground="White" VerticalAlignment="Center" Margin="10,1,0,0"/>
                    </StackPanel>
                  </Border>
                </Grid>
              </Button>
              <Button x:Name="BtnPick2021" Style="{StaticResource Ticket}" Margin="6,0,6,0" Height="256" AutomationProperties.Name="Instalar Office 2021">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="108"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border CornerRadius="18,0,0,18" Background="{StaticResource StubGrad}">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                      <TextBlock Text="OFFICE" Foreground="#FECACA" FontSize="13" FontWeight="Bold" HorizontalAlignment="Center"/>
                      <TextBlock Text="2021" Foreground="White" FontSize="32" FontWeight="Black" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Border>
                  <Line Grid.Column="1" X1="0" Y1="16" X2="0" Y2="238" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
                  <Ellipse Grid.ColumnSpan="2" Width="24" Height="24" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="96,-12,0,0"/>
                  <Ellipse Grid.ColumnSpan="2" Width="24" Height="24" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="96,0,0,-12"/>
                  <StackPanel Grid.Column="1" Margin="24,20,18,0">
                    <TextBlock Text="Office 2021" FontSize="26" FontWeight="SemiBold" FontFamily="Segoe UI Variable Display, Segoe UI"/>
                    <TextBlock Text="Word, Excel, PowerPoint y más" Style="{StaticResource Small}" FontSize="16" Margin="0,2,0,14"/>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="24" Height="24" CornerRadius="12" Background="#26059669" VerticalAlignment="Top"><TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                      <TextBlock Grid.Column="1" Text="Licencia para siempre" FontSize="16" Foreground="#DBE5F1" Margin="10,1,0,0" TextWrapping="Wrap"/>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="24" Height="24" CornerRadius="12" Background="#26059669" VerticalAlignment="Top"><TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                      <TextBlock Grid.Column="1" Text="Se activa con tu código" FontSize="16" Foreground="#DBE5F1" Margin="10,1,0,0" TextWrapping="Wrap"/>
                    </Grid>
                  </StackPanel>
                  <Border Grid.Column="1" VerticalAlignment="Bottom" Margin="24,0,18,18" CornerRadius="12" Background="{StaticResource RedGrad}" Height="54">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                      <TextBlock Text="Instalar Office 2021" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                      <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" FontSize="13" Foreground="White" VerticalAlignment="Center" Margin="10,1,0,0"/>
                    </StackPanel>
                  </Border>
                </Grid>
              </Button>
              <Button x:Name="BtnPick365" Style="{StaticResource Ticket}" Margin="12,0,0,0" Height="256" AutomationProperties.Name="Instalar Microsoft 365">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="108"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border CornerRadius="18,0,0,18" Background="{StaticResource StubGrad}">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                      <TextBlock Text="MICROSOFT" Foreground="#FECACA" FontSize="12" FontWeight="Bold" HorizontalAlignment="Center"/>
                      <TextBlock Text="365" Foreground="White" FontSize="34" FontWeight="Black" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Border>
                  <Line Grid.Column="1" X1="0" Y1="16" X2="0" Y2="238" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
                  <Ellipse Grid.ColumnSpan="2" Width="24" Height="24" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="96,-12,0,0"/>
                  <Ellipse Grid.ColumnSpan="2" Width="24" Height="24" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="96,0,0,-12"/>
                  <StackPanel Grid.Column="1" Margin="24,20,18,0">
                    <TextBlock Text="Microsoft 365" FontSize="26" FontWeight="SemiBold" FontFamily="Segoe UI Variable Display, Segoe UI"/>
                    <TextBlock Text="Word, Excel, PowerPoint y más" Style="{StaticResource Small}" FontSize="16" Margin="0,2,0,14"/>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="24" Height="24" CornerRadius="12" Background="#26059669" VerticalAlignment="Top"><TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                      <TextBlock Grid.Column="1" Text="Con tu cuenta de SafeTech" FontSize="16" Foreground="#DBE5F1" Margin="10,1,0,0" TextWrapping="Wrap"/>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="24" Height="24" CornerRadius="12" Background="#26059669" VerticalAlignment="Top"><TextBlock Text="&#xE73E;" FontFamily="{StaticResource Icons}" FontSize="12" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                      <TextBlock Grid.Column="1" Text="Te guiamos paso a paso" FontSize="16" Foreground="#DBE5F1" Margin="10,1,0,0" TextWrapping="Wrap"/>
                    </Grid>
                  </StackPanel>
                  <Border Grid.Column="1" VerticalAlignment="Bottom" Margin="24,0,18,18" CornerRadius="12" Background="{StaticResource RedGrad}" Height="54">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                      <TextBlock Text="Instalar Microsoft 365" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                      <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" FontSize="13" Foreground="White" VerticalAlignment="Center" Margin="10,1,0,0"/>
                    </StackPanel>
                  </Border>
                </Grid>
              </Button>
            </UniformGrid>
            <Border Grid.Row="2" Style="{StaticResource Panel}" Margin="0,18,0,0" Padding="24,12">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center" Margin="0,0,20,0">
                  <TextBlock Text="¿Ya tienes Office instalado en este computador?" FontWeight="SemiBold" FontSize="18" TextWrapping="Wrap"/>
                  <TextBlock Style="{StaticResource Small}" FontSize="16" Text="Puedes activarlo o conectar tu cuenta sin volver a instalar." Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                  <Button x:Name="BtnOnlyActivate" Style="{StaticResource BtnPrimary}" MinHeight="50" Padding="20,10" Margin="0,0,12,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE8D7;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                      <TextBlock Text="Solo activar Office 2021 o 2024"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnOnlyTutorial" Style="{StaticResource BtnPrimary}" MinHeight="50" Padding="20,10">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE77B;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                      <TextBlock Text="Conectar mi cuenta 365"/>
                    </StackPanel>
                  </Button>
                </StackPanel>
              </Grid>
            </Border>
          </Grid>
        </ScrollViewer>
      </Grid>

      <!-- DATOS DE LA COMPRA -->
      <Grid x:Name="PageData" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Border MaxWidth="940" Margin="40,8,40,18" CornerRadius="22" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="2" VerticalAlignment="Top">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="118"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border CornerRadius="20,0,0,20" Background="{StaticResource StubGrad}">
                <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                  <TextBlock x:Name="TxtDataStubLabel" Text="OFFICE" Foreground="#FECACA" FontSize="13" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock x:Name="TxtDataStubYear" Text="2024" Foreground="White" FontSize="34" FontWeight="Black" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Line Grid.Column="1" X1="0" Y1="0" X2="0" Y2="1" Stretch="Fill" Margin="0,20" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
              <Ellipse Grid.ColumnSpan="2" Width="26" Height="26" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="105,-13,0,0"/>
              <Ellipse Grid.ColumnSpan="2" Width="26" Height="26" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="105,0,0,-13"/>
              <StackPanel Grid.Column="1" Margin="36,26,36,26">
                <TextBlock Style="{StaticResource H1}" FontSize="30" Text="Escribe los datos de tu compra"/>
                <TextBlock Style="{StaticResource Body}" Foreground="#C9D5E3" Margin="0,6,0,0" Text="Están en el correo que te envió SafeTech. Puedes copiarlos y pulsar «Pegar»."/>
                <TextBlock Style="{StaticResource Label}" Text="Clave de producto" Margin="0,20,0,8"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBox x:Name="InKey" Style="{StaticResource Field}" Tag="XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" MaxLength="200" AutomationProperties.Name="Clave de producto"/>
                  <Button x:Name="BtnPasteKey" Grid.Column="1" Style="{StaticResource BtnSoft}" Margin="12,0,0,0" MinHeight="58" Padding="18,8">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE77F;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                      <TextBlock Text="Pegar"/>
                    </StackPanel>
                  </Button>
                </Grid>
                <TextBlock x:Name="TxtKeyHint" Style="{StaticResource Small}" Margin="2,8,0,0" Text="25 letras y números. Los guiones se agregan solos."/>
                <TextBlock Style="{StaticResource Label}" Text="Código de activación SafeTech" Margin="0,18,0,8"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBox x:Name="InCode" Style="{StaticResource Field}" Tag="ST-XXXXXX" MaxLength="200" AutomationProperties.Name="Código de activación SafeTech"/>
                  <Border x:Name="ChipCode" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,14,0" CornerRadius="14" Height="28" Padding="12,0" Visibility="Collapsed" Background="#17243A" BorderBrush="#059669" BorderThickness="1.5">
                    <TextBlock x:Name="TxtCodeStatus" FontSize="14" FontWeight="SemiBold" Foreground="#6EE7B7" VerticalAlignment="Center" Text=""/>
                  </Border>
                  <Button x:Name="BtnPasteCode" Grid.Column="1" Style="{StaticResource BtnSoft}" Margin="12,0,0,0" MinHeight="58" Padding="18,8">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE77F;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                      <TextBlock Text="Pegar"/>
                    </StackPanel>
                  </Button>
                </Grid>
                <TextBlock Style="{StaticResource Small}" Margin="2,8,0,0" Text="Empieza con ST-, por ejemplo ST-AB12CD."/>
                <Grid Margin="0,24,0,0">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Button x:Name="BtnDataBack" Style="{StaticResource BtnSoft}" Content="Atrás" Margin="0,0,12,0"/>
                  <Button x:Name="BtnDataNext" Grid.Column="1" Style="{StaticResource BtnPrimary}" IsEnabled="False">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock x:Name="TxtDataNext" Text="Continuar"/>
                      <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="12,1,0,0" FontSize="14"/>
                    </StackPanel>
                  </Button>
                  <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0,0,0">
                    <TextBlock Text="&#xE72E;" FontFamily="{StaticResource Icons}" FontSize="15" Foreground="#AEBDD0" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Style="{StaticResource Small}" Text="Tus datos solo se usan para activar Office." VerticalAlignment="Center"/>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Grid>
          </Border>
        </ScrollViewer>
      </Grid>

      <!-- CONFIRMACION -->
      <Grid x:Name="PageConfirm" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Border MaxWidth="980" Margin="40,8,40,18" CornerRadius="22" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="2" VerticalAlignment="Top">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="108"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border CornerRadius="20,0,0,20" Background="{StaticResource StubGrad}">
              <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                <TextBlock x:Name="TxtConfirmStubLabel" Text="OFFICE" Foreground="#FECACA" FontSize="13" FontWeight="Bold" HorizontalAlignment="Center"/>
                <TextBlock x:Name="TxtConfirmStubYear" Text="2024" Foreground="White" FontSize="32" FontWeight="Black" HorizontalAlignment="Center"/>
              </StackPanel>
            </Border>
            <Line Grid.Column="1" X1="0" Y1="0" X2="0" Y2="1" Stretch="Fill" Margin="0,20" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
            <Ellipse Grid.ColumnSpan="2" Width="26" Height="26" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="95,-13,0,0"/>
            <Ellipse Grid.ColumnSpan="2" Width="26" Height="26" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="95,0,0,-13"/>
          <StackPanel Grid.Column="1" Margin="32,20,32,14">
            <TextBlock x:Name="TxtConfirmTitle" Style="{StaticResource H1}" FontSize="28" Text="Antes de empezar"/>
            <TextBlock x:Name="TxtConfirmLead" Style="{StaticResource Lead}" Text="" Visibility="Collapsed"/>
            <Border x:Name="PanelDetected" Style="{StaticResource Panel}" Background="#111C2F" Margin="0,12,0,0" Padding="18,12">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Border x:Name="DetectedIconBg" Width="48" Height="48" CornerRadius="14" Background="#26F59E0B" VerticalAlignment="Top">
                  <TextBlock x:Name="DetectedIcon" Text="&#xE7BA;" FontFamily="{StaticResource Icons}" FontSize="20" Foreground="#FCD34D" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <StackPanel Grid.Column="1" Margin="18,0,0,0" VerticalAlignment="Center">
                  <TextBlock x:Name="TxtDetectedTitle" Text="" FontWeight="SemiBold" FontSize="19" TextWrapping="Wrap"/>
                  <TextBlock x:Name="TxtDetected" Style="{StaticResource Body}" Foreground="#C9D5E3" Margin="0,4,0,0" Text=""/>
                </StackPanel>
              </Grid>
            </Border>
            <StackPanel x:Name="StackConfirmList" Margin="0,14,0,0"/>
            <CheckBox x:Name="ChkConfirm" Style="{StaticResource Check}" Margin="0,4,0,0" Content="Guardé mis documentos y leí estas indicaciones."/>
            <WrapPanel Margin="0,16,0,0">
              <Button x:Name="BtnConfirmBack" Style="{StaticResource BtnSoft}" Content="Atrás" Margin="0,0,12,10"/>
              <Button x:Name="BtnConfirmAlt" Style="{StaticResource BtnSoft}" Content="Solo activar" Margin="0,0,12,10" Visibility="Collapsed"/>
              <Button x:Name="BtnConfirmStart" Style="{StaticResource BtnPrimary}" IsEnabled="False" Margin="0,0,0,10">
                <StackPanel Orientation="Horizontal">
                  <TextBlock x:Name="TxtConfirmStart" Text="Comenzar instalación"/>
                  <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="12,1,0,0" FontSize="14"/>
                </StackPanel>
              </Button>
            </WrapPanel>
          </StackPanel>
          </Grid>
          </Border>
        </ScrollViewer>
      </Grid>

      <!-- PROGRESO -->
      <Grid x:Name="PageProgress" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid MaxWidth="1240" Margin="40,8,40,18">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="340"/></Grid.ColumnDefinitions>
            <StackPanel Margin="0,0,36,0">
              <TextBlock x:Name="TxtProgTitle" Style="{StaticResource H1}" FontSize="30" Text="Instalando Office"/>
              <TextBlock x:Name="TxtProgDetail" Style="{StaticResource Lead}" FontSize="19" Text="Preparando..." MinHeight="54"/>
              <Grid Margin="0,14,0,0">
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <Border x:Name="ProgTrack" Height="14" CornerRadius="7" Background="#20304A" VerticalAlignment="Center">
                  <Border x:Name="ProgFill" HorizontalAlignment="Left" Width="0" CornerRadius="7" Background="{StaticResource RedGrad}"/>
                </Border>
                <TextBlock x:Name="TxtProgCount" Grid.Column="1" Style="{StaticResource Small}" FontSize="16" Margin="16,0,0,0" VerticalAlignment="Center" Text=""/>
              </Grid>
              <StackPanel x:Name="StackSteps" Margin="0,18,0,0"/>
              <Border x:Name="PanelTutorialPromo" Style="{StaticResource Panel}" Margin="0,10,0,0" Padding="20,14" Visibility="Collapsed">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <StackPanel VerticalAlignment="Center" Margin="0,0,16,0">
                    <TextBlock Text="Aprovecha la espera" FontWeight="SemiBold" FontSize="18"/>
                    <TextBlock Style="{StaticResource Small}" Text="Puedes ir preparando tu cuenta. La instalación sigue sola." Margin="0,2,0,0"/>
                  </StackPanel>
                  <Button x:Name="BtnTutorialEarly" Grid.Column="1" Style="{StaticResource BtnSoft}" Content="Preparar mi cuenta"/>
                </Grid>
              </Border>
            </StackPanel>
            <StackPanel Grid.Column="1">
              <Border CornerRadius="20" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="2">
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="88"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border CornerRadius="18,0,0,18" Background="{StaticResource StubGrad}">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                      <TextBlock x:Name="TxtProgStubLabel" Text="OFFICE" Foreground="#FECACA" FontSize="13" FontWeight="Bold" HorizontalAlignment="Center"/>
                      <TextBlock x:Name="TxtProgStubYear" Text="2024" Foreground="White" FontSize="26" FontWeight="Black" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Border>
                  <Line Grid.Column="1" X1="0" Y1="0" X2="0" Y2="1" Stretch="Fill" Margin="0,16" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
                  <Ellipse Grid.ColumnSpan="2" Width="22" Height="22" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="77,-11,0,0"/>
                  <Ellipse Grid.ColumnSpan="2" Width="22" Height="22" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="77,0,0,-11"/>
                  <StackPanel Grid.Column="1" Margin="22,18,18,18">
                    <TextBlock Text="Llevamos" Style="{StaticResource Small}"/>
                    <TextBlock x:Name="TxtElapsed" Text="00:00" FontSize="36" FontWeight="SemiBold" FontFamily="Segoe UI Variable Display, Segoe UI"/>
                    <TextBlock x:Name="TxtElapsedNote" Style="{StaticResource Small}" Text="Suele tardar entre 10 y 40 minutos, según tu Internet."/>
                    <Border x:Name="ProgStamp" HorizontalAlignment="Left" Margin="2,14,0,0" CornerRadius="10" BorderBrush="#F59E0B" BorderThickness="3" Padding="12,4" RenderTransformOrigin="0.5,0.5">
                      <Border.RenderTransform><TransformGroup><ScaleTransform x:Name="ProgStampScale" ScaleX="1" ScaleY="1"/><RotateTransform Angle="-6"/></TransformGroup></Border.RenderTransform>
                      <TextBlock x:Name="TxtProgStamp" Text="PENDIENTE" FontSize="17" FontWeight="Black" Foreground="#FCD34D"/>
                    </Border>
                  </StackPanel>
                </Grid>
              </Border>
              <Border Style="{StaticResource Panel}" Margin="0,16,0,0" BorderBrush="#80F59E0B" Background="#1AF59E0B" Padding="20,16">
                <StackPanel>
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE7BA;" FontFamily="{StaticResource Icons}" FontSize="18" Foreground="#FCD34D" VerticalAlignment="Center"/>
                    <TextBlock Text="Mientras tanto" FontWeight="SemiBold" FontSize="18" Margin="10,0,0,0"/>
                  </StackPanel>
                  <TextBlock x:Name="TxtWhile" Style="{StaticResource Body}" FontSize="16" Margin="0,8,0,0" Text="No apagues ni reinicies el computador."/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </Grid>

      <!-- LISTO -->
      <Grid x:Name="PageDone" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Border MaxWidth="900" Margin="40,8,40,18" CornerRadius="22" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="2" VerticalAlignment="Top">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="118"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border CornerRadius="20,0,0,20" Background="{StaticResource StubGrad}">
                <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                  <TextBlock x:Name="TxtDoneStubLabel" Text="OFFICE" Foreground="#FECACA" FontSize="13" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock x:Name="TxtDoneStubYear" Text="2024" Foreground="White" FontSize="34" FontWeight="Black" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Line Grid.Column="1" X1="0" Y1="0" X2="0" Y2="1" Stretch="Fill" Margin="0,20" Stroke="#405776" StrokeThickness="2" StrokeDashArray="2.5 2.5" HorizontalAlignment="Left"/>
              <Ellipse Grid.ColumnSpan="2" Width="26" Height="26" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="105,-13,0,0"/>
              <Ellipse Grid.ColumnSpan="2" Width="26" Height="26" Fill="#09111F" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="105,0,0,-13"/>
              <StackPanel Grid.Column="1" Margin="36,28,36,28">
                <TextBlock x:Name="TxtDoneTitle" Style="{StaticResource H1}" FontSize="32" Text="¡Listo!" Margin="0,0,170,0"/>
                <TextBlock x:Name="TxtDoneText" Style="{StaticResource Lead}" Text="" Margin="0,8,170,0"/>
                <StackPanel x:Name="StackDoneList" Margin="0,18,0,0"/>
                <WrapPanel Margin="0,20,0,0">
                  <Button x:Name="BtnDoneWord" Style="{StaticResource BtnPrimary}" Margin="0,0,12,10">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE8A5;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                      <TextBlock Text="Abrir Word"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnDoneClose" Style="{StaticResource BtnSoft}" Content="Cerrar el instalador" Margin="0,0,0,10"/>
                </WrapPanel>
                <Border x:Name="DoneNote" Margin="0,8,0,0">
                  <TextBlock x:Name="TxtDoneNote" Style="{StaticResource Small}" FontSize="16" Text=""/>
                </Border>
              </StackPanel>
              <Border x:Name="DoneStamp" Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,30,30,0" CornerRadius="12" BorderBrush="#34D399" BorderThickness="3" Padding="16,6" RenderTransformOrigin="0.5,0.5">
                <Border.RenderTransform><TransformGroup><ScaleTransform x:Name="DoneStampScale" ScaleX="1" ScaleY="1"/><RotateTransform Angle="-9"/></TransformGroup></Border.RenderTransform>
                <TextBlock x:Name="TxtDoneStamp" Text="ACTIVADO" FontSize="24" FontWeight="Black" Foreground="#34D399"/>
              </Border>
            </Grid>
          </Border>
        </ScrollViewer>
      </Grid>

      <!-- ERROR -->
      <Grid x:Name="PageError" Visibility="Collapsed">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <StackPanel MaxWidth="840" Margin="40,4,40,18">
            <StackPanel Orientation="Horizontal">
              <Border Width="56" Height="56" CornerRadius="16" Background="#26F59E0B" VerticalAlignment="Top">
                <TextBlock Text="&#xE7BA;" FontFamily="{StaticResource Icons}" FontSize="24" Foreground="#FCD34D" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <TextBlock x:Name="TxtErrTitle" Style="{StaticResource H1}" FontSize="30" Margin="18,0,0,0" VerticalAlignment="Center" MaxWidth="740" Text="No pudimos terminar"/>
            </StackPanel>
            <TextBlock Style="{StaticResource Lead}" Margin="0,12,0,0" Text="No te preocupes: no se borró nada. Puedes reintentar ahora o pedirnos ayuda."/>
            <Border CornerRadius="16" Background="#17243A" BorderBrush="#80F59E0B" BorderThickness="1.5" Padding="20,16" Margin="0,16,0,0">
              <TextBox x:Name="TxtErrDetail" IsReadOnly="True" TextWrapping="Wrap" Background="#00000000" BorderThickness="0"
                       Foreground="#DBE5F1" FontSize="16" MaxHeight="190" VerticalScrollBarVisibility="Auto" Text=""/>
            </Border>
            <WrapPanel Margin="0,20,0,0">
              <Button x:Name="BtnErrRetry" Style="{StaticResource BtnPrimary}" Content="Reintentar" Margin="0,0,12,10"/>
              <Button x:Name="BtnErrWa" Style="{StaticResource BtnWa}" Margin="0,0,12,10">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE8BD;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                  <TextBlock Text="Pedir ayuda por WhatsApp"/>
                </StackPanel>
              </Button>
            </WrapPanel>
            <WrapPanel Margin="-10,0,0,0">
              <Button x:Name="BtnErrHome" Style="{StaticResource BtnLink}" Content="Volver al inicio" Margin="0,0,8,0"/>
              <Button x:Name="BtnErrZip" Style="{StaticResource BtnLink}" Content="Guardar informe para soporte en el Escritorio"/>
            </WrapPanel>
          </StackPanel>
        </ScrollViewer>
      </Grid>

      <!-- TUTORIAL MICROSOFT 365 -->
      <Grid x:Name="PageTutorial" Visibility="Collapsed">
        <Grid MaxWidth="1380" Margin="32,18,32,16">
          <Grid.ColumnDefinitions><ColumnDefinition Width="250"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border Style="{StaticResource Panel}" Padding="12,16" Margin="0,0,26,0">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <StackPanel Margin="8,0,8,8">
                <TextBlock Text="Cuenta Microsoft 365" FontWeight="SemiBold" FontSize="18" TextWrapping="Wrap"/>
                <Border x:Name="ChipInstall" CornerRadius="14" Padding="12,5" Margin="0,10,0,0" HorizontalAlignment="Left" Background="#1AF59E0B" BorderBrush="#80F59E0B" BorderThickness="1.5" Visibility="Collapsed">
                  <TextBlock x:Name="TxtChipInstall" FontSize="14" FontWeight="SemiBold" Foreground="#FCD34D" Text="Office instalándose"/>
                </Border>
                <Button x:Name="BtnTutProgress" Style="{StaticResource BtnLink}" FontSize="15" Margin="-10,4,0,0" HorizontalAlignment="Left" Visibility="Collapsed" Content="Ver el avance"/>
                <Button x:Name="BtnTutHome" Style="{StaticResource BtnLink}" FontSize="15" Margin="-10,2,0,0" HorizontalAlignment="Left">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE72B;" FontFamily="{StaticResource Icons}" FontSize="13" VerticalAlignment="Center" Margin="0,1,8,0"/>
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
            <Grid x:Name="TutColumns">
              <Grid.ColumnDefinitions><ColumnDefinition Width="1.2*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <ScrollViewer x:Name="TutScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel x:Name="TutContent" Margin="0,2,18,10">
                  <TextBlock x:Name="TxtTutTitle" Style="{StaticResource H1}" FontSize="28" Text=""/>
                  <StackPanel x:Name="StackTutBody" Margin="0,16,0,0"/>
                  <WrapPanel x:Name="StackTutActions" Margin="0,6,0,4"/>
                  <Border x:Name="TutNote" CornerRadius="14" BorderThickness="1.5" Padding="18,14" Margin="0,8,0,0">
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <TextBlock x:Name="TutNoteIcon" FontFamily="{StaticResource Icons}" FontSize="20" Margin="0,2,14,0" Text="&#xE946;"/>
                      <StackPanel Grid.Column="1">
                        <TextBlock x:Name="TxtTutNoteLabel" FontSize="14" FontWeight="Bold" Margin="0,1,0,4" Text="CONSEJO"/>
                        <TextBlock x:Name="TxtTutNote" Style="{StaticResource Body}" FontSize="16" Text=""/>
                      </StackPanel>
                    </Grid>
                  </Border>
                </StackPanel>
              </ScrollViewer>
              <Border x:Name="TutImageFrame" Grid.Column="1" CornerRadius="16" BorderBrush="#2A3C57" BorderThickness="1.5" Background="#0C1626" Margin="10,0,0,10" Padding="10" VerticalAlignment="Top" HorizontalAlignment="Left" MaxHeight="{Binding ActualHeight, ElementName=TutColumns}">
                <Image x:Name="ImgTut" Stretch="Uniform" StretchDirection="DownOnly" RenderOptions.BitmapScalingMode="HighQuality"/>
              </Border>
            </Grid>
            <Border Grid.Row="1" BorderBrush="#2A3C57" BorderThickness="0,1,0,0" Padding="0,14,0,0">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center" Margin="0,0,24,0">
                  <TextBlock x:Name="TxtTutProgress" Style="{StaticResource Small}" FontSize="16" Margin="0,0,0,8" Text=""/>
                  <Border Height="10" CornerRadius="5" Background="#20304A">
                    <Border x:Name="TutProgFill" HorizontalAlignment="Left" Width="0" CornerRadius="5" Background="{StaticResource RedGrad}"/>
                  </Border>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                  <Button x:Name="BtnTutBack" Style="{StaticResource BtnSoft}" Margin="0,0,12,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE72B;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,1,10,0" FontSize="14"/>
                      <TextBlock Text="Atrás"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnTutNext" Style="{StaticResource BtnPrimary}" MinWidth="170">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock x:Name="TxtTutNext" Text="Siguiente"/>
                      <TextBlock Text="&#xE72A;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="12,1,0,0" FontSize="14"/>
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
    <Border Grid.Row="3" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="0,1,0,0" Padding="0,4">
      <Grid Margin="24,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="¿Dudas?" Foreground="#DBE5F1" FontSize="16" VerticalAlignment="Center" Margin="0,0,12,0"/>
          <Button x:Name="BtnHelpWa" Style="{StaticResource BtnWa}" MinHeight="48" Padding="18,6" FontSize="16">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE8BD;" FontFamily="{StaticResource Icons}" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock Text="Escríbenos por WhatsApp"/>
            </StackPanel>
          </Button>
          <Button x:Name="BtnHelpCenter" Style="{StaticResource BtnLink}" FontSize="16" Margin="10,0,0,0" Content="Centro de Ayuda"/>
        </StackPanel>
        <TextBlock x:Name="TxtVersion" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="#AEBDD0" FontSize="14" Text="SafeTech SpA"/>
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
        FontSize="17" Foreground="#F8FAFC" UseLayoutRounding="True" TextOptions.TextFormattingMode="Ideal">
  <Window.Resources>
    __RESOURCES__
  </Window.Resources>
  <Border x:Name="DlgCard" Margin="24" CornerRadius="24" Background="#111C2F" BorderBrush="#405776" BorderThickness="1.5" MinWidth="480" MaxWidth="640">
    <Border.Effect><DropShadowEffect Color="#000000" BlurRadius="40" ShadowDepth="10" Direction="270" Opacity="0.55"/></Border.Effect>
    <Grid>
      <Border Height="2" VerticalAlignment="Top" Margin="40,0" Background="{StaticResource TopLine}"/>
      <StackPanel Margin="30,28,30,26">
        <Grid x:Name="DlgDrag" Background="#00000000">
          <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <Border x:Name="DlgIconBg" Width="52" Height="52" CornerRadius="16" Background="#262563EB" VerticalAlignment="Top">
            <TextBlock x:Name="DlgIcon" FontFamily="{StaticResource Icons}" FontSize="22" Foreground="#93C5FD" HorizontalAlignment="Center" VerticalAlignment="Center" Text="&#xE946;"/>
          </Border>
          <StackPanel Grid.Column="1" Margin="18,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="DlgKicker" Visibility="Collapsed" Text=""/>
            <TextBlock x:Name="DlgTitle" Style="{StaticResource H2}" FontSize="23" Text=""/>
          </StackPanel>
        </Grid>
        <TextBlock x:Name="DlgMessage" Style="{StaticResource Body}" Margin="0,18,0,0" Text=""/>
        <ContentControl x:Name="DlgExtra" Margin="0,16,0,0"/>
        <WrapPanel x:Name="DlgButtons" HorizontalAlignment="Right" Margin="0,22,0,0"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
'@

# ----- SafeBOT vectorial y animado (mismo dibujo plano de los videos de SafeTech) ------------------
# Lienzo de 400x520 con partes separadas: cuerpo, brazo que saluda, cabeza y dos juegos de ojos.
$StBotXaml = @'
<Viewbox xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Stretch="Uniform">
  <Canvas x:Name="BotRoot" Width="400" Height="520" RenderTransformOrigin="0.5,1">
    <Canvas.RenderTransform>
      <TransformGroup>
        <ScaleTransform x:Name="BotSquash" ScaleX="1" ScaleY="1"/>
        <TranslateTransform x:Name="BotBob" Y="0"/>
        <TranslateTransform x:Name="BotJump" Y="0"/>
      </TransformGroup>
    </Canvas.RenderTransform>
    <Rectangle Canvas.Left="58" Canvas.Top="352" Width="44" Height="108" RadiusX="22" RadiusY="22" Fill="#E9E9E6"/>
    <Rectangle Canvas.Left="58" Canvas.Top="352" Width="44" Height="30" RadiusX="15" RadiusY="15" Fill="#DC2626"/>
    <Ellipse Canvas.Left="56" Canvas.Top="434" Width="48" Height="48" Fill="#F4F4F2"/>
    <Path Data="M110,340 H290 Q312,340 310,362 L300,478 Q298,500 276,500 H124 Q102,500 100,478 L90,362 Q88,340 110,340 Z" Fill="#F4F4F2"/>
    <Path Data="M100,470 Q102,500 124,500 H276 Q298,500 300,470 Z" Fill="#D9D9D5"/>
    <Rectangle Canvas.Left="168" Canvas.Top="476" Width="64" Height="10" RadiusX="5" RadiusY="5" Fill="#DC2626"/>
    <TextBlock Canvas.Left="120" Canvas.Top="366" Width="160" TextAlignment="Center" FontFamily="Segoe UI Black, Segoe UI" FontWeight="Black" FontSize="64"><Run Text="S" Foreground="#172033"/><Run Text="T" Foreground="#DC2626"/></TextBlock>
    <Rectangle Canvas.Left="172" Canvas.Top="312" Width="56" Height="34" RadiusX="10" RadiusY="10" Fill="#3A3F47"/>
    <Canvas x:Name="BotArm">
      <Canvas.RenderTransform><RotateTransform x:Name="BotArmRot" CenterX="320" CenterY="364" Angle="0"/></Canvas.RenderTransform>
      <Rectangle Canvas.Left="298" Canvas.Top="352" Width="44" Height="108" RadiusX="22" RadiusY="22" Fill="#E9E9E6"/>
      <Rectangle Canvas.Left="298" Canvas.Top="352" Width="44" Height="30" RadiusX="15" RadiusY="15" Fill="#DC2626"/>
      <Ellipse Canvas.Left="296" Canvas.Top="434" Width="48" Height="48" Fill="#F4F4F2"/>
    </Canvas>
    <Canvas x:Name="BotHead">
      <Canvas.RenderTransform><RotateTransform x:Name="BotHeadRot" CenterX="200" CenterY="316" Angle="0"/></Canvas.RenderTransform>
      <Path Data="M70,120 L44,22 Q42,10 54,14 L104,92 Z" Fill="#F4F4F2"/>
      <Path Data="M48,36 L44,22 Q42,10 54,14 L72,42 Z" Fill="#DC2626"/>
      <Path Data="M330,120 L356,22 Q358,10 346,14 L296,92 Z" Fill="#F4F4F2"/>
      <Path Data="M352,36 L356,22 Q358,10 346,14 L328,42 Z" Fill="#DC2626"/>
      <Ellipse Canvas.Left="18" Canvas.Top="156" Width="88" Height="88" Fill="#3A3F47"/>
      <Ellipse Canvas.Left="32" Canvas.Top="170" Width="60" Height="60" Fill="#F4F4F2"/>
      <Ellipse Canvas.Left="42" Canvas.Top="180" Width="40" Height="40" Stroke="#DC2626" StrokeThickness="8"/>
      <Ellipse Canvas.Left="294" Canvas.Top="156" Width="88" Height="88" Fill="#3A3F47"/>
      <Ellipse Canvas.Left="308" Canvas.Top="170" Width="60" Height="60" Fill="#F4F4F2"/>
      <Ellipse Canvas.Left="318" Canvas.Top="180" Width="40" Height="40" Stroke="#DC2626" StrokeThickness="8"/>
      <Rectangle Canvas.Left="72" Canvas.Top="56" Width="256" Height="262" RadiusX="118" RadiusY="118" Fill="#F4F4F2"/>
      <Path Data="M80,250 Q110,316 200,318 Q290,316 320,250 Q300,300 200,304 Q100,300 80,250 Z" Fill="#E2E2DE"/>
      <Path Data="M160,58 H240 L232,102 H168 Z" Fill="#DC2626"/>
      <Rectangle Canvas.Left="104" Canvas.Top="126" Width="192" Height="150" RadiusX="66" RadiusY="66" Fill="#0D0F14"/>
      <Rectangle Canvas.Left="118" Canvas.Top="138" Width="60" Height="18" RadiusX="9" RadiusY="9" Fill="#FFFFFF" Opacity="0.10"/>
      <Canvas x:Name="BotEyes">
        <Canvas.RenderTransform><ScaleTransform x:Name="BotEyesScale" CenterX="200" CenterY="201" ScaleX="1" ScaleY="1"/></Canvas.RenderTransform>
        <Rectangle Canvas.Left="148" Canvas.Top="176" Width="30" Height="50" RadiusX="15" RadiusY="15" Fill="#FF5A5A"/>
        <Rectangle Canvas.Left="222" Canvas.Top="176" Width="30" Height="50" RadiusX="15" RadiusY="15" Fill="#FF5A5A"/>
      </Canvas>
      <Path x:Name="BotEyesHappy" Data="M145,212 Q163,182 181,212 M219,212 Q237,182 255,212" Stroke="#FF5A5A" StrokeThickness="12" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Opacity="0"/>
    </Canvas>
  </Canvas>
</Viewbox>
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

function Open-StWhatsApp {
    param([string]$Message)
    $url = $StConfig.WhatsAppUrl
    if ($Message) { $url += '?text=' + [Uri]::EscapeDataString($Message) }
    Open-StUnelevated $url
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
    Invoke-StBotAction -Action $(if ($Name -eq 'Done') { 'happy' } else { 'normal' })
    Update-StGuide -Page $Name
    if ($Name -eq 'Data') { Invoke-StCue 'datos' -Once }
    if ($Name -eq 'Confirm') { Invoke-StCue 'confirmar' -Once }
    $St.UI.GuideBar.Visibility = $(if ($Name -eq 'Tutorial') { 'Collapsed' } else { 'Visible' })
    if ($St.NoAnim) { $el.Opacity = 1; return }
    $tt = New-Object System.Windows.Media.TranslateTransform 0, 18
    $el.RenderTransform = $tt
    $ease = New-Object System.Windows.Media.Animation.QuinticEase
    $ease.EasingMode = 'EaseOut'
    $dur = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(360))
    $fade = [System.Windows.Media.Animation.DoubleAnimation]::new([double]0, [double]1, $dur)
    $fade.EasingFunction = $ease
    $move = [System.Windows.Media.Animation.DoubleAnimation]::new([double]18, [double]0, $dur)
    $move.EasingFunction = $ease
    $el.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fade)
    $tt.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $move)
}

# ----- Guía de SafeBOT y etapas ----------------------------------------------------------------
function Set-StGuide {
    param([string]$Text)
    if ($St.UI.TxtGuide.Text -ne $Text) { Invoke-StBotAction -Action 'nod' }
    $St.UI.TxtGuide.Text = $Text
}

# ----- SafeBOT animado ------------------------------------------------------------------------------
function New-StBot {
    try {
        $view = [Windows.Markup.XamlReader]::Parse($StBotXaml)
        $bot = @{ View = $view }
        foreach ($n in @('BotRoot', 'BotSquash', 'BotBob', 'BotJump', 'BotArmRot', 'BotHeadRot', 'BotEyes', 'BotEyesScale', 'BotEyesHappy')) {
            $bot[$n] = $view.FindName($n)
        }
        return $bot
    } catch {
        try { Write-StLog "No se pudo dibujar a SafeBOT: $($_.Exception.Message)" } catch { }
        return $null
    }
}

function New-StKeyAnim {
    # Cuadros "milisegundos=valor" con transición suave; la duración total es el último tiempo.
    param([string[]]$Frames, [switch]$Forever, [int]$BeginMs = 0)
    $a = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames
    foreach ($f in $Frames) {
        $parts = $f.Split('=')
        $k = New-Object System.Windows.Media.Animation.EasingDoubleKeyFrame
        $k.KeyTime = [System.Windows.Media.Animation.KeyTime]::FromTimeSpan([TimeSpan]::FromMilliseconds([double]$parts[0]))
        $k.Value = [double]$parts[1]
        $e = New-Object System.Windows.Media.Animation.SineEase
        $e.EasingMode = 'EaseInOut'
        $k.EasingFunction = $e
        [void]$a.KeyFrames.Add($k)
    }
    if ($Forever) { $a.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever }
    if ($BeginMs -gt 0) { $a.BeginTime = [TimeSpan]::FromMilliseconds($BeginMs) }
    return $a
}

function Start-StBotIdle {
    # Respira (sube y baja muy poco) y pestañea de vez en cuando.
    param($Bot)
    if (-not $Bot -or $St.NoAnim) { return }
    $Bot.BotBob.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, (New-StKeyAnim @('0=0', '1700=-16', '3400=0') -Forever))
    $Bot.BotEyesScale.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty,
        (New-StKeyAnim @('0=1', '3600=1', '3680=0.1', '3760=1', '6900=1', '6980=0.1', '7060=1', '7140=1', '7220=0.1', '7300=1', '9000=1') -Forever))
}

function Invoke-StBotAction {
    # wave = saluda, jump = salta de alegría, nod = asiente al hablar, shake = niega suave,
    # happy / normal = cambia los ojos.
    param($Bot, [ValidateSet('wave', 'jump', 'celebrate', 'nod', 'shake', 'happy', 'normal')][string]$Action, [int]$BeginMs = 0)
    if (-not $Bot) { $Bot = $St.Bot }
    if (-not $Bot) { return }
    try {
        if ($Action -eq 'happy' -or $Action -eq 'celebrate') { $Bot.BotEyes.Opacity = 0; $Bot.BotEyesHappy.Opacity = 1 }
        if ($Action -eq 'normal') { $Bot.BotEyes.Opacity = 1; $Bot.BotEyesHappy.Opacity = 0 }
        if ($St.NoAnim -or $Action -eq 'happy' -or $Action -eq 'normal') { return }
        $angle = [System.Windows.Media.RotateTransform]::AngleProperty
        $wave = @('0=0', '350=-140', '600=-112', '850=-140', '1100=-112', '1350=-140', '1800=0')
        switch ($Action) {
            'wave' { $Bot.BotArmRot.BeginAnimation($angle, (New-StKeyAnim $wave -BeginMs $BeginMs)) }
            'nod' { $Bot.BotHeadRot.BeginAnimation($angle, (New-StKeyAnim @('0=0', '180=-5', '400=4', '600=0') -BeginMs $BeginMs)) }
            'shake' { $Bot.BotHeadRot.BeginAnimation($angle, (New-StKeyAnim @('0=0', '120=7', '260=-7', '400=5', '540=-4', '680=0') -BeginMs $BeginMs)) }
            { $_ -eq 'jump' -or $_ -eq 'celebrate' } {
                $Bot.BotJump.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty,
                    (New-StKeyAnim @('0=0', '140=10', '420=-80', '700=0', '820=6', '940=0') -BeginMs $BeginMs))
                $Bot.BotSquash.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty,
                    (New-StKeyAnim @('0=1', '140=0.9', '320=1.05', '700=1', '820=0.94', '940=1') -BeginMs $BeginMs))
                if ($Action -eq 'celebrate') { $Bot.BotArmRot.BeginAnimation($angle, (New-StKeyAnim $wave -BeginMs ($BeginMs + 700))) }
            }
        }
    } catch { }
}

function New-StBotForDialog {
    # SafeBOT propio para una ventana emergente: respira, pestañea y saluda al aparecer.
    $b = New-StBot
    if (-not $b) { return $null }
    Start-StBotIdle $b
    Invoke-StBotAction -Bot $b -Action 'wave' -BeginMs 450
    return $b.View
}

# ----- Sonidos y voz de SafeBOT -------------------------------------------------------------------
# Sonidos cortos en assets/sonidos y frases de SafeBOT en assets/voz (se descargan al abrir).
# Si una frase todavía no existe, solo suena el aviso. El botón de la barra superior los apaga.
$StCues = @{
    'bienvenida'    = @{ Voice = 'bienvenida' }
    'datos'         = @{ Voice = 'datos' }
    'confirmar'     = @{ Voice = 'confirmar' }
    'instalando'    = @{ Voice = 'instalando' }
    'instalado'     = @{ Sound = 'exito'; Voice = 'instalado'; Bot = 'jump' }
    'instalado-365' = @{ Sound = 'exito'; Voice = 'instalado-365'; Bot = 'jump' }
    'activado'      = @{ Sound = 'logro'; Voice = 'activado'; Bot = 'celebrate' }
    'conectado'     = @{ Sound = 'logro'; Voice = 'conectado'; Bot = 'celebrate' }
    'error'         = @{ Sound = 'error'; Voice = 'error'; Bot = 'shake' }
    'codigo'        = @{ Sound = 'aviso'; Voice = 'codigo'; Bot = 'shake' }
    'atencion'      = @{ Sound = 'aviso'; Bot = 'nod' }
}
$StVoiceDir = 'assets/voz/v1'

function Get-StCueAssets {
    $list = @('assets/sonidos/aviso.wav', 'assets/sonidos/exito.wav', 'assets/sonidos/logro.wav', 'assets/sonidos/error.wav')
    foreach ($c in $StCues.Values) { if ($c.Voice) { $list += "$StVoiceDir/$($c.Voice).mp3" } }
    return @($list | Select-Object -Unique)
}

function Get-StAssetFile {
    # Ruta local de un recurso ya descargado (no descarga: no debe frenar la ventana).
    param([string]$RelPath)
    if ($env:SAFETECH_OFFICE_ASSETS) { $f = Join-Path $env:SAFETECH_OFFICE_ASSETS $RelPath }
    else { $f = Join-Path $St.CacheDir ($RelPath -replace '[\\/]', '_') }
    if (Test-Path -LiteralPath $f) { return $f }
    return $null
}

function Set-StSound {
    param([bool]$On, [switch]$Save)
    $St.SoundOn = $On
    $St.UI.TxtSound.Text = $(if ($On) { 'Sonido activado' } else { 'Sonido apagado' })
    $St.UI.TxtSoundIcon.Text = [string][char]$(if ($On) { 0xE767 } else { 0xE74F })
    if (-not $On -and $St.VoicePlayer) { try { $St.VoicePlayer.Stop() } catch { } }
    if (-not $Save) { return }
    $mark = Join-Path $StConfig.DataDir 'sonido-apagado'
    try {
        if ($On) { Remove-Item -LiteralPath $mark -ErrorAction SilentlyContinue }
        else { New-Item -ItemType Directory -Force -Path $StConfig.DataDir | Out-Null; Set-Content -LiteralPath $mark -Value '1' }
    } catch { }
}

function Start-StVoice {
    param([string]$Voice)
    if (-not $Voice -or -not $St.SoundOn) { return }
    $f = Get-StAssetFile "$StVoiceDir/$Voice.mp3"
    if (-not $f) { return }
    try {
        if (-not $St.VoicePlayer) { $St.VoicePlayer = New-Object System.Windows.Media.MediaPlayer }
        $St.VoicePlayer.Stop()
        $St.VoicePlayer.Open((New-Object Uri $f))
        $St.VoicePlayer.Volume = 1.0
        $St.VoicePlayer.Play()
        Invoke-StBotAction -Action 'nod'
    } catch { }
}

function Invoke-StCue {
    # Un aviso: SafeBOT reacciona, suena un tono y, si existe, habla.
    param([string]$Name, [switch]$Once)
    $cue = $StCues[$Name]
    if (-not $cue) { return }
    if ($Once) {
        if ($St.CuesDone[$Name]) { return }
        $St.CuesDone[$Name] = $true
    }
    if ($cue.Bot) { Invoke-StBotAction -Action $cue.Bot }
    if (-not $St.SoundOn -or $St.SelfTest) { return }
    $delay = 0
    if ($cue.Sound) {
        $w = Get-StAssetFile "assets/sonidos/$($cue.Sound).wav"
        if ($w) {
            try { $St.SoundPlayer = New-Object System.Media.SoundPlayer $w; $St.SoundPlayer.Play(); $delay = 900 } catch { }
        }
    }
    if (-not $cue.Voice) { return }
    if ($delay -eq 0) { Start-StVoice $cue.Voice; return }
    $St.PendingVoice = $cue.Voice
    $St.VoiceTimer.Stop()
    $St.VoiceTimer.Interval = [TimeSpan]::FromMilliseconds($delay)
    $St.VoiceTimer.Start()
}

function Show-StWindowFront {
    # Trae la ventana al frente (por ejemplo, cuando termina la ventana de Microsoft).
    try {
        if ($St.Window.WindowState -eq 'Minimized') { $St.Window.WindowState = 'Maximized' }
        $St.Window.Topmost = $true
        [void]$St.Window.Activate()
        $St.Window.Topmost = $false
    } catch { }
}

function Set-StJourney {
    # Stage: 0 = oculto, 1..3 = etapa actual, 4 = todo listo.
    param([int]$Stage, [string]$Label2 = 'Tus datos')
    $ui = $St.UI
    if ($Stage -le 0) { $ui.Journey.Visibility = 'Collapsed'; return }
    $ui.Journey.Visibility = 'Visible'
    $ui.JLbl2.Text = $Label2
    $conv = New-Object System.Windows.Media.BrushConverter
    for ($i = 1; $i -le 3; $i++) {
        $dot = $ui["JDot$i"]; $num = $ui["JNum$i"]; $lbl = $ui["JLbl$i"]
        if ($i -lt $Stage) {
            $dot.Background = $conv.ConvertFromString('#059669'); $dot.BorderBrush = $conv.ConvertFromString('#059669')
            $num.Text = [string][char]0xE73E; $num.FontFamily = 'Segoe Fluent Icons, Segoe MDL2 Assets'; $num.FontSize = 14
            $num.Foreground = $conv.ConvertFromString('#FFFFFF'); $lbl.Foreground = $conv.ConvertFromString('#C9D5E3'); $lbl.FontWeight = 'Normal'
        } elseif ($i -eq $Stage) {
            $dot.Background = $St.Window.FindResource('RedGrad'); $dot.BorderBrush = $conv.ConvertFromString('#DC2626')
            $num.Text = [string]$i; $num.FontFamily = 'Segoe UI'; $num.FontSize = 16
            $num.Foreground = $conv.ConvertFromString('#FFFFFF'); $lbl.Foreground = $conv.ConvertFromString('#F8FAFC'); $lbl.FontWeight = 'SemiBold'
        } else {
            $dot.Background = $conv.ConvertFromString('#00000000'); $dot.BorderBrush = $conv.ConvertFromString('#8FA3BE')
            $num.Text = [string]$i; $num.FontFamily = 'Segoe UI'; $num.FontSize = 16
            $num.Foreground = $conv.ConvertFromString('#AEBDD0'); $lbl.Foreground = $conv.ConvertFromString('#AEBDD0'); $lbl.FontWeight = 'Normal'
        }
    }
    $ui.JLine1.Background = $conv.ConvertFromString($(if ($Stage -gt 1) { '#059669' } else { '#405776' }))
    $ui.JLine2.Background = $conv.ConvertFromString($(if ($Stage -gt 2) { '#059669' } else { '#405776' }))
}

function Update-StGuide {
    param([string]$Page)
    $p = $St.Product
    $label2 = $(if ($p -and -not $p.NeedsKey) { 'Revisar' } else { 'Tus datos' })
    switch ($Page) {
        'Home' {
            Set-StGuide '¡Hola! Soy SafeBOT y te acompaño en todo el proceso. Lo que compraste aparece en el correo que te envió SafeTech.'
            Set-StJourney 1 'Tus datos'
        }
        'Data' {
            Set-StGuide 'Copia la clave y el código desde tu correo. Si te equivocas no pasa nada: lo puedes corregir.'
            Set-StJourney 2 $label2
        }
        'Confirm' {
            Set-StGuide 'Revisa esto con calma y marca la casilla. Después yo me encargo de todo.'
            Set-StJourney 2 $label2
        }
        'Progress' {
            if ($St.Mode -eq 'activate' -or ($St.InstallDone -and $p -and $p.NeedsKey)) {
                Set-StGuide 'Estoy activando Office con tu código. Tarda menos de un minuto.'
            } else {
                Set-StGuide 'Estoy instalando Office. Puedes dejar el computador solo: te aviso cuando termine.'
            }
            Set-StJourney 3 $label2
        }
        'Done' {
            Set-StGuide '¡Lo logramos! Gracias por confiar en SafeTech.'
            Set-StJourney 4 $label2
        }
        'Error' {
            Set-StGuide 'Esto tiene solución. Si prefieres, te ayudamos por WhatsApp.'
            Set-StJourney 0
        }
        'Tutorial' {
            Set-StGuide 'Sigue cada paso con calma. Si algo no coincide con tu pantalla, escríbenos por WhatsApp.'
            Set-StJourney 0
        }
    }
}

function Update-StScale {
    $w = $St.Window.ActualWidth; $h = $St.Window.ActualHeight
    if ($w -le 0 -or $h -le 0) { return }
    $scale = Get-StScale $w $h
    $St.UI.RootScale.ScaleX = $scale
    $St.UI.RootScale.ScaleY = $scale
}

function Update-StWindowFrame {
    # Sin el gancho de maximizado, la ventana maximizada sobresale unos pixeles de la pantalla: se compensa con un margen.
    # Con el gancho ocupa justo el area de trabajo y no hace falta.
    if ($St.Window.WindowState -eq 'Maximized') {
        $St.UI.RootBorder.Margin = $(if ($St.MaximizeFix) { '0' } else { '7' })
        $St.UI.RootBorder.BorderThickness = '0'
        $St.UI.BtnMax.Content = [string][char]0xE923
        $St.UI.BtnMax.ToolTip = 'Restaurar'
    } else {
        $St.UI.RootBorder.Margin = '0'
        $St.UI.RootBorder.BorderThickness = '1'
        $St.UI.BtnMax.Content = [string][char]0xE922
        $St.UI.BtnMax.ToolTip = 'Maximizar'
    }
}

# ----- Ventanas emergentes -------------------------------------------------------------------
function Show-StDialog {
    param(
        [string]$Title,
        [string]$Message,
        [ValidateSet('info', 'success', 'warn', 'error', 'question', 'key')][string]$Kind = 'info',
        [string[]]$Buttons = @('Aceptar'),
        [string]$Primary,
        [string]$Kicker = '',
        $Extra,
        [int]$Width = 0,
        [string]$LinkLabel,
        [scriptblock]$OnLink,
        $Image
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
    $card = $dlg.FindName('DlgCard')
    if ($Width -gt 0) { $card.MaxWidth = $Width; $card.MinWidth = $Width }
    $dscale = 1.0
    try { $dscale = [Math]::Max(1.0, [double]$St.UI.RootScale.ScaleX) } catch { }
    if ($dscale -gt 1.01) { $card.LayoutTransform = New-Object System.Windows.Media.ScaleTransform $dscale, $dscale }
    if ($Image) {
        $ib = $dlg.FindName('DlgIconBg')
        $ib.Background = $conv.ConvertFromString('#00000000'); $ib.Width = 64; $ib.Height = 80
        if ($Image -is [System.Windows.UIElement]) {
            $ib.Child = $Image
        } else {
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $Image; $img.Stretch = 'Uniform'
            $ib.Child = $img
        }
    }
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
        $b.Margin = '12,0,0,0'
        $b.MinWidth = 140
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
    if ($LinkLabel) {
        $lnk = New-Object System.Windows.Controls.Button
        $lnk.Content = $LinkLabel
        $lnk.Style = $dlg.FindResource('BtnLink')
        $lnk.Margin = '0,0,12,0'
        $lnk.Tag = $(if ($OnLink) { $OnLink } else { $LinkLabel })
        $lnk.Add_Click({
            param($s, $e)
            if ($s.Tag -is [scriptblock]) { & $s.Tag }
            else { $w = [System.Windows.Window]::GetWindow($s); $w.Tag = [string]$s.Tag; $w.Close() }
        })
        $panel.Children.Insert(0, $lnk)
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
        $h.FontSize = 16
        $h.TextWrapping = 'Wrap'
        $h.Margin = '2,8,0,0'
        [void]$sp.Children.Add($h)
    }
    return @{ Panel = $sp; Box = $box }
}

# ----- Tareas en segundo plano -------------------------------------------------------------------
function Start-StAsync {
    param([Parameter(Mandatory)][scriptblock]$Work, [hashtable]$Params = @{}, [scriptblock]$OnDone)
    $script = New-StAsyncScript -EngineText $St.EngineText -WorkText $Work.ToString()
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
        $diag = New-Object System.Collections.Generic.List[string]
        try {
            $out = @($job.PS.EndInvoke($job.Handle))
            $diag.Add("Objetos devueltos: $($out.Count)")
            $res = Get-StAsyncResult -Output $out
            if (-not ($res -is [hashtable])) {
                foreach ($item in $out) { if ($null -ne $item) { $diag.Add("Objeto: $($item.GetType().FullName)") } else { $diag.Add('Objeto: nulo') } }
            }
        } catch {
            $res = @{ Ok = $false; Result = $null; Error = $_.Exception.Message }
        }
        if (-not ($res -is [hashtable])) {
            try {
                $info = $job.PS.InvocationStateInfo
                $diag.Add("Estado: $($info.State)")
                if ($info.Reason) { $diag.Add("Motivo: $($info.Reason.GetType().Name): $($info.Reason.Message)") }
                foreach ($e in @($job.PS.Streams.Error)) {
                    $diag.Add("Error: $($e.Exception.GetType().Name): $($e.Exception.Message) | $($e.InvocationInfo.PositionMessage)")
                }
                foreach ($w in @($job.PS.Streams.Warning)) { $diag.Add("Aviso: $w") }
            } catch { $diag.Add("Sin datos del proceso: $($_.Exception.Message)") }
            $tail = ''
            try {
                if ($St.Sync.LogFile -and (Test-Path -LiteralPath $St.Sync.LogFile)) {
                    $lines = @(Get-Content -LiteralPath $St.Sync.LogFile -Encoding UTF8 -Tail 300)
                    $heads = @($lines | Where-Object { $_ -match '^\[\d{4}-\d\d-\d\d ' })
                    if ($heads.Count -gt 20) { $heads = $heads[($heads.Count - 20)..($heads.Count - 1)] }
                    if ($lines.Count -gt 60) { $lines = $lines[($lines.Count - 60)..($lines.Count - 1)] }
                    $tail = "Pasos:`n" + ($heads -join "`n") + "`n`nDetalle:`n" + ($lines -join "`n")
                }
            } catch { }
            $detail = "La tarea en segundo plano se detuvo sin entregar un resultado.`n`nDiagnóstico (v$($StConfig.Version))`n" + ($diag -join "`n") + "`n`nÚLTIMAS LÍNEAS DEL REGISTRO`n" + $tail
            Write-StLog ("Tarea sin resultado. " + ($diag -join ' | '))
            $res = @{ Ok = $false; Result = $null; Error = $detail }
        }
        try { $job.PS.Dispose() } catch { }
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
    $St.StepSummary = New-StXamlElement @'
<Border CornerRadius="14" Padding="14,10" Margin="0,0,0,6" Background="#12059669" BorderBrush="#4D059669" BorderThickness="1" Visibility="Collapsed">
  <StackPanel Orientation="Horizontal">
    <Border Width="32" Height="32" CornerRadius="16" Background="#059669">
      <TextBlock Text="&#xE73E;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="14" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <TextBlock x:Name="T" Foreground="#DBE5F1" FontSize="17" FontWeight="SemiBold" VerticalAlignment="Center" Margin="14,0,0,0"/>
  </StackPanel>
</Border>
'@
    [void]$St.UI.StackSteps.Children.Add($St.StepSummary)
    $n = 0
    foreach ($id in $Ids) {
        $n++
        $def = $StStepDefs[$id]
        $title = $def[0]
        if ($id -eq 'install' -and $St.Product) { $title = "Instalar $($St.Product.Name)" }
        $row = New-StXamlElement @"
<Border CornerRadius="14" Padding="14,10" Margin="0,0,0,6" Background="#00000000" BorderBrush="#00000000" BorderThickness="1">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="48"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Grid Width="32" Height="32" VerticalAlignment="Top">
      <Ellipse x:Name="Ring" Stroke="#8FA3BE" StrokeThickness="2" Fill="#00000000"/>
      <TextBlock x:Name="Num" Text="$n" Foreground="#AEBDD0" FontSize="15" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      <TextBlock x:Name="Glyph" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="14" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
      <Path x:Name="Spin" Width="32" Height="32" Data="M 16,1.5 A 14.5,14.5 0 0 1 30.5,16" Stroke="#DC2626" StrokeThickness="3.5" StrokeStartLineCap="Round" StrokeEndLineCap="Round" RenderTransformOrigin="0.5,0.5" Visibility="Collapsed">
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
      <TextBlock x:Name="Title" Foreground="#AEBDD0" FontSize="17" FontWeight="SemiBold" TextWrapping="Wrap"/>
      <TextBlock x:Name="Sub" Foreground="#C9D5E3" FontSize="15" TextWrapping="Wrap" Margin="0,2,0,0" Visibility="Collapsed"/>
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
    $St.StepMore = New-StXamlElement @'
<TextBlock Foreground="#AEBDD0" FontSize="16" Margin="62,4,0,0" Visibility="Collapsed"/>
'@
    [void]$St.UI.StackSteps.Children.Add($St.StepMore)
    Update-StStepLayout
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
    $spin.Visibility = 'Collapsed'; $glyph.Visibility = 'Collapsed'; $num.Visibility = 'Visible'; $sub.Visibility = 'Collapsed'
    $row.Background = $conv.ConvertFromString('#00000000'); $row.BorderBrush = $conv.ConvertFromString('#00000000')
    $ring.Fill = $conv.ConvertFromString('#00000000'); $ring.Stroke = $conv.ConvertFromString('#8FA3BE')
    $title.Foreground = $conv.ConvertFromString('#AEBDD0'); $sub.Text = $r.Sub
    switch ($State) {
        'running' {
            $spin.Visibility = 'Visible'; $num.Foreground = $conv.ConvertFromString('#FCA5A5'); $sub.Visibility = 'Visible'
            $ring.Stroke = $conv.ConvertFromString('#405776')
            $row.Background = $conv.ConvertFromString('#17243A'); $row.BorderBrush = $conv.ConvertFromString('#405776')
            $title.Foreground = $conv.ConvertFromString('#F8FAFC')
        }
        'done' {
            $num.Visibility = 'Collapsed'; $glyph.Visibility = 'Visible'; $glyph.Text = [string][char]0xE73E
            $glyph.Foreground = $conv.ConvertFromString('#FFFFFF')
            $ring.Fill = $conv.ConvertFromString('#059669'); $ring.Stroke = $conv.ConvertFromString('#059669')
            $title.Foreground = $conv.ConvertFromString('#DBE5F1')
        }
        'skip' {
            $num.Visibility = 'Collapsed'; $glyph.Visibility = 'Visible'; $glyph.Text = [string][char]0xE73E
            $glyph.Foreground = $conv.ConvertFromString('#FFFFFF')
            $ring.Fill = $conv.ConvertFromString('#059669'); $ring.Stroke = $conv.ConvertFromString('#059669')
            $sub.Text = 'No hizo falta'
        }
        'error' {
            $num.Visibility = 'Collapsed'; $glyph.Visibility = 'Visible'; $glyph.Text = [string][char]0xE711
            $glyph.Foreground = $conv.ConvertFromString('#FFFFFF'); $sub.Visibility = 'Visible'
            $ring.Fill = $conv.ConvertFromString('#B45309'); $ring.Stroke = $conv.ConvertFromString('#B45309')
            $title.Foreground = $conv.ConvertFromString('#FCD34D')
        }
        default { $num.Foreground = $conv.ConvertFromString('#AEBDD0') }
    }
    Update-StStepLayout
}

function Update-StStepLayout {
    # Muestra un resumen de lo terminado, el paso actual y solo los próximos pasos.
    if (-not $St.StepOrder -or -not $St.StepSummary) { return }
    $total = @($St.StepOrder).Count
    $finished = 0; $current = -1; $i = 0
    foreach ($id in $St.StepOrder) {
        $state = $St.StepRows[$id].State
        if ($state -in @('done', 'skip')) { $finished++ }
        elseif ($current -lt 0) { $current = $i }
        $i++
    }
    $shownPending = 0; $hidden = 0; $i = 0
    foreach ($id in $St.StepOrder) {
        $r = $St.StepRows[$id]
        if ($r.State -in @('done', 'skip')) { $r.Row.Visibility = 'Collapsed' }
        elseif ($r.State -in @('running', 'error') -or $i -eq $current) { $r.Row.Visibility = 'Visible' }
        elseif ($shownPending -lt 2) { $r.Row.Visibility = 'Visible'; $shownPending++ }
        else { $r.Row.Visibility = 'Collapsed'; $hidden++ }
        $i++
    }
    $St.StepSummary.Visibility = $(if ($finished -gt 0) { 'Visible' } else { 'Collapsed' })
    $St.StepSummary.FindName('T').Text = $(if ($finished -eq 1) { '1 tarea lista' } else { "$finished tareas listas" })
    $St.StepMore.Visibility = $(if ($hidden -gt 0) { 'Visible' } else { 'Collapsed' })
    $St.StepMore.Text = $(if ($hidden -eq 1) { 'y 1 paso más' } else { "y $hidden pasos más" })
    $stepNow = [Math]::Min($total, $finished + 1)
    $St.UI.TxtProgCount.Text = $(if ($finished -ge $total) { '¡Todo listo!' } else { "Paso $stepNow de $total" })
    $fraction = $(if ($total -gt 0) { ($finished + 0.5 * [int]($finished -lt $total)) / $total } else { 0 })
    $trackWidth = $St.UI.ProgTrack.ActualWidth
    if ($trackWidth -le 0) { $trackWidth = 600 }
    $target = [double]$trackWidth * [Math]::Min([double]1, [double]$fraction)
    if ($St.NoAnim) { $St.UI.ProgFill.Width = $target }
    else {
        $anim = [System.Windows.Media.Animation.DoubleAnimation]::new([double]$St.UI.ProgFill.ActualWidth, $target, [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(450)))
        $e = New-Object System.Windows.Media.Animation.QuinticEase; $e.EasingMode = 'EaseOut'; $anim.EasingFunction = $e
        $St.UI.ProgFill.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $anim)
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
    # La barra de avance se mantiene: muestra cuánto se completó.
    Update-StStepLayout
}

# ----- Errores ---------------------------------------------------------------------------------
function Show-StError {
    param([string]$Title = 'No pudimos completar la instalación', [string]$Detail, [string]$RetryLabel = 'Reintentar', [scriptblock]$OnRetry)
    $St.Busy = $false
    Stop-StProgressAnimation
    Invoke-StCue 'error'
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
    Set-StStub -Label $St.UI.TxtDataStubLabel -Year $St.UI.TxtDataStubYear -Product $p
    $St.UI.TxtDataNext.Text = $(if ($St.Mode -eq 'activate') { 'Activar Office' } else { 'Continuar' })
    Update-StDataValidity
    Show-StPage 'Data'
    [void]$St.UI.InKey.Focus()
}

function Set-StStub {
    param($Label, $Year, $Product)
    if (-not $Product) { return }
    if ($Product.Id -eq '365') { $Label.Text = 'MICROSOFT'; $Year.Text = '365' }
    else { $Label.Text = 'OFFICE'; $Year.Text = $Product.Year }
}

function Invoke-StPaste {
    param([string]$Target)
    $text = ''
    try { $text = [System.Windows.Clipboard]::GetText() } catch { }
    if (-not $text) {
        [void](Show-StDialog -Kind 'info' -Title 'No hay nada copiado' -Message 'Primero copia el dato desde tu correo: selecciónalo con el mouse, haz clic derecho y elige «Copiar». Después vuelve aquí y pulsa «Pegar».' -Buttons @('Entendido'))
        return
    }
    $found = Find-StPurchaseData -Text $text
    if ($found.Key) { $St.UI.InKey.Text = $found.Key }
    if ($found.Code) { $St.UI.InCode.Text = $found.Code }
    if (-not $found.Key -and -not $found.Code) {
        if ($Target -eq 'key') { $St.UI.InKey.Text = $text.Trim() } else { $St.UI.InCode.Text = $text.Trim() }
    }
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
    $ui.TxtConfirmLead.Text = ''
    $ui.TxtConfirmLead.Visibility = 'Collapsed'
    $St.ConfirmAction = 'install'
    Set-StStub -Label $ui.TxtConfirmStubLabel -Year $ui.TxtConfirmStubYear -Product $p

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
        $ui.TxtConfirmLead.Visibility = 'Visible'
        $items.Add('E73E|No se desinstala ni se cambia nada en tu Office.')
        $items.Add('E77B|No es necesario cerrar tu cuenta personal de OneDrive o Teams.')
        $items.Add('E8A5|Te guiaremos para agregar la cuenta de SafeTech en Word.')
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
        $items.Add('E8A5|**Guarda los documentos que tengas abiertos**: vamos a cerrar Word, Excel y Outlook.')
        $items.Add('E8B7|Tus documentos y archivos **no se borran**.')
        $items.Add('E701|Deja el computador **enchufado y con Internet**: Office se descarga desde Microsoft (unos 3 GB).')
        if ($p.NeedsKey) { $items.Add('E8D7|Al terminar, **activamos Office solos** con tu código SafeTech.') }
        else { $items.Add('E77B|Mientras se instala, puedes ir preparando tu cuenta con el tutorial.') }
        if ($p.Id -eq '365' -and $state.HasAny) {
            $items.Add('E77B|Si usabas otra cuenta en Office, puede que tengas que volver a elegirla después.')
        }
    }

    $ui.StackConfirmList.Children.Clear()
    foreach ($t in $items) {
        $parts = $t.Split('|', 2)
        $row = New-StXamlElement @"
<Grid Margin="0,0,0,8">
  <Grid.ColumnDefinitions><ColumnDefinition Width="46"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Width="34" Height="34" CornerRadius="10" Background="#20304A" VerticalAlignment="Top">
    <TextBlock x:Name="I" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="#C9D5E3" HorizontalAlignment="Center" VerticalAlignment="Center"/>
  </Border>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="17" TextWrapping="Wrap" LineHeight="26" VerticalAlignment="Center" Margin="0,3,0,0"/>
</Grid>
"@
        $row.FindName('I').Text = [string][char][Convert]::ToInt32($parts[0], 16)
        Set-StRichText -Target $row.FindName('T') -Text $parts[1]
        [void]$ui.StackConfirmList.Children.Add($row)
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
    $St.UI.TxtWhile.Text = 'No apagues ni reinicies el computador. Verás una ventana de Microsoft con el avance: cuando diga «¡Hemos terminado!», presiona «Cerrar» o espera, la cerramos por ti.'
    $St.UI.TxtProgTitle.Text = "Instalando $($p.Name)"
    $St.UI.TxtProgDetail.Text = 'Preparando...'
    $St.Sync.Detail = 'Preparando...'
    $St.UI.TxtElapsedNote.Text = 'Suele tardar entre 10 y 40 minutos, según tu Internet.'
    $St.UI.PanelTutorialPromo.Visibility = $(if ($p.Id -eq '365') { 'Visible' } else { 'Collapsed' })
    $St.UI.ChipInstall.Visibility = $(if ($p.Id -eq '365') { 'Visible' } else { 'Collapsed' })
    $St.StartTime = Get-Date
    $St.Busy = $true
    $St.InstallDone = $false
    Set-StProgStamp -Text 'PENDIENTE' -Kind 'pending'
    Show-StPage 'Progress'
    Invoke-StCue 'instalando'

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
        Set-StProgStamp -Text 'INSTALADO' -Kind 'done' -Animate
        Show-StWindowFront
        if ($St.Product.NeedsKey) {
            # Sin ventana intermedia: la activación sigue sola (antes un aviso quedaba escondido detrás de Microsoft).
            Invoke-StCue 'instalado'
            Start-StActivationFlow -KeepSteps
        } else {
            Invoke-StCue 'instalado-365'
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
        $St.UI.TxtWhile.Text = 'No apagues el computador. La activación suele tardar menos de un minuto.'
        $St.StartTime = Get-Date
        $St.UI.PanelTutorialPromo.Visibility = 'Collapsed'
    }
    $St.UI.TxtProgTitle.Text = "Activando $($p.Name)"
    $St.UI.TxtElapsedNote.Text = $(if ($KeepSteps) { 'Ya casi: solo falta activar Office (menos de un minuto).' } else { 'La activación tarda menos de un minuto.' })
    $St.Busy = $true
    if (-not $KeepSteps) { Set-StProgStamp -Text 'PENDIENTE' -Kind 'pending' }
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
        if ($kind -eq 'code') {
            # Un solo paso: el cliente corrige el código en la misma ventana.
            $hint = 'Revisa que empiece con ST- y que no confundas 0 (cero) con la letra O.'
            $msg = "El código no es válido o ya se usó.`n$message"
            while ($true) {
                $f = New-StDialogField -Label 'Código de activación SafeTech' -Placeholder 'ST-XXXXXX' -Value $St.Act.Code -MaxLength 40 -Hint $hint
                Invoke-StCue 'codigo'
                $choice = Show-StDialog -Kind 'warn' -Title 'Tu código SafeTech no funcionó' -Message $msg -Extra $f.Panel `
                    -Buttons @('Cancelar', 'Probar este código') -Primary 'Probar este código' -LinkLabel 'Activar de otra forma' -Width 640
                if ($choice -eq 'Activar de otra forma') { Show-StManualActivation; return }
                if ($choice -ne 'Probar este código') { Stop-StActivationPending; return }
                $newCode = ConvertTo-StAccessCode $f.Box.Text
                if ($newCode) { $St.Act.Code = $newCode; Request-StCidFlow; return }
                $hint = 'Ese código está incompleto. Debe tener el formato ST-XXXXXX, por ejemplo ST-AB12CD.'
                $msg = 'Escribe el código tal como aparece en tu correo.'
            }
        }
        Invoke-StCue 'atencion'
        $choice = Show-StDialog -Kind 'warn' -Title 'No pudimos conectarnos con SafeTech' `
            -Message ($message + "`n`nRevisa tu Internet y vuelve a intentarlo. No se gastó ningún crédito si no llegó la respuesta.") `
            -Buttons @('Cancelar', 'Reintentar') -Primary 'Reintentar' -LinkLabel 'Activar de otra forma' -Width 640
        switch ($choice) {
            'Reintentar' { Request-StCidFlow }
            'Activar de otra forma' { Show-StManualActivation }
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

$StWelcomeCards = @(
    @{ Icon = [char]0xE896; Color = '#6EE7B7'; Back = '#26059669'; Title = 'Office original de Microsoft'
       Text = 'Se descarga directo desde Microsoft, sin cambios ni archivos modificados.' },
    @{ Icon = [char]0xE72E; Color = '#6EE7B7'; Back = '#26059669'; Title = 'Nada pirata'
       Text = 'No usamos cracks ni activadores. Solo tu licencia, como lo haría un técnico.' },
    @{ Icon = [char]0xE8D7; Color = '#FCA5A5'; Back = '#26DC2626'; Title = 'Se activa con tu código'
       Text = 'Con tu clave y tu código SafeTech activamos Office por ti. No copias números.' },
    @{ Icon = [char]0xE8B7; Color = '#93C5FD'; Back = '#262563EB'; Title = 'Tus archivos no se tocan'
       Text = 'Tus documentos, fotos y programas quedan igual que antes.' }
)

function Show-StAbout {
    param([switch]$Welcome)
    $panel = New-StXamlElement @'
<StackPanel>
  <UniformGrid x:Name="Cards" Columns="2"/>
  <Border CornerRadius="14" Background="#1AF59E0B" BorderBrush="#80F59E0B" BorderThickness="1.5" Padding="18,12" Margin="0,4,0,0">
    <Grid>
      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <TextBlock Text="&#xE7EF;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="20" Foreground="#FCD34D" VerticalAlignment="Center" Margin="0,0,14,0"/>
      <TextBlock Grid.Column="1" FontSize="17" Foreground="#F8FAFC" TextWrapping="Wrap" VerticalAlignment="Center">Windows te pedirá permiso para hacer cambios: responde <Bold>«Sí»</Bold>. Es normal.</TextBlock>
    </Grid>
  </Border>
  <TextBlock Text="SafeTech no es distribuidor oficial de Microsoft. Microsoft y Office son marcas de Microsoft Corporation." FontSize="15" Foreground="#AEBDD0" TextWrapping="Wrap" Margin="2,12,0,0"/>
</StackPanel>
'@
    $cards = $panel.FindName('Cards')
    $i = 0
    foreach ($c in $StWelcomeCards) {
        $m = $(if ($i % 2 -eq 0) { '0,0,7,14' } else { '7,0,0,14' })
        $row = New-StXamlElement @"
<Border CornerRadius="16" Background="#17243A" BorderBrush="#2A3C57" BorderThickness="1.5" Padding="18,16" Margin="$m">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="52" Height="52" CornerRadius="26" Background="$($c.Back)" VerticalAlignment="Top">
      <TextBlock x:Name="Ico" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="22" Foreground="$($c.Color)" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <StackPanel Grid.Column="1" Margin="16,0,0,0">
      <TextBlock x:Name="Ttl" Foreground="#F8FAFC" FontSize="19" FontWeight="SemiBold" TextWrapping="Wrap"/>
      <TextBlock x:Name="Txt" Foreground="#C9D5E3" FontSize="17" TextWrapping="Wrap" LineHeight="24" Margin="0,4,0,0"/>
    </StackPanel>
  </Grid>
</Border>
"@
        $row.FindName('Ico').Text = [string]$c.Icon
        $row.FindName('Ttl').Text = $c.Title
        $row.FindName('Txt').Text = $c.Text
        [void]$cards.Children.Add($row)
        $i++
    }
    if ($Welcome) { Invoke-StCue 'bienvenida' -Once }
    $label = $(if ($Welcome) { 'Entendido, empezar' } else { 'Entendido' })
    $title = $(if ($Welcome) { '¡Hola! Así funciona este instalador' } else { '¿Cómo funciona este instalador?' })
    return (Show-StDialog -Kind 'info' -Title $title -Message 'Léelo con calma: son solo 4 cosas.' -Extra $panel -Buttons @($label) -Width 920 `
        -Image (New-StBotForDialog) -LinkLabel 'Ver detalles técnicos' -OnLink { Show-StAboutDetails })
}

function Show-StAboutDetails {
    $panel = New-StXamlElement @'
<ScrollViewer MaxHeight="440" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,8,0">
  <StackPanel x:Name="List"/>
</ScrollViewer>
'@
    $list = $panel.FindName('List')
    foreach ($sec in $StAboutSections) {
        $row = New-StXamlElement @"
<Border CornerRadius="14" Background="#0C1626" BorderBrush="#2A3C57" BorderThickness="1" Padding="16,14" Margin="0,0,0,10">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="46"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="34" Height="34" CornerRadius="10" Background="#20304A" VerticalAlignment="Top">
      <TextBlock x:Name="Ico" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="$($sec.Color)" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <StackPanel Grid.Column="1">
      <TextBlock x:Name="Ttl" Foreground="#F8FAFC" FontSize="17" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,4,0,6"/>
      <TextBlock x:Name="Txt" Foreground="#DBE5F1" FontSize="16" TextWrapping="Wrap" LineHeight="23"/>
    </StackPanel>
  </Grid>
</Border>
"@
        $row.FindName('Ico').Text = [string]$sec.Icon
        $row.FindName('Ttl').Text = $sec.Title
        Set-StRichText -Target $row.FindName('Txt') -Text $sec.Text
        [void]$list.Children.Add($row)
    }
    $r = Show-StDialog -Kind 'info' -Title 'Detalles técnicos' -Message '' -Extra $panel -Buttons @('Cerrar', 'Ver el código en GitHub') -Width 760
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
    <TextBlock Text="1. Tu ID de instalación" Foreground="#F8FAFC" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,6"/>
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
    $St.UI.TxtDoneTitle.Text = "¡Listo! $($p.Name) quedó activado"
    $St.UI.TxtDoneText.Text = 'Ya puedes usar Word, Excel y PowerPoint en este computador.'
    Set-StDoneList @('E8A5|Word, Excel, PowerPoint y más, **instalados**', 'E8D7|**Activado** con tu licencia', 'E895|Se actualiza solo con Microsoft', 'E721|Para abrirlos después: botón **Inicio** y escribe **Word**')
    $St.UI.TxtDoneNote.Text = 'Guarda el correo de tu compra. Si algún día reinstalas Windows, escríbenos y te ayudamos.'
    Show-StDonePage -Stamp 'ACTIVADO'
}

function Set-StProgStamp {
    param([string]$Text, [ValidateSet('pending', 'done')][string]$Kind = 'pending', [switch]$Animate)
    $conv = New-Object System.Windows.Media.BrushConverter
    $St.UI.TxtProgStamp.Text = $Text
    $c = $(if ($Kind -eq 'done') { @('#34D399', '#34D399') } else { @('#F59E0B', '#FCD34D') })
    $St.UI.ProgStamp.BorderBrush = $conv.ConvertFromString($c[0])
    $St.UI.TxtProgStamp.Foreground = $conv.ConvertFromString($c[1])
    Set-StStub -Label $St.UI.TxtProgStubLabel -Year $St.UI.TxtProgStubYear -Product $St.Product
    if (-not $Animate -or $St.NoAnim) { return }
    $ease = New-Object System.Windows.Media.Animation.BackEase; $ease.EasingMode = 'EaseOut'; $ease.Amplitude = 0.35
    $dur = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(300))
    foreach ($prop in @([System.Windows.Media.ScaleTransform]::ScaleXProperty, [System.Windows.Media.ScaleTransform]::ScaleYProperty)) {
        $a = [System.Windows.Media.Animation.DoubleAnimation]::new([double]1.7, [double]1, $dur); $a.EasingFunction = $ease
        $St.UI.ProgStampScale.BeginAnimation($prop, $a)
    }
}

function Show-StDonePage {
    param([string]$Stamp)
    Set-StStub -Label $St.UI.TxtDoneStubLabel -Year $St.UI.TxtDoneStubYear -Product $St.Product
    $St.UI.TxtDoneStamp.Text = $Stamp
    $St.UI.DoneNote.Visibility = 'Visible'
    Show-StPage 'Done'
    Show-StWindowFront
    Invoke-StCue $(if ($Stamp -eq 'ACTIVADO') { 'activado' } else { 'conectado' })
    if ($St.NoAnim) { return }
    # El sello cae sobre el comprobante.
    $St.UI.DoneStamp.Opacity = 0
    $ease = New-Object System.Windows.Media.Animation.BackEase; $ease.EasingMode = 'EaseOut'; $ease.Amplitude = 0.35
    $dur = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(320))
    $begin = [TimeSpan]::FromMilliseconds(260)
    $sx = [System.Windows.Media.Animation.DoubleAnimation]::new([double]1.6, [double]1, $dur); $sx.EasingFunction = $ease; $sx.BeginTime = $begin
    $sy = [System.Windows.Media.Animation.DoubleAnimation]::new([double]1.6, [double]1, $dur); $sy.EasingFunction = $ease; $sy.BeginTime = $begin
    $op = [System.Windows.Media.Animation.DoubleAnimation]::new([double]0, [double]1, [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(180))); $op.BeginTime = $begin
    $St.UI.DoneStampScale.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleXProperty, $sx)
    $St.UI.DoneStampScale.BeginAnimation([System.Windows.Media.ScaleTransform]::ScaleYProperty, $sy)
    $St.UI.DoneStamp.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $op)
}

function Set-StDoneList {
    # Cada elemento: "<código del ícono>|<texto con **negrita**>".
    param([string[]]$Items)
    $St.UI.StackDoneList.Children.Clear()
    foreach ($t in $Items) {
        $parts = $t.Split('|', 2)
        $row = New-StXamlElement @"
<Grid Margin="0,5">
  <Grid.ColumnDefinitions><ColumnDefinition Width="46"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Width="34" Height="34" CornerRadius="17" Background="#26059669">
    <TextBlock x:Name="I" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="15" Foreground="#6EE7B7" HorizontalAlignment="Center" VerticalAlignment="Center"/>
  </Border>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="17" VerticalAlignment="Center" TextWrapping="Wrap"/>
</Grid>
"@
        $row.FindName('I').Text = [string][char][Convert]::ToInt32($parts[0], 16)
        Set-StRichText -Target $row.FindName('T') -Text $parts[1]
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
       Note = 'si este computador es de tu **empresa, colegio o universidad**, consulta antes con quien lo administra: la cuenta puede entrar en conflicto con sus políticas.'; NoteKind = 'warn'
       Image = $null; Actions = @(); NeedsOffice = $false },
    @{ Short = 'Entrar al portal'; Title = 'Inicia sesión en el portal de Microsoft 365'
       Body = @('Haz clic en **«Abrir el portal»**. Se abrirá tu navegador.',
                'Escribe el **correo entregado por SafeTech** y presiona **Siguiente**.',
                'Escribe la **contraseña temporal** que te enviamos.',
                '!**No uses tu correo personal** en este paso.')
       Note = 'Si el navegador entra solo con otra cuenta tuya, ciérrala arriba a la derecha («Cerrar sesión») y vuelve a pulsar «Abrir el portal». Si te enredas, escríbenos.'; NoteKind = 'info'
       Image = 'assets/tutorial/02-portal-inicio-sesion.jpg'; Actions = @('portal', 'copyPortal'); NeedsOffice = $false },
    @{ Short = 'Contraseña'; Title = 'Cambia la contraseña temporal'
       Body = @('La primera vez, Microsoft te pedirá **crear una contraseña nueva**.',
                'En el primer campo escribe la **contraseña temporal**; en los otros dos, **tu contraseña nueva**.',
                'Usa mayúsculas, minúsculas, números y un símbolo. Ejemplo: Perro#Azul2026 (inventa la tuya).')
       Note = '**anótala ahora en un papel o sácale una foto.** SafeTech no puede verla ni recuperarla; si la pierdes, pierdes la cuenta.'; NoteKind = 'critical'
       Image = 'assets/tutorial/03-cambiar-contrasena.jpg'; Actions = @(); NeedsOffice = $false; Confirm = 'Ya anoté mi nueva contraseña' },
    @{ Short = 'Seguridad'; Title = '¿Te apareció este mensaje?'
       Body = @('Después de cambiar la contraseña, algunas cuentas muestran **«Vamos a garantizar la seguridad de la cuenta»** y piden instalar Microsoft Authenticator.',
                'Otras cuentas no lo piden y entran directo. Depende de la cuenta que te entregamos.',
                '!**Elige lo que te pasó a ti** con uno de los dos botones de abajo.')
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
       Note = '**no borres la app ni la cuenta de tu celular**: Microsoft puede pedirte aprobar el ingreso otra vez. Si cambias de celular, **avísanos antes por WhatsApp**.'; NoteKind = 'critical'
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
       Note = '**no elijas «Sí».** La cuenta quedaría ligada al computador como cuenta de empresa y cambiaría su configuración. Esta cuenta solo debe usarse para activar Office.'; NoteKind = 'critical'
       Image = 'assets/tutorial/09-solo-esta-aplicacion.jpg'; Actions = @(); NeedsOffice = $true },
    @{ Short = 'Comprobar'; Title = 'Comprueba que Office quedó activado'
       Body = @('En Word ve a **Archivo > Cuenta**.',
                '!Debe decir **«Producto de suscripción»** y **«Aplicaciones de Microsoft 365»**.',
                'Si aparece «Producto sin licencia», **cierra todos los programas de Office**, vuelve a abrir Word y revisa otra vez.')
       Note = '**no guardes archivos importantes en el OneDrive de esta cuenta**: Microsoft puede reducir o eliminar ese espacio sin aviso. Usa tu OneDrive personal.'; NoteKind = 'warn'
       Image = 'assets/tutorial/10-producto-activado.jpg'; Actions = @('check', 'word'); NeedsOffice = $true }
)

function Initialize-StTutorialNav {
    $St.UI.StackTutSteps.Children.Clear()
    $St.TutRows = @()
    $i = 0
    foreach ($s in $StTutorial) {
        $i++
        $row = New-StXamlElement @"
<Button Cursor="Hand" Margin="0,0,0,2" FocusVisualStyle="{x:Null}" Background="#00000000" BorderThickness="0" HorizontalContentAlignment="Stretch">
  <Button.Template>
    <ControlTemplate TargetType="Button">
      <Border x:Name="Bd" CornerRadius="12" Padding="10,5" MinHeight="42" Background="{TemplateBinding Background}" BorderBrush="#00000000" BorderThickness="2"><ContentPresenter VerticalAlignment="Center"/></Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#17243A"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Bd" Property="BorderBrush" Value="#FCD34D"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Button.Template>
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="40"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Grid Width="28" Height="28">
      <Ellipse x:Name="Dot" Stroke="#8FA3BE" StrokeThickness="2"/>
      <TextBlock x:Name="Num" Text="$i" FontSize="13" FontWeight="Bold" Foreground="#AEBDD0" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Grid>
    <TextBlock Grid.Column="1" x:Name="Lbl" Foreground="#AEBDD0" FontSize="16" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
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
    $ui.TxtTutTitle.Text = $s.Title
    $ui.StackTutBody.Children.Clear()
    $k = 0
    foreach ($line in $s.Body) {
        $k++
        $must = $line.StartsWith('!')
        $text = $line.TrimStart('!')
        if ($must) {
            $row = New-StXamlElement @"
<Border CornerRadius="14" Background="#20304A" BorderBrush="#405776" BorderThickness="1.5" Padding="14,12" Margin="-14,0,0,12">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="40"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="28" Height="28" CornerRadius="14" Background="#DC2626" VerticalAlignment="Top">
      <TextBlock Text="&#xE72A;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <TextBlock Grid.Column="1" x:Name="T" Foreground="#F8FAFC" FontSize="17" FontWeight="SemiBold" TextWrapping="Wrap" LineHeight="26" Margin="0,1,0,0"/>
  </Grid>
</Border>
"@
        } else {
            $row = New-StXamlElement @"
<Grid Margin="0,0,0,12">
  <Grid.ColumnDefinitions><ColumnDefinition Width="26"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Ellipse Width="8" Height="8" Fill="#AEBDD0" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="4,10,0,0"/>
  <TextBlock Grid.Column="1" x:Name="T" Foreground="#DBE5F1" FontSize="17" TextWrapping="Wrap" LineHeight="26"/>
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
    $noteMap = @{
        info     = @('#17243A', '#405776', '#AEBDD0', [char]0xE946, '', '1.5')
        warn     = @('#1AF59E0B', '#80F59E0B', '#FCD34D', [char]0xE7BA, 'Importante: ', '1.5')
        critical = @('#26F59E0B', '#F59E0B', '#FCD34D', [char]0xE7BA, 'Muy importante: ', '2.5')
    }[$noteKind]
    Set-StRichText -Target $ui.TxtTutNote -Text $(if ($noteMap[4]) { '**' + $noteMap[4] + '**' + $noteText } else { $noteText })
    $ui.TutNote.Background = $conv.ConvertFromString($noteMap[0])
    $ui.TutNote.BorderBrush = $conv.ConvertFromString($noteMap[1])
    $ui.TutNote.BorderThickness = $noteMap[5]
    $ui.TutNoteIcon.Foreground = $conv.ConvertFromString($noteMap[2])
    $ui.TutNoteIcon.Text = [string]$noteMap[3]
    $ui.TxtTutNoteLabel.Visibility = 'Collapsed'
    $ui.TxtTutNote.Foreground = $conv.ConvertFromString($(if ($noteKind -eq 'info') { '#DBE5F1' } else { '#F8FAFC' }))
    # El aviso crítico va antes de los botones (se lee antes de confirmar); los consejos, después.
    [void]$ui.TutContent.Children.Remove($ui.TutNote)
    $pos = $ui.TutContent.Children.IndexOf($ui.StackTutActions)
    if ($noteKind -eq 'critical') { $ui.TutContent.Children.Insert($pos, $ui.TutNote) } else { $ui.TutContent.Children.Insert($pos + 1, $ui.TutNote) }
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
            'authYes' { $b.Content = 'Sí, me apareció'; $b.Style = $St.Window.FindResource($(if ($St.Tut.Auth -eq $true) { 'BtnPrimary' } else { 'BtnSoft' })); $b.MinWidth = 220 }
            'authNo' { $b.Content = 'No, entré directo'; $b.Style = $St.Window.FindResource($(if ($St.Tut.Auth -eq $false) { 'BtnPrimary' } else { 'BtnSoft' })); $b.MinWidth = 220 }
        }
        $b.Add_Click({ param($sender, $e) Invoke-StTutorialAction ([string]$sender.Tag) $sender })
        [void]$ui.StackTutActions.Children.Add($b)
    }
    if ($s.ContainsKey('Confirm') -and $s.Confirm) {
        $chk = New-Object System.Windows.Controls.CheckBox
        $chk.Style = $St.Window.FindResource('Check')
        $chk.Content = $s.Confirm
        $chk.Margin = '0,4,0,10'
        $chk.MinWidth = 380
        $chk.IsChecked = [bool]$St.Tut.Checked[$Index]
        $chk.Tag = $Index
        $chk.Add_Checked({ param($sender, $e) $St.Tut.Checked[[int]$sender.Tag] = $true; $St.UI.BtnTutNext.IsEnabled = $true })
        $chk.Add_Unchecked({ param($sender, $e) $St.Tut.Checked[[int]$sender.Tag] = $false; $St.UI.BtnTutNext.IsEnabled = $false })
        [void]$ui.StackTutActions.Children.Add($chk)
    }
    $ui.StackTutActions.Visibility = $(if ($ui.StackTutActions.Children.Count) { 'Visible' } else { 'Collapsed' })

    $img = $null
    if ($s.Image) { $img = Get-StImage $s.Image }
    $ui.ImgTut.Source = $img
    $ui.TutImageFrame.Visibility = $(if ($img) { 'Visible' } else { 'Collapsed' })
    [System.Windows.Controls.Grid]::SetColumnSpan($ui.TutScroll, $(if ($img) { 1 } else { 2 }))
    $ui.TutContent.MaxWidth = $(if ($img) { [double]::PositiveInfinity } else { 760 })

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
            $dot.Fill = $conv.ConvertFromString('#00000000'); $dot.Stroke = $conv.ConvertFromString('#8FA3BE')
            $num.Foreground = $conv.ConvertFromString('#AEBDD0'); $num.Text = [string]($i + 1); $num.FontFamily = 'Segoe UI'
            $lbl.Foreground = $conv.ConvertFromString('#AEBDD0'); $lbl.FontWeight = 'Normal'
            if (Test-StTutorialBlocked $i) { $lbl.Foreground = $conv.ConvertFromString('#8FA3BE') }
        }
        if ($i -ne $Index -and (Test-StAuthSkipped $i)) {
            $dot.Fill = $conv.ConvertFromString('#00000000'); $dot.Stroke = $conv.ConvertFromString('#2A3C57')
            $num.Foreground = $conv.ConvertFromString('#8FA3BE'); $num.Text = [string][char]0xE738; $num.FontFamily = 'Segoe Fluent Icons, Segoe MDL2 Assets'
            $lbl.Foreground = $conv.ConvertFromString('#8FA3BE'); $lbl.FontWeight = 'Normal'
            $lbl.Text = $StTutorial[$i].Short
            $r.ToolTip = 'No hizo falta en tu cuenta'
        } else {
            $lbl.Text = $StTutorial[$i].Short
        }
    }
    # En el primer paso, «Atrás» vuelve a la pantalla principal (salvo mientras se instala Microsoft 365).
    $ui.BtnTutBack.IsEnabled = ($Index -gt 0 -or -not (Test-StInstalling365))
    $ui.TxtTutNext.Text = $(if ($Index -eq $n - 1) { 'Terminar' } else { 'Siguiente' })
    # En la pregunta de seguridad se responde con los dos botones; en la contraseña hay que confirmar que se anotó.
    $ui.BtnTutNext.Visibility = $(if ($Index -eq 3 -and $null -eq $St.Tut.Auth) { 'Collapsed' } else { 'Visible' })
    $ui.BtnTutNext.IsEnabled = -not ($s.ContainsKey('Confirm') -and $s.Confirm -and -not $St.Tut.Checked[$Index])
    $ui.TxtTutProgress.Text = "Paso $($Index + 1) de $n"
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
    if (-not $St.Product) { $St.Product = Get-StProduct '365' }
    $St.UI.TxtDoneTitle.Text = '¡Listo! Microsoft 365 quedó conectado'
    $St.UI.TxtDoneText.Text = 'Ya puedes usar Word, Excel, PowerPoint y Outlook con la cuenta que te entregamos.'
    if ($St.Tut.Auth -eq $false) {
        Set-StDoneList @('E8A5|Aplicaciones de Microsoft 365 **instaladas**', 'E77B|Cuenta de SafeTech **conectada** en Office', 'E72E|Contraseña nueva creada')
        $St.UI.TxtDoneNote.Text = 'Guarda tu nueva contraseña: SafeTech no puede recuperarla. La garantía dura 1 año desde la compra.'
    } else {
        Set-StDoneList @('E8A5|Aplicaciones de Microsoft 365 **instaladas**', 'E77B|Cuenta de SafeTech **conectada** en Office', 'E8EA|Microsoft Authenticator listo en tu celular')
        $St.UI.TxtDoneNote.Text = 'Guarda tu nueva contraseña y no borres Microsoft Authenticator de tu celular. La garantía dura 1 año desde la compra.'
    }
    if ($St.StepRows -and $St.StepRows.ContainsKey('account')) { Set-StUiStep 'account' 'done' }
    Show-StDonePage -Stamp 'CONECTADO'
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
$St.Bot = New-StBot
if ($St.Bot) { $St.UI.BotHost.Content = $St.Bot.View; Start-StBotIdle $St.Bot } else { $St.UI.BotHost.Visibility = 'Collapsed' }
$icon = Get-StImage 'assets/icono.png'
if ($icon) { $St.Window.Icon = $icon }

# Sonido: la voz espera a que termine el tono de aviso.
$St.VoiceTimer = New-Object System.Windows.Threading.DispatcherTimer
$St.VoiceTimer.Add_Tick({ $St.VoiceTimer.Stop(); Start-StVoice $St.PendingVoice })
Set-StSound -On (-not (Test-Path -LiteralPath (Join-Path $StConfig.DataDir 'sonido-apagado')))

# Descarga en segundo plano los sonidos, la voz y las imagenes del tutorial
if (-not $env:SAFETECH_OFFICE_ASSETS -and $StMode -eq 'app') {
    $assetList = @(Get-StCueAssets) + @($StTutorial | Where-Object { $_.Image } | ForEach-Object { $_.Image })
    Start-StAsync -Params @{ Paths = $assetList; Cache = $St.CacheDir } -Work { Save-StAssets -Paths $P.Paths -CacheDir $P.Cache }
}
if ($StMode -eq 'app') {
    Start-StAsync -Work { Get-StSignInIssues } -OnDone { param($ok, $r, $err) if ($ok) { $St.SignInIssues = @($r) } }
}

# ----- Eventos ---------------------------------------------------------------------------------
$ui = $St.UI
$ui.BtnMin.Add_Click({ $St.Window.WindowState = 'Minimized' })
$ui.BtnMax.Add_Click({ $St.Window.WindowState = $(if ($St.Window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }) })
$St.Window.Add_StateChanged({ Update-StWindowFrame })
$St.Window.Add_SizeChanged({ Update-StScale })
$St.Window.Add_ContentRendered({
    if ($St.SelfTest -or $St.AboutShown) { return }
    $St.AboutShown = $true
    Show-StAbout -Welcome
})
$ui.BtnClose.Add_Click({ $St.Window.Close() })
$ui.BtnHelpCenter.Add_Click({ Open-StUnelevated $StConfig.HelpUrl })
$ui.BtnAbout.Add_Click({ Show-StAbout })
$ui.BtnSound.Add_Click({ Set-StSound -On (-not $St.SoundOn) -Save; if ($St.SoundOn) { Invoke-StCue 'atencion' } })
$ui.BtnHelpWa.Add_Click({ Open-StWhatsApp -Message 'Hola SafeTech, necesito ayuda con el instalador de Office.' })

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
        $raw = $St.UI.InKey.Text
        if ($raw.Length -gt 29) {
            $found = Find-StPurchaseData -Text $raw
            if ($found.Key) { $raw = $found.Key }
            if ($found.Code -and -not $St.UI.InCode.Text) { $St.UI.InCode.Text = $found.Code }
        }
        if ($raw -ne $St.UI.InKey.Text) { $St.UI.InKey.Text = $raw }
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
    if ($t.Length -gt 12) {
        $found = Find-StPurchaseData -Text $t
        if ($found.Code) { $St.UI.InCode.Text = $found.Code; return }
    }
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
$ui.BtnPasteKey.Add_Click({ Invoke-StPaste -Target 'key' })
$ui.BtnPasteCode.Add_Click({ Invoke-StPaste -Target 'code' })
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
            $c = Show-StDialog -Kind 'success' -Title 'Informe guardado' -Message "Lo dejamos en tu Escritorio:`n$zip`n`nEnvíalo por WhatsApp al $($StConfig.WhatsApp) o a $($StConfig.Email)." -Buttons @('Mostrar el archivo', 'Cerrar')
            if ($c -eq 'Mostrar el archivo') { Start-Process (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList "/select,`"$zip`"" }
        } else {
            [void](Show-StDialog -Kind 'error' -Title 'No pudimos crear el ZIP' -Message "$err`n`nLos registros están en: $($St.LogDir)" -Buttons @('Cerrar'))
        }
    }
})

$ui.BtnErrWa.Add_Click({
    Initialize-StLogs
    $St.UI.BtnErrWa.IsEnabled = $false
    $name = $(if ($St.Product) { $St.Product.Name } else { 'Office' })
    $St.WaProduct = $name
    Start-StAsync -Params @{ Dir = $St.LogDir } -Work { New-StDiagnosticZip -LogDir $P.Dir } -OnDone {
        param($ok, $zip, $err)
        $St.UI.BtnErrWa.IsEnabled = $true
        $msg = "Hola SafeTech, el instalador de $($St.WaProduct) no pudo terminar."
        if ($ok -and $zip) {
            $msg += ' Te envío el archivo del informe que quedó en mi Escritorio.'
            try { Start-Process (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList "/select,`"$zip`"" } catch { }
        }
        Open-StWhatsApp -Message $msg
        if ($ok -and $zip) {
            [void](Show-StDialog -Kind 'info' -Title 'Adjunta el informe en WhatsApp' -Message "Abrimos WhatsApp y la carpeta con el informe.`n`nEn WhatsApp, pulsa el clip (Adjuntar), elige «Documento» y selecciona el archivo que empieza con SafeTech-Diagnostico." -Buttons @('Entendido'))
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
    function Save-StShotAt {
        # Dibuja la ventana como si estuviera maximizada en una pantalla de W x H (el runner tiene pantalla chica).
        param([string]$Name, [int]$W = 1366, [int]$H = 728)
        $root = $St.Window.Content
        $St.Window.Content = $null
        $holder = New-Object System.Windows.Controls.Border
        $holder.Background = $St.Window.Background
        try { $holder.Resources.MergedDictionaries.Add($St.Window.Resources) } catch { Write-Host "Aviso recursos: $($_.Exception.Message)" }
        [System.Windows.Documents.TextElement]::SetFontFamily($holder, $St.Window.FontFamily)
        [System.Windows.Documents.TextElement]::SetFontSize($holder, $St.Window.FontSize)
        [System.Windows.Documents.TextElement]::SetForeground($holder, $St.Window.Foreground)
        [System.Windows.Media.TextOptions]::SetTextFormattingMode($holder, 'Ideal')
        $holder.Child = $root
        $St.UI.RootBorder.Margin = '0'; $St.UI.RootBorder.BorderThickness = '0'
        $scale = Get-StScale $W $H
        $St.UI.RootScale.ScaleX = $scale; $St.UI.RootScale.ScaleY = $scale
        try {
            $holder.Measure((New-Object System.Windows.Size $W, $H))
            $holder.Arrange((New-Object System.Windows.Rect 0, 0, $W, $H))
            $holder.UpdateLayout()
            if ($St.CurrentPage -eq 'Progress') { Update-StStepLayout; $holder.UpdateLayout() }
            if ($St.CurrentPage -eq 'Tutorial') { Set-StTutorialStep -Index $St.Tut.Index; $holder.UpdateLayout() }
            $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap $W, $H, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
            $rtb.Render($holder)
            $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
            $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
            $fs = [IO.File]::Create((Join-Path $out "$Name.png"))
            $enc.Save($fs); $fs.Close()
            Write-Host "Captura: $Name ($W x $H, escala $([Math]::Round($scale, 2)))"
        } finally {
            $holder.Child = $null
            $St.Window.Content = $root
        }
    }
    Show-StPage 'Home'
    Save-StShotAt '01-inicio' 1366 728
    Save-StShotAt '01b-inicio-fullhd' 1920 1040
    Save-StShotAt '01c-inicio-chico' 1280 680
    $St.Product = Get-StProduct '2024'; $St.Mode = 'install'
    Show-StDataPage
    $St.UI.InKey.Text = 'Tu clave: abcde-12345-fghij-67890-klmno  Código: ST-AB12CD'
    $St.CodeTimer.Stop()
    Set-StCodeChip 'valid' 'Válido · 3 créditos'
    Save-StShotAt '02-datos'
    Show-StConfirmPage
    $St.UI.ChkConfirm.IsChecked = $true
    Save-StShotAt '03-confirmar'
    Initialize-StSteps @('check', 'restore', 'download', 'clean', 'remove', 'install', 'verify', 'act_key', 'act_iid', 'act_cid', 'act_confirm')
    foreach ($id in @('check', 'restore', 'download', 'clean')) { Set-StUiStep $id 'done' }
    Set-StUiStep 'remove' 'skip'
    Set-StUiStep 'install' 'running' (Get-StInstallHint -ProductId '2024')
    $St.UI.TxtWhile.Text = 'No apagues ni reinicies el computador. Verás una ventana de Microsoft con el avance: cuando diga «¡Hemos terminado!», presiona «Cerrar» o espera, la cerramos por ti.'
    $St.UI.TxtProgTitle.Text = 'Instalando Office 2024'
    $St.UI.TxtElapsed.Text = '07:42'
    Set-StProgStamp -Text 'PENDIENTE' -Kind 'pending'
    Show-StPage 'Progress'; Save-StShotAt '04-avance'
    $St.Product = Get-StProduct '365'
    $St.UI.TxtProgTitle.Text = 'Instalando Microsoft 365'
    Set-StProgStamp -Text 'PENDIENTE' -Kind 'pending'
    $St.UI.PanelTutorialPromo.Visibility = 'Visible'
    Save-StShotAt '05-avance-365'
    $St.Mode = 'tutorial'
    Show-StTutorial -Index 1; Save-StShotAt '06-tutorial-paso2'
    Save-StShotAt '06c-tutorial-fullhd' 1920 1040
    Set-StTutorialStep -Index 2; Save-StShotAt '06a-tutorial-contrasena'
    Set-StTutorialStep -Index 3; Save-StShotAt '06b-tutorial-pregunta'
    $St.Tut.Auth = $false
    Set-StTutorialStep -Index 8; Save-StShotAt '07-tutorial-paso9'
    Set-StTutorialStep -Index 6; Save-StShotAt '07b-tutorial-sin-imagen'
    $St.Product = Get-StProduct '2024'
    $St.UI.TxtDoneTitle.Text = '¡Listo! Office 2024 quedó activado'
    $St.UI.TxtDoneText.Text = 'Ya puedes usar Word, Excel y PowerPoint en este computador.'
    Set-StDoneList @('E8A5|Word, Excel, PowerPoint y más, **instalados**', 'E8D7|**Activado** con tu licencia', 'E895|Se actualiza solo con Microsoft', 'E721|Para abrirlos después: botón **Inicio** y escribe **Word**')
    $St.UI.TxtDoneNote.Text = 'Guarda el correo de tu compra. Si algún día reinstalas Windows, escríbenos y te ayudamos.'
    Show-StDonePage -Stamp 'ACTIVADO'; Save-StShotAt '08-listo'
    $St.ErrRetry = { }
    Show-StError -Title 'No pudimos instalar Office 2024' -Detail 'No pudimos conectarnos con los servidores de Microsoft. Revisa tu conexión a Internet.' -OnRetry { }
    Save-StShotAt '09-error'
    $St.Act = @{ Iid = '123456712345671234567123456712345671234567123456712345671234567' }
    $dlg = Show-StDialog -Kind 'success' -Title 'Office 2024 quedó instalado' -Message 'Ahora lo activaremos automáticamente con tu clave y tu código SafeTech. No necesitas abrir Word ni escribir números.' -Buttons @('Activar ahora')
    $dlg.Show(); Save-StShot $dlg '10-ventana-emergente'; $dlg.Close()
    $f = New-StDialogField -Label 'Código de activación SafeTech' -Placeholder 'ST-XXXXXX' -Value 'ST-AB12CO' -MaxLength 40 -Hint 'Revisa que empiece con ST- y que no confundas 0 (cero) con la letra O.'
    $dlg = Show-StDialog -Kind 'warn' -Title 'Tu código SafeTech no funcionó' -Message "El código no es válido o ya se usó.`nCódigo no válido. Verifica que esté escrito correctamente." -Extra $f.Panel -Buttons @('Cancelar', 'Probar este código') -Primary 'Probar este código' -LinkLabel 'Activar de otra forma' -Width 640
    $dlg.Show(); Save-StShot $dlg '11-ventana-codigo'; $dlg.Close()
    $dlg = Show-StAbout -Welcome
    $dlg.Show(); Save-StShot $dlg '12-bienvenida'; $dlg.Close()
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
$St.MaximizeFix = Enable-StMaximizeFix -Window $St.Window
$St.Window.WindowState = 'Maximized'
Update-StWindowFrame
[void]$St.Window.ShowDialog()
if ($env:SAFETECH_OFFICE_CHILD -eq '1') { [Environment]::Exit(0) }
