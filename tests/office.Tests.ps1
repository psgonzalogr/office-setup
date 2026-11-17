# Pruebas del motor del instalador (Pester 5, Windows PowerShell 5.1).
# Cargan office.ps1 en modo biblioteca: no abre ventanas ni pide permisos.

BeforeAll {
    $env:SAFETECH_OFFICE_MODE = 'lib'
    $text = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\office.ps1'), [Text.Encoding]::UTF8)
    . ([scriptblock]::Create($text))
    $ErrorActionPreference = 'Stop'
}

AfterAll {
    Remove-Item Env:\SAFETECH_OFFICE_MODE -ErrorAction SilentlyContinue
}

Describe 'Productos' {
    It 'Office 2024 usa ProPlus2024Volume y el canal PerpetualVL2024' {
        $p = Get-StProduct '2024'
        $p.OdtId | Should -Be 'ProPlus2024Volume'
        $p.Channel | Should -Be 'PerpetualVL2024'
        $p.NeedsKey | Should -BeTrue
    }
    It 'Office 2021 usa ProPlus2021Retail' {
        (Get-StProduct '2021').OdtId | Should -Be 'ProPlus2021Retail'
    }
    It 'Microsoft 365 no pide clave' {
        $p = Get-StProduct '365'
        $p.OdtId | Should -Be 'O365ProPlusRetail'
        $p.NeedsKey | Should -BeFalse
    }
    It 'Un producto desconocido falla' {
        { Get-StProduct '2019' } | Should -Throw
    }
}

Describe 'Clave de producto' {
    It 'Agrega guiones y mayúsculas' {
        Format-StProductKey 'abcde12345fghij67890klmno' | Should -BeExactly 'ABCDE-12345-FGHIJ-67890-KLMNO'
    }
    It 'Ignora espacios y guiones mal puestos y corta en 25 caracteres' {
        Format-StProductKey ' ab-cde 12345-FGHIJ67890KLMNOXYZ' | Should -BeExactly 'ABCDE-12345-FGHIJ-67890-KLMNO'
    }
    It 'Formatea claves incompletas' {
        Format-StProductKey 'abcde1' | Should -BeExactly 'ABCDE-1'
        Format-StProductKey '' | Should -BeExactly ''
    }
    It 'Valida solo claves completas' {
        Test-StProductKey 'ABCDE-12345-FGHIJ-67890-KLMNO' | Should -BeTrue
        Test-StProductKey 'ABCDE-12345-FGHIJ-67890-KLMN' | Should -BeFalse
        Test-StProductKey 'abcde-12345-fghij-67890-klmno' | Should -BeFalse
    }
    It 'Obtiene los últimos 5 caracteres' {
        Get-StKeyLast5 'ABCDE-12345-FGHIJ-67890-KLMNO' | Should -BeExactly 'KLMNO'
    }
}

Describe 'Código de activación SafeTech' {
    It 'Normaliza distintas formas de escribirlo' {
        ConvertTo-StAccessCode 'st-ab12cd' | Should -BeExactly 'ST-AB12CD'
        ConvertTo-StAccessCode 'ST AB12CD' | Should -BeExactly 'ST-AB12CD'
        ConvertTo-StAccessCode 'AB12CD' | Should -BeExactly 'ST-AB12CD'
        ConvertTo-StAccessCode 'ST-ST1234' | Should -BeExactly 'ST-ST1234'
    }
    It 'Rechaza códigos incompletos' {
        ConvertTo-StAccessCode 'ST-AB1' | Should -BeNullOrEmpty
        ConvertTo-StAccessCode '' | Should -BeNullOrEmpty
    }
    It 'Valida el formato' {
        Test-StAccessCode 'ST-AB12CD' | Should -BeTrue
        Test-StAccessCode 'ST-AB12C' | Should -BeFalse
    }
}

Describe 'Ocultar datos sensibles en los registros' {
    It 'Oculta la clave salvo los últimos 5 caracteres' {
        Hide-StSecret '/inpkey:ABCDE-12345-FGHIJ-67890-KLMNO' | Should -BeExactly '/inpkey:XXXXX-XXXXX-XXXXX-XXXXX-KLMNO'
    }
    It 'Oculta parte del código ST' {
        Hide-StSecret 'codigo ST-AB12CD usado' | Should -BeExactly 'codigo ST-AB**** usado'
    }
}

Describe 'Lectura de OSPP.VBS' {
    BeforeAll {
        $dstatus = @'
---Processing--------------------------
---------------------------------------
SKU ID: 11111111-2222-3333-4444-555555555555
LICENSE NAME: Office 24, Office24ProPlus2024VL_MAK_AE edition
LICENSE DESCRIPTION: Office 24, VOLUME_MAK channel
LICENSE STATUS:  ---NOTIFICATIONS---
ERROR CODE: 0xC004F056
Last 5 characters of installed product key: KLMNO
---------------------------------------
SKU ID: 66666666-7777-8888-9999-000000000000
LICENSE NAME: Office 24, Office24ProPlus2024VL_KMS_Client_AE edition
LICENSE DESCRIPTION: Office 24, VOLUME_KMSCLIENT channel
LICENSE STATUS:  ---LICENSED---
Last 5 characters of installed product key: 2YT43
---------------------------------------
---Exiting-----------------------------
'@
        $dinstid = @'
---Processing--------------------------
---------------------------------------
Installation ID for: Office 24, Office24ProPlus2024VL_MAK_AE edition: 123456712345671234567123456712345671234567123456712345671234567
---------------------------------------
---Exiting-----------------------------
'@
    }
    It 'Separa cada licencia con su estado y clave' {
        $l = @(ConvertFrom-StOsppStatus $dstatus)
        $l.Count | Should -Be 2
        $l[0].Name | Should -BeExactly 'Office 24, Office24ProPlus2024VL_MAK_AE edition'
        $l[0].Status | Should -BeExactly 'NOTIFICATIONS'
        $l[0].Last5 | Should -BeExactly 'KLMNO'
        $l[1].Status | Should -BeExactly 'LICENSED'
    }
    It 'Devuelve una lista con dos licencias sin anidarla' {
        @(ConvertFrom-StOsppStatus ($dstatus + "`n" + $dstatus)).Count | Should -Be 4
    }
    It 'Devuelve una lista vacía si no hay licencias' {
        @(ConvertFrom-StOsppStatus "---Processing---`n---Exiting---").Count | Should -Be 0
    }
    It 'Reconoce los estados activados' {
        Test-StLicensedStatus 'LICENSED' | Should -BeTrue
        Test-StLicensedStatus 'UNLICENSED' | Should -BeFalse
        Test-StLicensedStatus 'NOTIFICATIONS' | Should -BeFalse
        Test-StLicensedStatus 'OOB_GRACE' | Should -BeFalse
        Test-StLicensedStatus '' | Should -BeFalse
    }
    It 'Lee el ID de instalación y su licencia' {
        $i = @(ConvertFrom-StOsppIid $dinstid)
        $i.Count | Should -Be 1
        $i[0].Name | Should -BeExactly 'Office 24, Office24ProPlus2024VL_MAK_AE edition'
        $i[0].Iid.Length | Should -Be 63
    }
    It 'Encuentra un ID de instalación aunque cambie el formato' {
        $i = @(ConvertFrom-StOsppIid ('ID: ' + ('9' * 54)))
        $i[0].Iid | Should -BeExactly ('9' * 54)
    }
    It 'Extrae el código de error' {
        Get-StErrorCode 'ERROR CODE: 0xc004f069 algo' | Should -BeExactly '0xC004F069'
        Get-StErrorCode 'sin error' | Should -BeNullOrEmpty
    }
    It 'Reconoce el éxito de OSPP' {
        Test-StOsppSuccess '<Product activation successful>' | Should -BeTrue
        Test-StOsppSuccess 'ERROR CODE: 0xC004F04D' | Should -BeFalse
    }
    It 'Tiene mensajes en español para los errores conocidos' {
        Get-StOsppMessage '0xC004C060' | Should -Match 'bloqueada'
        Get-StOsppMessage '0x12345678' | Should -Match '0x12345678'
    }
}

Describe 'Elegir el ID de instalación correcto' {
    BeforeAll {
        $script:fakeStatus = @'
---------------------------------------
LICENSE NAME: Office 24, Office24ProPlus2024VL_KMS_Client_AE edition
LICENSE STATUS:  ---OOB_GRACE---
Last 5 characters of installed product key: 2YT43
---------------------------------------
LICENSE NAME: Office 24, Office24ProPlus2024VL_MAK_AE edition
LICENSE STATUS:  ---NOTIFICATIONS---
Last 5 characters of installed product key: KLMNO
---------------------------------------
'@
        $script:fakeIid = @'
---------------------------------------
Installation ID for: Office 24, Office24ProPlus2024VL_KMS_Client_AE edition: 111111111111111111111111111111111111111111111111111111111111111
---------------------------------------
Installation ID for: Office 24, Office24ProPlus2024VL_MAK_AE edition: 222222222222222222222222222222222222222222222222222222222222222
---------------------------------------
'@
    }
    It 'Devuelve solo el ID de la licencia de la clave del cliente' {
        Mock Invoke-StOspp { $script:fakeStatus } -ParameterFilter { $Arguments -contains '/dstatus' }
        Mock Invoke-StOspp { $script:fakeIid } -ParameterFilter { $Arguments -contains '/dinstid' }
        $iid = Get-StInstallationId -ProductId '2024' -Key 'ABCDE-12345-FGHIJ-67890-KLMNO'
        $iid | Should -BeOfType [string]
        $iid | Should -BeExactly ('2' * 63)
    }
    It 'Falla con un mensaje claro si Office no entrega el ID' {
        Mock Invoke-StOspp { $script:fakeStatus } -ParameterFilter { $Arguments -contains '/dstatus' }
        Mock Invoke-StOspp { '---Exiting---' } -ParameterFilter { $Arguments -contains '/dinstid' }
        { Get-StInstallationId -ProductId '2024' -Key 'ABCDE-12345-FGHIJ-67890-KLMNO' } | Should -Throw '*ID de instalaci*'
    }
}

Describe 'Configuración de ODT' {
    It 'Genera un XML válido para cada producto' {
        foreach ($id in @('2024', '2021', '365')) {
            $path = Join-Path $TestDrive "$id.xml"
            [void](New-StOdtXml -Path $path -ProductId $id -Edition '64')
            [xml]$x = Get-Content -LiteralPath $path -Raw
            $p = Get-StProduct $id
            $x.Configuration.Add.Product.ID | Should -BeExactly $p.OdtId
            $x.Configuration.Add.Channel | Should -BeExactly $p.Channel
            $x.Configuration.Add.OfficeClientEdition | Should -BeExactly '64'
            $x.Configuration.Add.Product.Language.ID | Should -BeExactly 'es-es'
            @($x.Configuration.Property | Where-Object { $_.Name -eq 'AUTOACTIVATE' }).Value | Should -BeExactly '0'
            ($null -ne $x.Configuration.SelectSingleNode('RemoveMSI')) | Should -BeTrue
        }
    }
    It 'Excluye OneDrive solo en Microsoft 365' {
        $path = Join-Path $TestDrive '365b.xml'
        [void](New-StOdtXml -Path $path -ProductId '365')
        [xml]$x = Get-Content -LiteralPath $path -Raw
        @($x.Configuration.Add.Product.ExcludeApp.ID) | Should -Contain 'OneDrive'
    }
    It 'Genera el XML de desinstalación silenciosa' {
        $path = Join-Path $TestDrive 'quitar.xml'
        [void](New-StOdtXml -Path $path -Remove)
        [xml]$x = Get-Content -LiteralPath $path -Raw
        $x.Configuration.Remove.All | Should -BeExactly 'TRUE'
        $x.Configuration.Display.Level | Should -BeExactly 'None'
    }
}

Describe 'Nombres amigables' {
    It 'Traduce los ProductReleaseIds' {
        Get-StFriendlyName 'O365ProPlusRetail' | Should -BeExactly 'Microsoft 365'
        Get-StFriendlyName 'O365HomePremRetail' | Should -BeExactly 'Microsoft 365 Familia o Personal'
        Get-StFriendlyName 'ProPlus2019Retail' | Should -BeExactly 'Office 2019'
        Get-StFriendlyName 'VisioPro2021Retail' | Should -BeExactly 'Visio 2021'
        Get-StFriendlyName 'ProPlusRetail' | Should -BeExactly 'Office (ProPlusRetail)'
    }
    It 'Agrupa dígitos para mostrarlos' {
        Format-StDigitGroups '123456789012' 6 | Should -BeExactly '123456 789012'
    }
}

Describe 'Detección en este equipo' {
    It 'Devuelve un estado aunque no haya Office' {
        $s = Get-StOfficeState
        $s.PSObject.Properties.Name | Should -Contain 'HasAny'
        $s.PSObject.Properties.Name | Should -Contain 'Has365'
    }
}

Describe 'Motor dentro de una tarea en segundo plano (como en la app)' {
    BeforeAll {
        $script:fakeOspp = @'
function Invoke-StOspp {
    param([string[]]$Arguments, [string]$OsppPath)
    if ($Arguments -contains '/dinstid') {
        return "Installation ID for: Office 24, Office24ProPlus2024VL_KMS_Client_AE edition: " + ('1' * 63) + "`nInstallation ID for: Office 24, Office24ProPlus2024VL_MAK_AE edition: " + ('2' * 63)
    }
    return "---------------------------------------`nLICENSE NAME: Office 24, Office24ProPlus2024VL_KMS_Client_AE edition`nLICENSE STATUS:  ---OOB_GRACE---`nLast 5 characters of installed product key: 2YT43`n---------------------------------------`nLICENSE NAME: Office 24, Office24ProPlus2024VL_MAK_AE edition`nLICENSE STATUS:  ---NOTIFICATIONS---`nLast 5 characters of installed product key: KLMNO`n---------------------------------------"
}
'@
        function Invoke-InRunspace([string]$Body) {
            $script = "param(`$sync, `$StConfig)`n" + $StEngine.ToString() + "`n" + $script:fakeOspp + "`n" + $Body
            $ps = [powershell]::Create()
            [void]$ps.AddScript($script).AddArgument($null).AddArgument($StConfig)
            $h = $ps.BeginInvoke()
            $out = $ps.EndInvoke($h)
            $ps.Dispose()
            return , @($out)
        }
    }
    It 'El motor no usa Select-Object -First (detiene la tarea en PowerShell 5.1)' {
        @($StEngine.ToString() -split "`n" | Where-Object { $_ -notmatch '^\s*#' -and $_ -match 'Select-Object\s+-First' }).Count | Should -Be 0
    }
    It 'Get-StMyLicense termina y devuelve la licencia de la clave' {
        $out = Invoke-InRunspace "`$l = Get-StMyLicense -Key 'ABCDE-12345-FGHIJ-67890-KLMNO'; 'nombre=' + `$l.Name; 'fin'"
        $out | Should -Contain 'fin'
        $out | Should -Contain 'nombre=Office 24, Office24ProPlus2024VL_MAK_AE edition'
    }
    It 'Get-StInstallationId termina y elige el ID correcto' {
        $out = Invoke-InRunspace "'iid=' + (Get-StInstallationId -ProductId '2024' -Key 'ABCDE-12345-FGHIJ-67890-KLMNO'); 'fin'"
        $out | Should -Contain 'fin'
        $out | Should -Contain ('iid=' + ('2' * 63))
    }
}
