@echo off
setlocal
rem ==========================================================================
rem  SafeTech - Instalador de Office
rem  Doble clic para abrir el instalador. Descarga la version mas reciente
rem  desde https://github.com/psgonzalogr/office-setup y la ejecuta.
rem ==========================================================================
title SafeTech - Instalador de Office
color 0F
echo.
echo   ============================================================
echo     SafeTech - Instalador de Office
echo   ============================================================
echo.
echo   Abriendo el instalador. Espera unos segundos...
echo   Windows te va a preguntar si permites cambios: responde SI.
echo.
set "SAFETECH_OFFICE_LAUNCHER=cmd"
set "ST_URL=https://raw.githubusercontent.com/psgonzalogr/office-setup/main/office.ps1"
if defined SAFETECH_OFFICE_URL set "ST_URL=%SAFETECH_OFFICE_URL%"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference = 'Stop'; try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $r = Invoke-WebRequest -UseBasicParsing -Uri $env:ST_URL; $t = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()).TrimStart([char]0xFEFF); Invoke-Expression $t } catch { Write-Host ''; Write-Host ('  No se pudo abrir el instalador: ' + $_.Exception.Message) -ForegroundColor Yellow; exit 1 }"
if errorlevel 1 goto fallo
exit /b 0

:fallo
echo.
echo   Si Windows te pidio permiso, vuelve a intentarlo y responde SI.
echo   Revisa tambien tu conexion a Internet.
echo   Si necesitas ayuda, escribenos por WhatsApp al +56 9 5101 7168.
echo.
pause
exit /b 1
