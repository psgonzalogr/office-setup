@echo off
rem SafeTech - Instalador de Office
rem Doble clic para abrir el instalador. Descarga la version mas reciente desde GitHub.
title SafeTech - Instalador de Office
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; irm https://raw.githubusercontent.com/psgonzalogr/office-setup/main/office.ps1 | iex"
