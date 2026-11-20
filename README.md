# Instalador de Office · SafeTech

Aplicación para que los clientes de SafeTech instalen **Office 2024**, **Office 2021** o **Microsoft 365** con un solo comando, sin tocar la consola después. Se abre maximizada, con el diseño oscuro de SafeTech, y SafeBOT guía cada paso. Está pensada para personas con poca experiencia en computadores: letra grande, una sola acción principal por pantalla y ayuda por WhatsApp siempre visible.

## Cómo se usa (cliente)

1. Clic derecho en el botón Inicio → **Terminal** o **Windows PowerShell**.
2. Pegar este comando y presionar Enter:

```powershell
irm https://raw.githubusercontent.com/psgonzalogr/office-setup/main/office.ps1 | iex
```

3. Windows pregunta si se permiten cambios: responder **Sí**. El instalador se abre en una ventana nueva y la consola ya se puede cerrar.

Alternativa sin consola: descargar [`Instalar-Office-SafeTech.cmd`](https://github.com/psgonzalogr/office-setup/releases/latest/download/Instalar-Office-SafeTech.cmd) y abrirlo con doble clic. Muestra lo que está haciendo y, si algo falla (por ejemplo, sin Internet), deja el mensaje en pantalla.

Al abrirse aparece la bienvenida «Así funciona este instalador» (4 mensajes cortos: Office original, nada pirata, se activa con tu código, tus archivos no se tocan).

## Qué hace

| Producto | Instalación (ODT) | Activación |
|---|---|---|
| Office 2024 | `ProPlus2024Volume`, canal `PerpetualVL2024`, 64 bits, español | Automática con la clave y el código `ST-XXXXXX` |
| Office 2021 | `ProPlus2021Retail`, canal `Current`, 64 bits, español | Automática con la clave y el código `ST-XXXXXX` |
| Microsoft 365 | `O365ProPlusRetail`, canal `Current`, 64 bits, español | Tutorial integrado paso a paso con la cuenta entregada |

Los productos, canales y exclusiones (Lync; en 365 también Groove y OneDrive) se tomaron de los asistentes Inno Setup anteriores (2024 v1.2.9, 2021 v1.0.0, 365 v1.0.5). En Windows de 32 bits instala la edición de 32 bits.

### Instalación

1. Revisa Windows 10/11, al menos 10 GB libres y conexión con Microsoft.
2. Crea un punto de restauración (si Windows lo permite; si no, sigue igual).
3. Descarga la Herramienta de implementación de Office desde `officecdn.microsoft.com` y **verifica su firma digital de Microsoft**.
4. Cierra Word, Excel, etc., retira claves anteriores y servidores KMS (`ospp.vbs /unpkey`, `/remhst`).
5. Si hay otro Office Click-to-Run, lo desinstala en silencio (`Remove All`) y lo verifica; las versiones MSI antiguas se quitan con `RemoveMSI`.
6. Instala el producto con la ventana de progreso de Microsoft y verifica que quedó instalado.

Microsoft 365: si ya está instalado (cualquier edición), no se toca nada y se pasa directo al tutorial. Mientras se instala, el cliente puede avanzar en la parte del tutorial que no necesita Word (portal, contraseña, Authenticator).

### Activación asistida (2021/2024)

Todo ocurre con ventanas emergentes; el cliente no escribe comandos ni copia números:

1. Registra la clave (`/inpkey`). Si la clave no corresponde al tipo de licencia instalada (error `0xC004F069`), instala las licencias de esa familia desde `Licenses16` y reintenta.
2. Prueba la activación en línea (`/act`). Si Microsoft informa que la clave está bloqueada (`0xC004C060`, `0xC004C003`, `0xC004C004`), se detiene sin gastar créditos.
3. Obtiene el ID de instalación (`/dinstid`) de la licencia de esa clave.
4. Pide el ID de confirmación a `getcid.safe-tech.cl/generate` con el código `ST-XXXXXX` del cliente (el mismo portal que ya usan).
5. Confirma con `/actcid` y verifica en `/dstatus` que la licencia quedó activada.

Si el código falla, se puede ingresar otro código, reintentar o **activar manualmente**: la ventana muestra el ID de instalación con botón para copiarlo, abre el portal y recibe el ID de confirmación.

El código se valida mientras se escribe con `getcid.safe-tech.cl/check-code` y muestra los créditos disponibles.

### Desinstalación profunda

Botón «Desinstalar Office por completo» en la pantalla de inicio (franja «¿Ya tienes Office?»). Pide confirmación con lo que se borra y lo que no, y luego:

1. Crea un punto de restauración.
2. Cierra Word, Excel, etc.
3. Retira claves y activaciones (`ospp.vbs /unpkey`, `/remhst`) y las cuentas de Office guardadas en el Administrador de credenciales (`MicrosoftOffice1x*`).
4. Desinstala todas las versiones: Click-to-Run con la Herramienta de implementación (`Remove All`; si falla, con el desinstalador propio de cada producto), MSI 2007–2016 con `msiexec /x` y Office de escritorio de la Microsoft Store.
5. Borra los restos: servicio `ClickToRunSvc`, tareas `\Microsoft\Office\`, accesos directos que apuntan a Office, carpetas (`Microsoft Office`, `microsoft shared\ClickToRun`, `OFFICE15/16`, `ProgramData\Microsoft\ClickToRun` y `\Office`), registro de máquina (`Office\ClickToRun`, `Office\15.0/16.0`, `AppVISV`, `OfficeSoftwareProtectionPlatform`) y, por usuario, solo licencias, identidad y registro de Office.
6. Comprueba que no quedó nada. Si algo estaba en uso, pide reiniciar.

**No toca nunca:** documentos, correos (`.pst`/`.ost`), la configuración personal de Office (plantillas, firmas, diccionario, autocorrección, perfiles de Outlook), OneDrive, Teams ni OneNote de la Tienda.

## Seguridad

- **No contiene secretos.** No usa el token de getcid.dev; solo llama a los endpoints públicos de `getcid.safe-tech.cl` con el código del cliente.
- Los registros nunca guardan la clave completa (solo los últimos 5 caracteres) ni el código completo.
- Office se descarga siempre desde Microsoft y se comprueba la firma del instalador.
- Los navegadores y Word se abren **sin** permisos de administrador.

## Registros y soporte

- Registros: `C:\ProgramData\SafeTech\Office\registros\<fecha>\proceso.log`
- Si algo falla, la pantalla de error tiene **Crear ZIP para soporte**: deja en el Escritorio un ZIP con el registro, información del equipo, estado de licencias y los registros de Microsoft (ODT).

## Desarrollo

Todo está en un solo archivo, [`office.ps1`](office.ps1) (Windows PowerShell 5.1 + WPF):

- **Arranque**: si no es administrador o no es la ventana propia, se relanza elevado (`-Verb RunAs`), oculto y en STA, descargando el script de nuevo en UTF-8.
- **Motor** (`$StEngine`): detección, ODT, OSPP, API. Se carga en la interfaz y en cada tarea en segundo plano (runspaces), así la ventana nunca se congela.
- **Interfaz**: XAML con los colores del modo oscuro de la guía de diseño de SafeTech (`#09111F`, superficies `#111C2F`/`#17243A`, rojo `#DC2626 → #991B1B`, textos `#F8FAFC`/`#AEBDD0`).
- **Imágenes** en [`assets/`](assets/): logo y capturas del tutorial (tomadas del Centro de Ayuda, con los correos difuminados).

Variables de entorno útiles:

| Variable | Uso |
|---|---|
| `SAFETECH_OFFICE_URL` | Probar otra rama: URL raw de `office.ps1` |
| `SAFETECH_OFFICE_MODE=lib` | Solo carga funciones (pruebas) |
| `SAFETECH_OFFICE_MODE=selftest` | Dibuja cada pantalla en PNG sin instalar nada |

GitHub Actions ([`validar.yml`](.github/workflows/validar.yml)) corre en Windows: sintaxis en PowerShell 5.1, PSScriptAnalyzer, pruebas Pester ([`tests/`](tests/)) y la autoprueba de la interfaz, que sube las capturas como artefacto.

## Versión

Versión actual: **1.6.2**. El historial completo está en [`CHANGELOG.md`](CHANGELOG.md).

---

SafeTech SpA · Antofagasta, Chile · [safe-tech.cl](https://safe-tech.cl) · No somos distribuidores oficiales de Microsoft.
