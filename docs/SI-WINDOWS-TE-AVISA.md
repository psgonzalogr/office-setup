# ¿Windows o tu antivirus muestran un aviso con el instalador de SafeTech?

**En corto:** el aviso significa *«no conozco a este autor»*, no *«encontré un virus»*. Aquí te explicamos por qué aparece y cómo puedes comprobar tú mismo que el instalador es el original.

## Los avisos que puedes ver

| Aviso | Qué significa | Qué hacer |
|---|---|---|
| **«Windows protegió su PC»** (pantalla azul al abrir `Instalar-Office-SafeTech.cmd`) | Windows SmartScreen no conoce al autor del archivo porque no tiene una firma digital de pago. Sale con cualquier programa nuevo sin esa firma. | Pulsa **Más información** y luego **Ejecutar de todas formas**. |
| **«¿Quieres permitir que esta aplicación haga cambios en el dispositivo?»** | Es el permiso de administrador. Instalar o desinstalar Office y registrar su licencia siempre lo pide, también cuando lo hace un técnico a mano. | Responde **Sí**. Es normal. |
| **Tu antivirus bloquea o pone en cuarentena el instalador** | Algunos antivirus desconfían de programas que descargan archivos, desinstalan programas y registran licencias, porque los programas maliciosos también lo hacen. Se llama «falso positivo». | **No desactives tu antivirus.** Escríbenos por WhatsApp al +56 9 5101 7168 y lo revisamos contigo. |

## Por qué no tiene firma digital

Las empresas grandes pagan cada año un certificado de «firma de código» para que Windows las reconozca. SafeTech es una empresa pequeña y este instalador todavía no tiene ese certificado. Por eso Windows muestra el aviso, igual que con cualquier programa nuevo sin firma, aunque no tenga nada malo.

## Cómo comprobarlo tú mismo

1. **El código es público.** Todo lo que hace el instalador está en este repositorio, en [`office.ps1`](../office.ps1), con comentarios en español.
2. **Compara la huella del archivo.** Cada [release](https://github.com/psgonzalogr/office-setup/releases/latest) publica `SHA256.txt`. En PowerShell, dentro de la carpeta donde descargaste el archivo:
   ```powershell
   Get-FileHash .\Instalar-Office-SafeTech.cmd
   ```
   Si el número coincide con el de `SHA256.txt`, el archivo es exactamente el publicado.
3. **Revísalo con más de 60 antivirus a la vez** en [VirusTotal](https://www.virustotal.com): sube el archivo y mira el resultado.

## Qué hace y qué no hace

- **Descarga Office directo desde Microsoft** (`officecdn.microsoft.com`) y comprueba la firma digital de Microsoft antes de usarlo.
- **Activa con los métodos oficiales de Microsoft**: tu clave de producto y, si hace falta, la activación por teléfono, con el ID de confirmación que obtiene con tu código SafeTech (`getcid.safe-tech.cl`).
- **No usa cracks, activadores KMS ni programas que modifiquen Windows u Office.**
- **No queda nada funcionando en segundo plano** al terminar: solo un registro del proceso en `C:\ProgramData\SafeTech`.
- **No toca tus documentos, fotos ni correos.**

### A qué sitios se conecta

| Sitio | Para qué |
|---|---|
| `raw.githubusercontent.com` | Descargar el instalador y sus imágenes, voces y sonidos desde este repositorio |
| `officecdn.microsoft.com` y servidores de Microsoft | Descargar la herramienta oficial de Microsoft y Office |
| `getcid.safe-tech.cl` | Validar tu código SafeTech y obtener el ID de confirmación (Office 2021 y 2024) |
| `portal.office.com`, `youtube.com`, `safe-tech.cl`, `wa.me` | Solo si pulsas el botón del portal, del video, de los términos o de WhatsApp |

## Para SafeTech: cómo reducir los avisos

- Enviar cada versión nueva del `.cmd` a Microsoft para análisis: <https://www.microsoft.com/wdsi/filesubmission> (opción *Software developer*). Si no encuentran nada, suelen quitar la detección de Microsoft Defender.
- Si otro antivirus da un falso positivo, la mayoría tiene un formulario para reportarlo.
- A futuro, firmar el `.cmd` con un certificado de firma de código.
