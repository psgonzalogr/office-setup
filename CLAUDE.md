# Instrucciones para Claude Code

## Commits y pull requests

- Autor de los commits: `SafeTech <contacto@safe-tech.cl>`. Antes del primer commit de la sesión: `git config user.name "SafeTech"` y `git config user.email "contacto@safe-tech.cl"`.
- No agregar líneas de firma a los mensajes de commit ni a las descripciones de pull requests: nada de `Co-Authored-By`, `Claude-Session`, «Generated with Claude Code» ni enlaces a sesiones.
- Mensajes en español: título corto (con la versión al inicio si cambia, por ejemplo «1.6.3: …») y, si hace falta, un párrafo que explique el porqué.

## Versiones

- Al publicar una versión: subir `Version` en `office.ps1`, anotarla al inicio de `CHANGELOG.md` y actualizar la sección «Versión» del `README.md`. La release se crea sola cuando el CI pasa en `main` (`.github/workflows/release.yml`).
- Al cambiar voces o imágenes, subir la carpeta (`assets/voz/v6` → `v7`) porque los equipos guardan los recursos en caché.
- Los enlaces dentro de `CHANGELOG.md` van con la dirección completa (`https://github.com/psgonzalogr/office-setup/blob/main/...`), porque también se copian a las notas de cada release.
