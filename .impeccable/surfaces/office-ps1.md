---
version: 1
slug: "office-ps1"
primary_target: "office.ps1"
related_targets: []
---

# Surface brief: instalador (office.ps1)

Scope: toda la app de escritorio (inicio, datos, confirmación, avance, tutorial 365, listo, error, ventanas emergentes). Mode: operate.
Audience/job: clientes de SafeTech con brecha digital, 40+, que instalan y activan Office una sola vez con el correo de compra abierto.
Constraints: WPF + Windows PowerShell 5.1, un archivo, solo fuentes de Windows, solo modo oscuro SafeTech, ventana maximizada al abrir, aviso "¿Cómo funciona?" al abrir y sin scroll en 1366x768.
Seed: 48a30e3f (surface, operate). Dealt 6 (lead), 1, 2. Unattended run: the owner asked to proceed while asleep; lead locked, ideas grafted from the other two dealt cards.

## Direction contract

THESIS: La app es el comprobante de compra SafeTech hecho objeto: el cliente elige el comprobante de lo que compró, completa sus dos datos tal como vienen en el correo y el comprobante se va timbrando (Instalado, Activado). Rechaza el asistente genérico con kicker, tarjetas iguales y letra chica.

OWN-WORLD: Noche SafeTech (#09111F) y superficies pizarra; el comprobante es una pieza #17243A con talón rojo de marca a la izquierda (número grande 2024/2021/365), perforación punteada con muescas, campos con letra de 20 px y sellos rotados (verde Activado, ámbar Pendiente). SafeBOT plano como guía con una frase por pantalla. Rojo solo para la acción principal y el paso actual; verde lo logrado; ámbar los avisos.

STORY: "¿Qué compraste?" → comprobante con tus dos datos → qué va a pasar → avance con sellos → comprobante timbrado y Word abierto. La ayuda humana (WhatsApp) siempre visible.

FIRST VIEWPORT: Ventana maximizada; sobre el inicio, la bienvenida de SafeBOT con 4 mensajes cortos (original de Microsoft, nada pirata, activación con tu código, tus archivos no se tocan) y un solo botón "Entendido, empezar".

SIGNATURE: el sello que cae sobre el comprobante al terminar cada etapa (escala 1.25 → 1 con leve rotación, 280 ms).

RISK: que la metáfora tape la tarea; por eso el comprobante solo enmarca, cada pantalla tiene una sola acción principal de 56 px y textos de 17 px como base.
