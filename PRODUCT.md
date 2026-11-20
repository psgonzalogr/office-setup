# Product

<!-- impeccable:product-schema 1 -->

<!-- Escrito sin entrevista: el dueño pidió avanzar de forma autónoma (2026-11-17). Lo marcado como (inferido) no fue confirmado por él. -->

## Platform

Windows desktop (WPF, Windows PowerShell 5.1). No es web: se abre con `irm … | iex` o con `Instalar-Office-SafeTech.cmd` y corre en Windows 10/11 con permisos de administrador.

## Users

Clientes de SafeTech (Antofagasta, Chile) que compraron una licencia de Office 2021, Office 2024 o una cuenta de Microsoft 365 y la instalan solos en su computador. El dueño los describe como usuarios básicos de Windows o con brecha digital, muchos sobre 40 años. Lo hacen una sola vez, en su casa o trabajo, con el correo de la compra abierto al lado (inferido: a menudo en notebooks de 1366×768 y con poca confianza en instaladores).

## Product Purpose

Que el cliente pase de "compré una licencia" a "Office instalado y activado" sin escribir comandos ni copiar números. El instalador descarga Office desde Microsoft (ODT), quita versiones anteriores, activa Office 2021/2024 por teléfono de forma automática con el código ST de SafeTech, y guía la cuenta de Microsoft 365 paso a paso. Éxito: Office activado y el cliente sin necesidad de escribir a soporte.

## Positioning

La activación telefónica automática con el código SafeTech (getcid.safe-tech.cl): el cliente no copia el ID de instalación ni el de confirmación. Todo lo instalado es original de Microsoft; el código es público.

## Operating Context

- El cliente tiene a mano el correo de compra con la clave (XXXXX-…) y el código `ST-XXXXXX`, o la cuenta de Microsoft 365 que compró y su contraseña temporal.
- Durante la instalación aparecen ventanas de Microsoft; el proceso dura de 10 a 40 minutos.
- En Microsoft 365 el cliente usa su celular (Microsoft Authenticator, solo en algunas cuentas).
- Soporte: WhatsApp +56 9 5101 7168, contacto@safe-tech.cl, ayuda.safe-tech.cl.

## Capabilities and Constraints

- Productos: Office 2024 (LTSC ProPlus, volumen), Office 2021 (ProPlus Retail), Microsoft 365 Apps. Español, 64 bits.
- Activación 2021/2024: clave + código ST; respaldo manual con el portal. Cada activación usa 1 crédito.
- Microsoft 365: tutorial de 11 pasos con voz e imágenes ampliables; «Ya configuré mi cuenta» para un segundo computador.
- Desinstalación profunda de Office que no toca documentos, correos ni la configuración personal.
- Un solo archivo `office.ps1`; recursos gráficos desde el repositorio público.
- Debe funcionar sin fuentes instaladas aparte de las de Windows (Segoe UI / Segoe UI Variable, Segoe Fluent Icons o MDL2).
- La ventana se abre maximizada (pedido del dueño, 2026-11-17).

## Brand Commitments

- SafeTech: modo oscuro de la guía `web/GUIA_DISENO_SAFETECH.md` del repo SafeTech. Fondo #09111F, superficies #111C2F / #17243A / #20304A, rojo de marca #DC2626 → #991B1B, textos #F8FAFC / #DBE5F1 / #AEBDD0. Solo modo oscuro en esta app (pedido del dueño).
- Logo `assets/logo-safetech-dark.png`.
- Español de Chile, claro y sin tecnicismos. SafeTech no es distribuidor oficial de Microsoft.

## Evidence on Hand

Ilustraciones de cada pantalla del tutorial en `assets/tutorial/v2` (rehechas con una cuenta de ejemplo, `tucuenta@tudominio.com`). No hay testimonios ni cifras de clientes: no inventarlos.

## Product Principles

1. Una sola decisión por pantalla, con un botón principal evidente.
2. Tranquilizar antes que instruir: explicar qué va a pasar y que es seguro, sin jerga.
3. Nada se lee con scroll si puede caber: textos cortos y grandes.
4. El cliente nunca queda sin salida: siempre hay "Atrás", "Reintentar" o ayuda humana visible.
5. Honestidad: lo que la app dice que hace es exactamente lo que hace.

## Accessibility & Inclusion

Público con brecha digital y sobre 40 años (pedido del dueño): texto base de al menos 16 px, objetivos táctiles/clic de al menos 48 px, alto contraste (≥ 4.5:1 en texto), lenguaje simple, sin depender solo del color, foco de teclado visible.
