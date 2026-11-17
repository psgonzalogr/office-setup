---
name: Instalador de Office SafeTech
description: El comprobante de compra SafeTech hecho objeto, en la noche SafeTech, para instalar y activar Office sin miedo.
colors:
  noche: "#09111F"
  pozo: "#0C1626"
  pizarra: "#111C2F"
  comprobante: "#17243A"
  pizarra-alta: "#20304A"
  linea: "#2A3C57"
  linea-fuerte: "#405776"
  borde-control: "#6B82A3"
  anillo-inactivo: "#8FA3BE"
  texto: "#F8FAFC"
  texto-suave: "#DBE5F1"
  texto-lead: "#C9D5E3"
  apagado: "#AEBDD0"
  rojo-marca: "#DC2626"
  rojo-profundo: "#991B1B"
  rojo-talon-fin: "#8E1818"
  rojo-texto: "#FCA5A5"
  talon-etiqueta: "#FECACA"
  verde-logrado: "#059669"
  verde-sello: "#34D399"
  verde-texto: "#6EE7B7"
  ambar: "#F59E0B"
  ambar-claro: "#FCD34D"
  ambar-error: "#B45309"
  azul-info: "#2563EB"
  azul-info-texto: "#93C5FD"
  whatsapp: "#25D366"
  whatsapp-tinta: "#052E16"
typography:
  display:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "34px"
    fontWeight: 600
  headline:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "28px"
    fontWeight: 600
  title:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "22px"
    fontWeight: 600
  lead:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "18px"
    fontWeight: 400
    lineHeight: "27px"
  body:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "17px"
    fontWeight: 400
    lineHeight: "26px"
  small:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "16px"
    fontWeight: 400
    lineHeight: "22px"
  label:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "17px"
    fontWeight: 600
  field:
    fontFamily: "Consolas, Segoe UI"
    fontSize: "21px"
    fontWeight: 400
  talon-numero:
    fontFamily: "Segoe UI"
    fontSize: "34px"
    fontWeight: 900
  talon-etiqueta:
    fontFamily: "Segoe UI"
    fontSize: "13px"
    fontWeight: 700
  sello:
    fontFamily: "Segoe UI"
    fontSize: "24px"
    fontWeight: 900
rounded:
  sm: "10px"
  md: "14px"
  lg: "16px"
  panel: "20px"
  comprobante: "22px"
  dialogo: "24px"
spacing:
  xs: "6px"
  sm: "8px"
  md: "12px"
  lg: "18px"
  xl: "24px"
  xxl: "36px"
  margen: "40px"
components:
  button-primary:
    backgroundColor: "{colors.rojo-marca}"
    textColor: "#FFFFFF"
    typography: "{typography.label}"
    rounded: "{rounded.md}"
    padding: "12px 28px"
    height: "56px"
  button-soft:
    backgroundColor: "{colors.comprobante}"
    textColor: "{colors.texto}"
    typography: "{typography.label}"
    rounded: "{rounded.md}"
    padding: "12px 24px"
    height: "56px"
  button-soft-hover:
    backgroundColor: "{colors.pizarra-alta}"
  button-whatsapp:
    backgroundColor: "{colors.whatsapp}"
    textColor: "{colors.whatsapp-tinta}"
    typography: "{typography.label}"
    rounded: "{rounded.md}"
    padding: "12px 24px"
    height: "56px"
  button-link:
    textColor: "{colors.rojo-texto}"
    rounded: "{rounded.sm}"
    padding: "8px 10px"
    height: "44px"
  field:
    backgroundColor: "{colors.pozo}"
    textColor: "{colors.texto}"
    typography: "{typography.field}"
    rounded: "{rounded.md}"
    padding: "0 18px"
    height: "58px"
  field-focus:
    backgroundColor: "{colors.pizarra}"
  check-row:
    backgroundColor: "{colors.comprobante}"
    textColor: "{colors.texto}"
    rounded: "{rounded.md}"
    padding: "14px 18px"
    height: "58px"
  ticket:
    backgroundColor: "{colors.comprobante}"
    textColor: "{colors.texto}"
    rounded: "{rounded.panel}"
    height: "256px"
  talon:
    backgroundColor: "{colors.rojo-marca}"
    textColor: "#FFFFFF"
    typography: "{typography.talon-numero}"
    width: "108px"
  panel:
    backgroundColor: "{colors.pizarra}"
    rounded: "{rounded.panel}"
    padding: "24px"
  dialog:
    backgroundColor: "{colors.pizarra}"
    textColor: "{colors.texto}"
    rounded: "{rounded.dialogo}"
    padding: "28px 30px 26px"
    width: "480px"
  sello-activado:
    textColor: "{colors.verde-sello}"
    typography: "{typography.sello}"
    rounded: "{rounded.md}"
    padding: "6px 16px"
  sello-pendiente:
    textColor: "{colors.ambar-claro}"
    rounded: "{rounded.sm}"
    padding: "4px 12px"
---

# Design System: Instalador de Office SafeTech

<!-- Registrado después de construir, a partir de office.ps1 (rama rediseno, 2026-11-17) y de las capturas de CI en Windows. Plataforma: WPF de escritorio, no web. Los valores son píxeles independientes de dispositivo (DIP) a escala 1; toda la vista crece o se encoge con la escala de la ventana. Claves de estilo WPF entre comillas invertidas. -->

## Overview

**Creative North Star: "El comprobante timbrado"**

La app es el comprobante de compra de SafeTech convertido en objeto. Sobre la noche SafeTech, cada pantalla importante es una pieza de pizarra con un talón rojo a la izquierda que muestra en grande lo que el cliente compró (2024, 2021, 365). Una línea punteada con muescas separa el talón del cuerpo. El cliente completa el comprobante con sus dos datos, tal como vienen en el correo, y al terminar cada etapa un sello rotado cae encima: ámbar «PENDIENTE» mientras se trabaja y verde «ACTIVADO» o «INSTALADO» al final.

El tono es tranquilo y grande, sin letra chica. El texto base mide 17 px, las acciones principales 56 px de alto y hay una sola acción roja evidente por decisión. SafeBOT acompaña desde arriba con una sola frase en un globo, y a su derecha tres etapas numeradas muestran dónde va el cliente. La ayuda humana (WhatsApp verde) vive siempre en el pie.

La metáfora solo enmarca. El comprobante aparece en la elección (tres boletos), en los datos, en la confirmación, en la tarjeta del tiempo transcurrido y en el «¡Listo!». Las pantallas de error, el tutorial y las ventanas emergentes no usan talón: son pizarra sobria con íconos en cuadros de color.

**Key Characteristics:**
- Solo modo oscuro: noche #09111F con capas de pizarra, sin superficies claras.
- Talón rojo con degradado vertical, perforación punteada y muescas: la firma de la app.
- Sellos rotados con borde grueso y letra Black en mayúsculas como único marcador de estado del comprobante.
- Rojo para actuar y para lo actual, verde para lo logrado, ámbar para avisos y pendientes.
- Letra grande (17 px base) y objetivos de clic de 56 px.
- Foco de teclado ámbar, grueso y separado del control.

## Colors

Una noche azul pizarra en capas, con un rojo de marca que reservamos para actuar y tres colores de estado que nunca se cruzan.

### Primary
- **Rojo SafeTech** (rojo-marca → rojo-profundo): degradado diagonal (`RedGrad`) de la acción principal, de la barra de avance, del punto de la etapa actual y del giro del paso en curso. Al pasar el ratón se aclara apenas (`RedGradHover`, #E23434 → #A11D1D).
- **Talón** (rojo-marca → rojo-talon-fin): degradado vertical (`StubGrad`) solo para el talón del comprobante.
- **Rojo legible** (rojo-texto): texto de enlaces (`BtnLink`), número del paso en curso y avisos de error sobre pizarra. El rojo de marca nunca se usa como color de texto sobre la noche.
- **Etiqueta del talón** (talon-etiqueta): la palabra «OFFICE» o «MICROSOFT» sobre el talón.

### Secondary
- **Verde logrado** (verde-logrado): marcas de etapa terminada, casillas marcadas, filas de pasos listos, resumen «N tareas listas».
- **Verde sello** (verde-sello): borde y texto del sello «ACTIVADO».
- **Verde texto** (verde-texto): íconos de visto bueno sobre fondo verde translúcido (beneficios del boleto, lista del «¡Listo!», chip del código válido).

### Tertiary
- **Ámbar aviso** (ambar): borde del sello «PENDIENTE» y de los paneles de aviso («Mientras tanto», notas importantes del tutorial).
- **Ámbar claro** (ambar-claro): texto e íconos de aviso. También es el color del foco de teclado (`Focus`).
- **Ámbar error** (ambar-error): relleno del círculo de un paso que falló. En los pasos y en la página de error, los fallos se dicen en ámbar, no en rojo, para no asustar.
- **Azul información** (azul-info, azul-info-texto): solo en el cuadro del ícono de las ventanas emergentes de tipo información o pregunta.
- **WhatsApp** (whatsapp, whatsapp-tinta): exclusivo del botón de ayuda humana.

### Neutral
- **Noche** (noche): fondo de la ventana y relleno de las muescas del comprobante.
- **Pozo** (pozo): barra superior, pie, campos de texto, marco de capturas y cajas de datos técnicos; es lo más hundido.
- **Pizarra** (pizarra): paneles (`Panel`), ventanas emergentes y campo con foco.
- **Pizarra del comprobante** (comprobante): cuerpo del comprobante y de los boletos, globo de SafeBOT, botón suave, filas de casilla, tarjetas de bienvenida.
- **Pizarra alta** (pizarra-alta): estados al pasar el ratón, pistas de avance, cuadros de ícono neutros, viñetas destacadas del tutorial.
- **Línea** (linea) y **línea fuerte** (linea-fuerte): bordes de superficies y divisores; la fuerte, para la perforación, la fila del paso en curso y bordes de ventana emergente.
- **Borde de control** (borde-control): borde de 2 px de botones suaves y campos en reposo; también el pulgar de la barra de desplazamiento.
- **Anillo inactivo** (anillo-inactivo): contorno de etapas y pasos pendientes, marcador de posición en campos y borde de la casilla vacía.
- **Texto** (texto), **texto suave** (texto-suave), **texto lead** (texto-lead), **apagado** (apagado): títulos; cuerpo; bajadas y descripciones; ayudas, notas y pasos pendientes. Todos superan 7:1 sobre la pizarra del comprobante.

### Named Rules
**The Tres Semáforos Rule.** Rojo es actuar o «aquí estás», verde es «ya está hecho» y ámbar es «atención» o «falta». Un color de estado nunca toma el papel de otro: el paso que falló y la página de error se pintan en ámbar. La única excepción es el cuadro de ícono de las ventanas emergentes de error y de clave, que usa rojo al 15 % con ícono rojo legible.

**The Tinte al 15 Rule.** Los fondos de estado son el mismo color semántico con transparencia: relleno al 15 % (alfa 26) en cuadros de ícono, al 10 % (alfa 1A) en paneles de aviso, y borde al 50 % (alfa 80). No se inventan tonos opacos intermedios.

## Typography

**Display Font:** Segoe UI Variable Display (con Segoe UI)
**Body Font:** Segoe UI Variable Text (con Segoe UI)
**Label/Mono Font:** Consolas, solo en los campos donde se escribe la clave y el código, y en el ID de instalación

**Character:** La familia del propio Windows, en semibold para todo lo que se lee primero, y monoespaciada grande para que la clave se vea igual que en el correo. Nada depende de fuentes instaladas aparte.

### Hierarchy
- **Display** (600, 34 px; `H1`): título de la elección «¿Qué compraste?». Las demás páginas usan el mismo estilo reducido a 28–32 px para caber en 1366×768.
- **Headline** (600, 28 px): títulos de comprobante, avance, error y tutorial.
- **Title** (600, 22 px; `H2`, 23 px en ventanas emergentes): título de cada ventana emergente. El nombre del producto en el boleto usa 26 px Display.
- **Lead** (400, 18 px / 27 px; `Lead`, color texto-lead): una sola bajada bajo el título. El globo de SafeBOT usa 18 px / 26 px en texto.
- **Body** (400, 17 px / 26 px; `Body`): todo el texto que explica, listas de confirmación y del tutorial.
- **Small** (400, 16 px / 22 px; `Small`, color apagado): ayudas bajo los campos, notas, contador de pasos. Es el mínimo para texto corrido.
- **Label** (600, 17 px; `Label`): rótulos de campo y texto de botones.
- **Campo** (400, 21 px, Consolas; `Field`): lo que el cliente escribe o pega.
- **Talón** (900, 32–34 px; etiqueta 700, 13 px en mayúsculas): el número del producto y su palabra sobre el talón.
- **Sello** (900, 24 px en «¡Listo!», 17 px en la tarjeta de avance, mayúsculas).
- **Reloj** (600, 36 px Display): el tiempo transcurrido en la tarjeta de avance.

### Named Rules
**The Diecisiete Rule.** El texto que se lee mide 17 px; 16 px es el piso para ayudas y notas. Por debajo de 16 px solo van chips de estado (14 px semibold), números dentro de círculos y la etiqueta del talón.

**The Clave Como en el Correo Rule.** Clave y código se escriben en Consolas de 21 px con marcador de posición del formato real (XXXXX-… y ST-XXXXXX); no se usa letra proporcional en esos campos.

## Layout

Ventana sin marco de Windows, abierta maximizada, diseñada en 1280×760 y escalada entera con una transformación de diseño: escala = mín(ancho/1280, alto/760), limitada entre 0,8 y 1,0: en pantallas chicas se achica, en las grandes no se agranda (a 1,37 en Full HD la letra se veía excesivamente grande; corregido en 1.1.1). El contenido queda centrado con su ancho máximo.

La ventana tiene cuatro franjas fijas: barra superior de 64 px (logo, «¿Cómo funciona?», botones de ventana), franja de guía (SafeBOT de 66 px de alto + globo de hasta 660 px + tres etapas), área de página y pie de 64 px (WhatsApp, Centro de Ayuda, versión). En el tutorial la franja de guía se oculta para ganar alto.

Las páginas centran su contenido con un ancho máximo: 1240 px para elección y avance, 900–980 px para los comprobantes de datos, confirmación y «¡Listo!», 840 px para error. El margen lateral es de 40 px. La elección son tres boletos iguales en fila (256 px de alto, 12 px entre ellos). El avance usa dos columnas: lista de pasos y una columna fija de 340 px con la tarjeta del reloj y el aviso. El tutorial usa un índice de 250 px, contenido a 1,2 partes y captura a 1 parte, con la navegación abajo tras un divisor.

El ritmo vertical se arma con 6, 8, 12, 18 y 24 px; el cuerpo del comprobante lleva 36 px de margen interior a los lados. La lista de pasos muestra solo un resumen de lo terminado, el paso actual y los próximos, con «y N pasos más», para no requerir desplazamiento.

### Named Rules
**The Sin Scroll en 1366 Rule.** Toda página debe caber sin desplazamiento en 1366×768 a la escala calculada; si no cabe, se acorta el texto o se resume la lista, no se agrega una barra.

## Elevation & Depth

La profundidad se logra con capas tonales (noche → pozo → pizarra → pizarra del comprobante → pizarra alta) y bordes de 1–2 px. Hay solo tres sombras, todas suaves y hacia abajo, y cada una marca algo que se puede tomar o que flota.

### Shadow Vocabulary
- **Brillo de la acción** (`DropShadowEffect` color #7F1D1D, desenfoque 22, distancia 6, opacidad 0,45): bajo el botón principal rojo.
- **Boleto levantado** (negro, desenfoque 30, distancia 10, opacidad 0,35): bajo cada boleto de la elección; al pasar el ratón sube 4 px y su borde se vuelve rojo.
- **Ventana emergente** (negro, desenfoque 40, distancia 10, opacidad 0,55): bajo la tarjeta de la ventana emergente.

### Named Rules
**The Plano Salvo Para Tomar Rule.** Paneles, comprobantes de página y tarjetas informativas no llevan sombra; solo la llevan el botón principal, los boletos que se eligen y las ventanas emergentes.

## Shapes

Esquinas generosas y constantes, sin ángulos rectos en nada que el cliente toque. Botones, campos, filas y notas usan 14 px; cuadros de ícono, enlaces y el sello pendiente 10 px; tarjetas de bienvenida y marco de capturas 16 px; paneles y boletos 20 px; comprobantes de página 22 px; ventana emergente 24 px. Etapas, pasos y beneficios son círculos.

El comprobante tiene una silueta propia: talón rojo con las esquinas izquierdas redondeadas y las derechas rectas, línea vertical punteada de 2 px (#405776, trazo 2,5 / 2,5) a lo largo del borde del talón, y dos muescas circulares de 22–26 px del color de la noche, centradas en la perforación y cortando los bordes superior e inferior.

### Named Rules
**The Muesca Verdadera Rule.** La muesca se pinta siempre con el color exacto del fondo sobre el que está el comprobante (noche #09111F); si el comprobante se coloca sobre otra superficie, la muesca cambia con ella.

## Components

### Buttons
Grandes, tranquilos y con un foco imposible de perder.
- **Shape:** 14 px de esquina; 56 px de alto mínimo (48 px en la barra superior y el pie, 50 px en el panel secundario de inicio).
- **Primary (`BtnPrimary`):** degradado rojo SafeTech, texto blanco semibold 17 px, relleno 12 × 28 px, brillo rojo debajo. Ícono Fluent opcional a la izquierda o flecha a la derecha.
- **Hover / Pressed:** el degradado se aclara; al presionar, escala a 0,97.
- **Focus:** anillo ámbar (#FCD34D) de 3 px, separado 5 px del botón y con 17 px de esquina. Igual en botones suaves, de WhatsApp, boletos (anillo de 25 px de esquina separado 6 px) y casillas.
- **Disabled:** opacidad 0,38; «Continuar» y «Comenzar instalación» parten así hasta que los datos o la casilla estén listos.
- **Soft (`BtnSoft`):** pizarra del comprobante con borde de 2 px (borde-control); al pasar el ratón, borde apagado y fondo pizarra alta. Para «Atrás», «Pegar», «Cerrar el instalador» y opciones no principales.
- **WhatsApp (`BtnWa`):** verde WhatsApp con tinta verde oscura; solo para pedir ayuda humana.
- **Link (`BtnLink`):** texto rojo legible semibold 16 px, sin fondo; al pasar el ratón un velo rojo al 15 %.
- **Ventana (`BtnChrome`, `BtnClose`):** 48 × 40 px, íconos Fluent de 12 px; cerrar se pinta #C42B1C al pasar el ratón, como en Windows.

### Ticket (boleto de elección)
La firma de la pantalla de inicio.
- **Forma:** botón completo (`Ticket`) de 256 px de alto, pizarra del comprobante, borde de 2 px (linea), esquina 20 px, sombra de boleto levantado.
- **Talón:** 108 px de ancho con la palabra y el número del producto; perforación y muescas de 24 px.
- **Cuerpo:** nombre en Display 26 px, «Word, Excel, PowerPoint y más» en small, dos beneficios con visto bueno verde en círculo de 24 px, y al pie una franja roja de 54 px con «Instalar …» y flecha. Toda la tarjeta es un solo objetivo de clic.
- **Hover:** borde rojo y sube 4 px.

### Comprobante de página
Datos, confirmación y «¡Listo!» son un comprobante grande: pizarra del comprobante, borde de 2 px, esquina 22 px, talón de 108–118 px que cambia con el producto elegido, muescas de 26 px y cuerpo con 32–36 px de margen. La tarjeta del reloj en el avance es la versión compacta (talón de 88 px, muescas de 22 px).

### Sello
- **Pendiente:** borde ámbar de 3 px, texto ámbar claro Black 17 px, esquina 10 px, rotado −6°. Vive en la tarjeta del reloj.
- **Activado / Instalado:** borde verde sello de 3 px, texto verde sello Black 24 px, esquina 12 px, rotado −9°, arriba a la derecha del «¡Listo!».
- **Movimiento:** el sello cae: escala de 1,7 a 1 en 300 ms (avance) o de 1,6 a 1 en 320 ms tras 260 ms de espera con aparición en 180 ms («¡Listo!»), con rebote suave (BackEase, amplitud 0,35, salida).

### Inputs / Fields
- **Style (`Field`):** pozo, borde de 2 px (borde-control), esquina 14 px, 58 px de alto, Consolas 21 px, marcador de posición en anillo-inactivo con el formato real.
- **Hover:** borde apagado. **Focus:** borde ámbar y fondo pizarra.
- **Validación:** chip dentro del campo, a la derecha (28 px de alto, esquina 14 px, borde verde de 1,5 px, texto verde 14 px semibold) cuando el código es válido. Al lado de cada campo, un botón suave «Pegar» de la misma altura.

### Check row (`Check`)
Una fila completa clicable de 58 px: pizarra del comprobante, borde de 2 px (línea fuerte), esquina 14 px, cuadro de 28 px con esquina 8 px. Marcada: cuadro y borde de fila en verde logrado con visto bueno blanco.

### Cards / Containers
- **Panel (`Panel`):** pizarra, borde de 1 px (linea), esquina 20 px, relleno 24 px. Variante de aviso: fondo ámbar al 10 % y borde ámbar al 50 %.
- **Nota del tutorial:** esquina 14 px, borde de 1,5 px; tres grados (información en pizarra, «Importante:» ámbar al 10 %, «Muy importante:» ámbar al 15 % con borde de 2,5 px). La nota crítica va antes de los botones.
- **Filas con ícono:** cuadro de 34 px con esquina 10 px en pizarra alta (confirmación, detalles técnicos) o círculo de 34 px verde al 15 % (lista del «¡Listo!»), ícono Fluent de 15–16 px y texto de 17 px.
- **Tarjetas de bienvenida:** 2 × 2, pizarra del comprobante, borde de 1,5 px, esquina 16 px, círculo de color de 52 px con ícono de 22 px, título 19 px semibold y texto 17 px.

### Navigation
- **Etapas (franja de guía):** tres círculos de 36 px unidos por líneas de 44 × 2 px. Hecho: relleno verde con visto bueno. Actual: degradado rojo con número blanco y rótulo semibold en texto. Pendiente: contorno anillo-inactivo y rótulo apagado.
- **Pasos del avance:** filas de esquina 14 px con círculo de 32 px. En curso: fondo pizarra del comprobante, borde línea fuerte, arco rojo de 3,5 px que gira cada 0,9 s y subtítulo visible. Hecho: círculo verde. Error: círculo ámbar error y título ámbar claro.
- **Índice del tutorial:** filas de 42 px con círculo numerado de 28 px y rótulo de 16 px recortado; barra de avance de 10 px abajo.
- **Avance:** pista de 14 px (pizarra alta, esquina 7 px) con relleno en degradado rojo que se anima en 450 ms con salida quíntica.

### SafeBOT y su globo
SafeBOT plano de 66 px a la izquierda, con un globo de pizarra del comprobante (borde 1,5 px, esquina 18 px, relleno 20 × 12 px, ancho máximo 660 px) y su pico triangular. Una sola frase por pantalla en 18 px.

### Ventana emergente
Tarjeta de pizarra, borde de 1,5 px (línea fuerte), esquina 24 px, sombra de ventana emergente y una línea roja de 2 px que se desvanece a los lados en el borde superior (`TopLine`). Cuadro de ícono de 52 px con esquina 16 px y tinte al 15 % según el tipo (información azul, éxito verde, aviso ámbar, error y clave rojo). Botones alineados a la derecha, mínimo 140 px de ancho; el primero es el principal.

### Movimiento
Las páginas entran con aparición y subida de 18 px en 360 ms con salida quíntica. Todo movimiento se desactiva en modo sin animación.

## Do's and Don'ts

### Do:
- **Do** poner el talón rojo con perforación y muescas en cada comprobante, mostrando el número del producto elegido.
- **Do** usar `BtnPrimary` para la acción que hace avanzar, con 56 px de alto, y `BtnSoft` para «Atrás» y alternativas.
- **Do** marcar el estado del comprobante con un sello rotado: ámbar «PENDIENTE» mientras se trabaja, verde «ACTIVADO» o «INSTALADO» al terminar.
- **Do** dar a todo control el anillo de foco ámbar de 3 px separado del borde.
- **Do** escribir texto corrido a 17 px y ayudas a 16 px como mínimo, en los colores texto, texto suave, texto lead o apagado.
- **Do** pintar avisos y errores en ámbar con el tinte al 10–15 % y el borde al 50 %.
- **Do** mantener el botón verde de WhatsApp visible en el pie en todas las páginas.
- **Do** usar solo íconos de Segoe Fluent Icons (con MDL2 de respaldo) dentro de cuadros o círculos de color.

### Don't:
- **Don't** usar el rojo de marca para adornos, para el paso que falló ni como color de texto; el rojo es actuar o «aquí estás», y como texto se usa rojo legible (#FCA5A5).
- **Don't** poner sombra en paneles, comprobantes de página o tarjetas informativas.
- **Don't** usar fondos claros ni un modo claro.
- **Don't** usar fuentes que no vengan con Windows.
- **Don't** poner rótulos en mayúsculas pequeñas sobre los títulos (kicker); el título y una bajada bastan. Las únicas mayúsculas son la palabra del talón y el texto de los sellos.
- **Don't** usar el verde de WhatsApp para nada que no sea pedir ayuda humana.
