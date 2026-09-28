# mindsoftdoc

Genera documentos Word con el formato de la casa Mindsoft a partir de un
archivo markdown. El formato no se reinventa: sale del XML literal de
`plantilla.docx`, así que fuentes, colores, encabezado, pie, viñetas y tablas
salen idénticos siempre.

    python3 mindsoftdoc.py --nuevo                   # menú: arma el .md de arranque
    python3 mindsoftdoc.py documento.md              # deja documento.docx al lado
    python3 mindsoftdoc.py documento.md -o final.docx

Solo necesita Python 3. No hay dependencias que instalar.

## Reglas de la casa

Antes de que cualquier texto llegue al documento pasa por un saneador que
reemplaza:

| Entra | Sale |
|---|---|
| guiones largos y medios | `-` |
| comillas tipográficas y angulares | `"` y `'` |
| puntos suspensivos de un carácter | `...` |
| espacios duros | espacio normal |
| `**negrita**` en cualquier texto | texto normal (la negrita solo vive en títulos y cabeceras de tabla) |

Al terminar informa qué corrigió, por ejemplo `Reglas de la casa: 28 guiones
largos`. Nunca cambia nada en silencio, y aplica sin importar quién escriba el
markdown.

Los acentos y la ñ se escriben siempre: va **"Actuación"**, no "Actuacion". El
saneador no los toca, y la tabla de arriba es la lista completa de lo prohibido:
no hay ninguna regla contra la tilde.

## Empezar un documento: `--nuevo`

El menú pregunta dos cosas y escribe el `.md` de arranque:

    $ python3 mindsoftdoc.py --nuevo

    1. Tipo de documento
      1) Propuesta con costos
      2) Propuesta sin costos
      3) Documento de análisis
      4) Guía / informativo
      5) Cotización
    > 1

    2. Bloques reutilizables
      1) Forma de pago 50/50 por fase           [x]
      2) Forma de pago 50/50 anticipo-final
      3) SEM / Google Ads
      4) Hosting dedicado (Debian)
      5) Hosting dedicado (AlmaLinux + cPanel)
      6) Manejo de hosting externo
      7) Plan de soporte mensual
      8) Portafolio de clientes                 [x]
    > 1 3 8

    3. Datos
       titulo: Lanzamiento + Agente AI
       subtitulo:
       encabezado:
       cliente: InmoWeb

    Escrito: inmoweb.md

Los bloques que se ofrecen dependen del tipo elegido, y las variables que se
preguntan salen de los propios bloques. Con `--nuevo propuesta-con-costos` se
salta la primera pregunta.

El menú está acá a propósito y no en la generación: **generar un `.md` tiene que
dar siempre el mismo `.docx`**. El `.md` es la fuente de verdad, se versiona y se
regenera cuando haga falta.

## Bloques reutilizables

Una línea sola, en el lugar del documento donde va el bloque:

    @incluir bloques/portafolio-clientes.md

Se busca junto al `.md` y, si no está, junto al generador; así `bloques/...`
funciona desde cualquier carpeta. Un bloque que no existe corta con un error, no
se ignora.

Dentro de un bloque, `{{variable}}` se reemplaza con el campo del front-matter
que tenga ese nombre:

    Se propone posicionar estratégicamente a {{cliente}} en Ecuador.

Lo que no tenga valor se deja a la vista en el documento y se avisa al terminar
(`AVISO - variables sin valor: cliente`). Nunca se borra en silencio.

| Bloque | Va en |
|---|---|
| `bloques/portafolio-clientes.md` | Propuestas y cotizaciones |
| `bloques/forma-de-pago-50-50.md` | Costo por fases: 50% al inicio y 50% al fin de cada fase |
| `bloques/forma-de-pago-anticipo-final.md` | Costo global: 50% a la firma y 50% contra entrega |
| `bloques/sem-google-ads.md` | Propuestas con campaña |
| `bloques/hosting-dedicado.md` | Infraestructura sobre Debian, sin panel |
| `bloques/hosting-dedicado-cpanel.md` | Infraestructura sobre AlmaLinux con WHM/cPanel |
| `bloques/hosting-externo.md` | Propuestas con infraestructura |
| `bloques/soporte-mensual.md` | Propuestas con soporte |

Las specs y los precios de los bloques de hosting y soporte salen de propuestas
reales: **revisalos en cada propuesta antes de mandarla**.

Hay dos pares de bloques que se excluyen entre sí: se elige uno, no los dos.

- Forma de pago: `forma-de-pago-50-50.md` es la versión por fases (InmoWeb, BYD);
  `forma-de-pago-anticipo-final.md` es la global (Siegfried PMC, Eliana).
- Hosting: `hosting-dedicado.md` es la oferta sobre Debian, sin panel (BYD);
  `hosting-dedicado-cpanel.md` es la de AlmaLinux con WHM/cPanel (Inmoweb). Este
  último pide `hosting_vcpu`, `hosting_ram`, `hosting_disco`, `hosting_mensual` y
  `hosting_anual` en el front-matter.

Los `<!-- comentarios -->` de los esqueletos no llegan al documento.

## El archivo de entrada

```
---
rotulo: Reporte
titulo: Análisis de Seguridad - Sistema de Cafetería
subtitulo: Incidente de reembolso · cotopaxi.k12.ec · Septiembre 2026
encabezado: Análisis de Seguridad - Cafetería AC | Septiembre 2026
indice: si
---

## Primera sección
...
```

| Campo | Para qué sirve |
|---|---|
| `rotulo` | La palabra sobre el título en la portada: Reporte, Manual, Especificación Funcional, Propuesta. Por defecto `Reporte` |
| `titulo` | Título grande de la portada. Si falta, se toma el primer `#` del documento |
| `subtitulo` | Línea gris bajo el título |
| `encabezado` | Línea superior de cada página. Si falta, se usa el título |
| `indice` | `si` inserta un índice automático después de la portada |

El pie con la dirección y los teléfonos de Mindsoft es fijo y no se configura.

## Qué reconoce del markdown

| Escribís | Sale como |
|---|---|
| `## Título` | Título de sección azul con la barra debajo, numerado 1, 2, 3 |
| `### Subtítulo` | Negrita azul de 11 puntos, numerado 1.1, 1.2 |
| texto suelto | Párrafo Verdana 10, justificado |
| `- item` | Viñeta |
| `1. item` | Lista numerada, cada lista arranca de nuevo en 1 |
| tabla con pipes | Cabecera azul y filas en blanco, texto a la izquierda |
| `*cursiva*` `` `código` `` | Formato en línea |
| `**negrita**` | Nada: sale como texto normal (regla de la casa) |
| `[texto](url)` | Texto, con la url entre paréntesis si es http |
| `![pie](captura.png)` | Imagen centrada con pie de figura |
| `> cita` | Párrafo normal: se le quita el `>` |
| `---` en una línea | Salto de página |
| bloque entre ` ``` ` | Recuadro gris en Consolas, literal: conserva la sangría y no interpreta marcado |

Las columnas de las tablas se reparten solas según cuánto texto lleva cada una.
Las imágenes se escalan al ancho útil de la página sin deformarse; PNG y JPEG.

Dentro de un bloque de código nada es marcado: un `#`, un `|`, un `-` o un `**`
salen tal cual, y los tabuladores pasan a cuatro espacios. Un ` ``` ` sin cerrar
corta con un error, porque si no se comería el resto del documento. Las reglas
de la casa sí aplican adentro (un JSON lleva comillas rectas igual).

La numeración la pone Word, no el markdown: en el `.md` los títulos van sin
número y Word los renumera solo cuando se mueve o se agrega una sección. Los
`###` entran al índice con su número (`2.1`), debajo de su sección.

## Verificación

    python3 test_mindsoftdoc.py

Comprueba las reglas de la casa, el marcado en línea, el reparto de columnas,
la medición de imágenes, la lectura del markdown, los `@incluir` y las
variables (incluidos el bloque que falta y el ciclo), que los comentarios no
lleguen al documento, que los 5 tipos generen un `.docx` válido con todos sus
bloques, y que el documento generado sea XML válido, con los estilos y las
listas resueltos contra la plantilla y las 21 partes del formato intactas.

Se validó además contra dos documentos reales ya aprobados:

- El reporte de la cafetería de Academia Cotopaxi: regenerado desde markdown da
  el mismo texto carácter por carácter una vez aplicadas las reglas de la casa,
  y la misma estructura (13 títulos, 14 tablas). Las 145 negritas del original
  hoy saldrían como texto normal: la regla de la casa es posterior a esa prueba.
- La propuesta de Siegfried PMC, ya con bloques: 45 líneas de markdown más un
  `@incluir` del portafolio.

## Limitaciones conocidas

- Sin listas anidadas ni notas al pie.
- `--nuevo` necesita una terminal interactiva. Para armar un `.md` sin menú, se
  copia a mano el esqueleto de `esqueletos/`.
- `{{variables}}` funciona en el cuerpo del documento, no dentro del
  front-matter.
- El índice lo arma Word al abrir el archivo. En LibreOffice u ONLYOFFICE hay
  que refrescarlo a mano con F9.
- Guardar el documento desde ONLYOFFICE reescribe los estilos. Si hay que
  editar a mano, conviene hacerlo en Word, o mejor editar el markdown y
  regenerar.
- Guardar `plantilla.docx` desde un editor renumera los ids internos de estilos
  y listas. El generador los resuelve por nombre, así que no se rompe, pero ese
  guardado se lleva las tipografías Nunito embebidas: si el documento tiene que
  verse igual en una máquina sin Nunito instalada, hay que volver a embeberlas.

## Archivos

| Archivo | Qué es |
|---|---|
| `mindsoftdoc.py` | El generador |
| `plantilla.docx` | El formato de la casa. No editar salvo para cambiar el formato de todos los documentos |
| `bloques/` | Bloques reutilizables. Texto, sin código |
| `esqueletos/` | Estructura de cada tipo de documento |
| `ejemplo.md` / `ejemplo.docx` | Muestra de cada elemento soportado |
| `test_mindsoftdoc.py` | Verificación |
