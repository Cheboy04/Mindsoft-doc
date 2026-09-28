# Handoff - mindsoftdoc

Última actualización: 2026-09-07, rama `main`. Documento para retomar el
proyecto en una sesión futura sin volver a leerlo todo. Si algo acá no coincide
con el código, gana el código: corre `python3 test_mindsoftdoc.py` primero.

## Qué es esto

Un generador de documentos Word con el formato de la casa Mindsoft a partir de
markdown. Nació para automatizar las propuestas comerciales, pero el mismo motor
sirve para cualquier documento de la empresa: reportes, manuales, guías de
usuario, guías de carga de contenido, especificaciones funcionales. Lo único que
cambia entre uno y otro es el campo `rotulo` de la portada y el contenido.

    python3 mindsoftdoc.py --nuevo                   # menú: arma el .md de arranque
    python3 mindsoftdoc.py documento.md              # deja documento.docx al lado
    python3 mindsoftdoc.py documento.md -o final.docx
    python3 test_mindsoftdoc.py                      # verificación

Python 3 puro, sin dependencias. Nada que instalar, nada que compilar.

## La decisión de diseño que sostiene todo

**El formato no se reinventa: se copia el XML literal de `plantilla.docx`.**

`plantilla.docx` es un reporte real ya aprobado (el de Siegfried Ecuador). El
generador lo abre como ZIP, se queda con `word/document.xml`, recorta la portada
y la barra de sección, y reconstruye un documento nuevo pegando ese XML con
bloques generados. Las otras 21 partes del paquete (estilos, tema, pie de
página, media) se copian byte a byte sin tocar.

Consecuencia práctica: fuentes, colores, viñetas, bordes de tabla y pie de página
salen idénticos siempre, porque nunca se escriben, se heredan. Un test verifica
que esas 21 partes queden intactas respecto de la plantilla.

Consecuencia igual de práctica: **cambiar el formato de todos los documentos =
editar `plantilla.docx` en Word**, no tocar Python. Y si se edita la plantilla,
hay que revisar los índices duros del recorte (ver "Trampas").

## La segunda decisión de diseño: el menú no vive en la generación

`--nuevo` abre un menú (tipo de documento, bloques reutilizables, datos) y
escribe un `.md`. Generar ese `.md` no pregunta nada.

Esto es deliberado y conviene no revertirlo: si los bloques se eligieran al
generar, el mismo `.md` daría documentos distintos según lo que se clickeó, se
perdería la regeneración (que es el flujo principal), los bloques solo podrían
ir al final, habría que pedir las variables en cada corrida, y los tests no
podrían llamar a `generar()` sin colgarse.

Con el menú en `--nuevo` el `.md` queda como fuente de verdad única:
versionable, diffeable, revisable por una persona. Y es el artefacto que la
futura automatización de redacción va a producir o pre-llenar.

## Arquitectura de `mindsoftdoc.py` (un solo archivo)

Ocho secciones, en orden en el archivo:

| Sección | Qué hace |
|---|---|
| Reglas de la casa | `limpiar()` normaliza caracteres prohibidos y lleva la cuenta |
| Bloques reutilizables | `incluir()` resuelve `@incluir` recursivo con tope de 10; `sustituir()` cambia `{{var}}` por el campo del front-matter |
| Plantilla | `_bloques()` parsea `<w:p>`/`<w:tbl>` de primer nivel; `Plantilla` recorta portada, barra, cabeza y cola, y resuelve los ids de la plantilla con `leer_estilos()` y `leer_numeracion()` |
| Piezas de XML | `runs()` para marcado en línea, y un `p_*()` por cada tipo de bloque |
| Lectura del markdown | `leer_markdown()` devuelve `(meta, bloques)`, cada bloque un `(tipo, dato)` |
| Imágenes | `medir()` lee dimensiones de PNG/JPEG a mano; `escalar()` ajusta al ancho útil |
| Armado | `generar()` recorre bloques y arma el XML; `espacios()` repone los `xmlns` que falten; `_escribir()` rearma el ZIP |
| `--nuevo` | `TIPOS` y `BLOQUES` (listas literales), `armar_nuevo()` puro y testeable, `menu_nuevo()` que es solo entrada/salida |

El flujo completo es lineal y cabe en una frase: markdown -> `leer_markdown` ->
lista de bloques `(tipo, dato)` -> un `p_*()` por bloque -> concatenar con
`tpl.cabeza` y `tpl.cola` -> `_escribir()` rearma el ZIP cambiando seis partes.

### Los tipos de bloque

`h2`, `sub`, `p`, `ul`, `ol`, `tabla`, `codigo`, `salto`, `img`. Agregar un elemento nuevo
al markdown son tres pasos: reconocerlo en `leer_markdown`, escribir su `p_*()`
copiando el XML de un documento real, y despacharlo en el `for` de `generar()`.

### Las seis partes del ZIP que se reescriben

Todo lo demás se copia tal cual (`_escribir`, al final del archivo):

- `word/document.xml` - el documento nuevo
- `word/header1.xml` - el primer `<w:t>` recibe el `encabezado`; el segundo se vacía
- `word/_rels/document.xml.rels` - relaciones de las imágenes agregadas (`rId1001+`)
- `word/settings.xml` - se inyecta `<w:updateFields>` si el documento lleva índice
- `word/numbering.xml` - se inyecta la lista multinivel de los títulos y un `w:num` por lista numerada
- `[Content_Types].xml` - se agregan los MIME de jpeg/jpg/gif si faltan

## Reglas de la casa

Antes de que cualquier texto llegue al documento pasa por `limpiar()`: guiones
largos y medios a `-`, comillas tipográficas y angulares a `"` y `'`, puntos
suspensivos de un carácter a `...`, espacios duros a espacio normal. Y en
`runs()`: la `**negrita**` sale como texto normal en párrafos, viñetas y celdas
(se cuenta y se reporta igual); sigue viva solo en títulos, subtítulos y la
cabecera azul de las tablas, que son estilo, no marcado.

No es cosmética: es un estándar de la empresa, y aplica sin importar quién
escriba el markdown (persona o modelo). Al terminar, el CLI informa qué corrigió
(`Reglas de la casa: 28 guiones largos`). Nunca cambia nada en silencio.

Este mismo criterio debería gobernar cualquier función nueva: **corregir siempre,
avisar siempre, nunca preguntar.**

## Campos del front-matter

| Campo | Para qué sirve |
|---|---|
| `rotulo` | Palabra sobre el título en la portada. Por defecto `Reporte` |
| `titulo` | Título grande. Si falta, se toma el primer `#` del documento |
| `subtitulo` | Línea gris bajo el título |
| `encabezado` | Línea superior de cada página. Si falta, se usa el título |
| `indice` | `si` inserta un índice automático después de la portada |

El pie con la dirección y los teléfonos de Mindsoft está fijo en
`word/footer1.xml` de la plantilla y no se configura desde el markdown.

## Trampas conocidas (lo que rompe si se toca sin mirar)

1. **Índices duros en `Plantilla.__init__`**. La portada son los primeros 23
   bloques del cuerpo, el 23 es el primer título del reporte original y se
   descarta, la barra de sección es el bloque 24, y el rótulo/título/subtítulo
   viven en los índices 19/20/21. Hay tres `assert` que
   avisan si la plantilla cambia, pero si se edita `plantilla.docx` y se agrega
   o quita un párrafo antes de la portada, estos números se corren.

2. **`poner()` reemplaza el primer `<w:t>` del bloque** (dentro de `generar()`). Si en
   Word se parte el texto de la portada en varios runs (pasa al editar y
   corregir ortografía), solo se reemplaza el primero y queda texto viejo
   pegado. Síntoma: la portada dice "ReporteAnalisis de Seguridad - Siegfried".

3. **ONLYOFFICE reescribe estilos al guardar.** Por eso se quitó el sombreado
   zebra de las tablas (commit `e0851cb`): no sobrevivía. Si hay que editar a
   mano, hacerlo en Word; mejor, editar el markdown y regenerar.

   Guardar `plantilla.docx` desde un editor también **renumera los ids**: los
   `styleId` (`Heading2` pasó a `788`, `Normal` a `786`) y las listas de
   `numbering.xml` (la viñeta pasó de `numId 90` a `5`). Un `pStyle` o un
   `numId` que no resuelve no da error: el título sale sin negrita, sin color y
   sin tamaño, y la viñeta sale sin punto. Por eso el generador ya no los cita
   por id: `leer_estilos()` los busca por `w:name` y `leer_numeracion()` por su
   definición (bullet / decimal). Del mismo guardado salió que faltara el
   `xmlns:pic` en `<w:document>` (`espacios()` lo repone) y que se perdieran las
   Nunito embebidas, que siguen citadas en `styles.xml` pero ya no viajan en el
   archivo.

4. **El índice lo arma Word al abrir.** Se inserta un campo TOC más
   `updateFields`. En LibreOffice u ONLYOFFICE hay que refrescar con F9.

5. **La numeración de títulos la pone Word, no el markdown.** `p_titulo` y
   `p_subtitulo` cuelgan de la lista multinivel `numId 92`, que `_escribir`
   inyecta en `word/numbering.xml` al vuelo (no vive en `plantilla.docx`: un
   guardado desde Word la borraría sin que nadie se entere). En el `.md` los
   títulos van sin número; Word renumera solo al mover una sección. Los `###`
   ahora sí entran al índice (`TOC \o "1-3"`, con `outlineLvl` explícito en
   cada título, que es lo que hace que LibreOffice y ONLYOFFICE también puedan
   armarlo). El título del propio índice va sin numerar y sin estilo `Heading2`
   (lleva el aspecto copiado a mano en `COMO_TITULO`): con el estilo puesto,
   Word mete el índice como primera entrada de sí mismo.

   Cada lista numerada del markdown cuelga de su propio `w:num` (100, 101, ...,
   los da `id_lista()`) sobre el abstractNum decimal de la plantilla, con
   `startOverride`. Sin eso Word las encadena y la segunda lista del documento
   arranca donde terminó la primera. Los numIds del documento y los inyectados
   en `numbering.xml` tienen que coincidir: si no, la lista sale sin números y
   en silencio, por eso el id sale de una sola función y `ejemplo.md` lleva dos
   listas numeradas seguidas.

   Los ids que se inyectan ya no son fijos: `leer_numeracion()` los corre por
   encima del máximo que trae la plantilla (`NUM_TITULOS`, `ABS_TITULOS` y
   `BASE_LISTAS`, que es de donde sale `id_lista()`). Con la plantilla de hoy
   dan 7, 6 y 8. Así un guardado que renumere `numbering.xml` no puede chocar
   contra ellos, que es lo que antes se perdia en silencio.

6. **Los bloques de hosting y soporte llevan precios y specs reales.** Salen de
   propuestas de 2026. Hay que revisarlos antes de mandar cada propuesta; el
   generador no sabe si están vigentes.

7. **Hay dos pares de bloques que se excluyen entre sí.** El menú los ofrece
   como opciones sueltas y no impide marcar los dos; si se marcan, el documento
   sale con dos formas de pago o dos ofertas de hosting. Se elige uno:

   - Hosting: `hosting-dedicado.md` (Debian 13 + CrowdSec + nftables, sin panel,
     la de BYD) contra `hosting-dedicado-cpanel.md` (AlmaLinux + WHM/cPanel +
     Imunify360, la de Inmoweb). No son versiones de lo mismo, son productos
     distintos.
   - Forma de pago: `forma-de-pago-50-50.md` (por fase: InmoWeb, BYD) contra
     `forma-de-pago-anticipo-final.md` (global: Siegfried PMC, Eliana).

8. **`hosting-dedicado-cpanel.md` se escribió sin la propuesta de Inmoweb a la
   vista** (`ejemplos/` no está en el disco). El stack sale del historial; el
   dimensionamiento y los precios quedaron como variables de front-matter
   (`hosting_vcpu`, `hosting_ram`, `hosting_disco`, `hosting_mensual`,
   `hosting_anual`) para que no viaje ningún número inventado. **Contrastar las
   filas de servicio contra la propuesta original antes de mandarlo.**

## Limitaciones actuales

- Sin listas anidadas ni notas al pie.
- Sin tabla de figuras, sin numeración automática de figuras (el pie de figura
  es texto libre que escribe el autor).
- El menú de `--nuevo` necesita una terminal interactiva. No hay forma no
  interactiva de armar un `.md`; para eso se copia un esqueleto a mano.
- `{{variables}}` funciona en el cuerpo, no dentro del front-matter.

## Verificación

`python3 test_mindsoftdoc.py` corre diez pruebas: reglas de la casa, marcado en
línea (incluido el caso de texto que parece marcado sin serlo, tipo
`[DEPOSIT][CONTROLLER]` o `SQLSTATE[42S22]`), reparto de columnas, medición de
imágenes, lectura del markdown, bloques de código (literales, con sangría y el
` ``` ` sin cerrar), `@incluir` y variables (con el bloque que no
existe y el ciclo), comentarios que no llegan al documento, los 5 tipos
generando `.docx` válido con todos sus bloques, y documento generado (XML
válido, cero caracteres prohibidos, conteos de títulos/tablas/listas/saltos,
estilos y listas resueltos contra la plantilla, imagen embebida, y las 21 partes
del formato intactas).

Al 2026-09-25 pasan las diez.

Los bloques de código (2026-09-25, para la entrega de InmoWeb al CRM) son una
tabla de una celda con fondo `FONDO_CODIGO`: el fondo de celda sobrevive a
ONLYOFFICE; el borde de párrafo no está probado. Una línea del `.md` es una
línea del `.docx` con `<w:br/>`, no un párrafo por línea, para que no quede
espacio entre ellas.

### El test no ve cómo se ve el documento

Esto es lo que más duele y no está en ningún assert: un `pStyle` que no resuelve,
un `numId` que no existe o un TOC vacío dan XML **válido**. El .docx se abre sin
quejarse y el error solo se ve mirando la página. Los tres bugs de formato del
2026-09-07 (títulos sin estilo, viñetas sin punto, índice vacío) pasaron los
nueve tests.

Para mirar de verdad, sin salir de la terminal:

    soffice --headless --convert-to pdf ejemplo.docx
    pdftoppm -r 75 -png -f 3 -l 3 ejemplo.pdf pag      # y abrir pag-3.png

El índice es aparte: es un campo, y ni LibreOffice ni el PDF lo llenan solos.
Hay que refrescarlo con la API de LibreOffice (`python3 -c "import uno"` ya viene
con el paquete) antes de exportar:

    soffice --headless --norestore --accept="socket,host=localhost,port=2002;urp;" &
    # conectar por uno, doc.getDocumentIndexes().getByIndex(0).update(), y
    # storeToURL(...writer_pdf_Export)

Si el índice sale vacío, el problema es el `outlineLvl` de los títulos, no el
campo TOC.

`ejemplo.md` es la muestra de todos los elementos **y** el fixture del test:
cualquier cambio de formato se verifica ahi, se regenera `ejemplo.docx` y se
commitea. Arreglar `ejemplo.docx` a mano no sirve de nada, se sobrescribe.

Se validó además contra dos documentos reales ya aprobados:

- **Reporte de la cafetería de Academia Cotopaxi.** Regenerado desde markdown da
  el mismo texto carácter por carácter una vez aplicadas las reglas de la casa, y
  la misma estructura (13 títulos, 14 tablas). Las 145 negritas del original hoy
  saldrían como texto normal: la regla de la casa es posterior a esa prueba.
- **Propuesta de Siegfried PMC** (2026-09-04, con el sistema de bloques ya
  puesto). Reconstruida en 45 líneas de markdown más un `@incluir` del
  portafolio. Vivía en `prueba/`, que ya no está en el disco.
- **Propuesta de actualización de la cafetería** (2026-09-07), en
  `ac_cafeteria_system/docs/`. Partida en `-cliente.md` y `-anexo-interno.md`
  porque el generador no corta un `.md` en dos, y con los números de título
  sacados a mano para que no chocaran con los de Word.

## Archivos

| Archivo | Qué es |
|---|---|
| `mindsoftdoc.py` | El generador. Un solo archivo, sin dependencias |
| `plantilla.docx` | El formato de la casa. Editar solo para cambiar el formato de todos los documentos |
| `bloques/` | Bloques reutilizables, extraídos de las propuestas reales. Texto, sin código |
| `esqueletos/` | Estructura de cada tipo de documento. Los `<!-- -->` guían a quien escribe y no llegan al `.docx` |
| `ejemplos/` | Propuestas reales de referencia. **Sin versionar: llevan precios de clientes.** Hoy no está en el disco; el estimation table quedó en `EstimationTable/` |
| `prueba/` | Reconstrucciones para comparar contra los originales. Sin versionar, descartable. Hoy no está en el disco |
| `ejemplo.md` / `ejemplo.docx` | Muestra de cada elemento soportado. Sirve de fixture del test |
| `ejemplo-captura.png` | Imagen del ejemplo, medida por el test (600x180) |
| `test_mindsoftdoc.py` | Verificación |
| `README.md` | Manual de uso para quien escribe markdown |
| `HANDOFF.md` | Este documento, para quien retoma el proyecto |

## Contexto del negocio

El cuello de botella no es el formato, ya está resuelto. El cuello de botella es
**escribir el contenido del markdown**. Hay un archivo de propuestas históricas
ya hechas que el dueño del proyecto puede aportar como material de referencia
para lo que venga.

## Qué sigue

1. **Estandarizar la tabla de estimación que ve el cliente.** Decisión del dueño
   del proyecto, en conversación con su equipo. Hoy los 4 ejemplos con costo usan
   4 formatos distintos: `Detalle/Tiempo/Complejidad`, `#/Descripcion/Pagina/Horas/Costo`,
   `Detalle/Descripcion/Tiempo/Costo` y `Fase/Tiempo/Inversion`. Los esqueletos
   ya proponen `Fase | Detalle | Tiempo | Inversión`, que es la de BYD.

   Al reconstruir Siegfried PMC apareció que **probablemente hagan falta dos
   formatos, no uno**: los proyectos grandes (BYD, InmoWeb) estiman por fase y
   llevan costo por fila; los trabajos chicos (Siegfried PMC, Eliana) estiman por
   tarea y solo llevan costo en el total. Forzar los dos a una sola tabla
   deforma uno de los dos.
2. **`estimacion.py`**: leer `Estimation Table Projects 2026.xlsx` y escribir
   `bloques/estimacion-<cliente>.md`. Script aparte a propósito, para no meter un
   lector de Excel dentro del generador. Bloqueado por el punto 1.
3. **Automatizar la redacción del `.md`.** El dueño del proyecto lo quiere pero no
   todavia. `--nuevo` ya deja el artefacto donde esa automatización va a escribir.

Sobre el estimation table, lo que ya se sabe de mirar las 5 hojas:

- El layout es estable: fila 7 es la cabecera (`Category | Task/Description |
  Best Case | Worst Case | Most likely | Cost | Requerido | MVP`), las tarifas
  por rol van en la fila 2, y `TOTAL VALUE` suma la columna MVP.
- En BYD el total del xlsx ($7.700) coincide exacto con el de la propuesta. El
  costo se puede automatizar.
- El tiempo **no** coincide: 3,5 semanas internas contra 16 semanas al cliente.
  El calendario es una decisión comercial, no un cálculo. No automatizarlo.
- Las fases se marcan en la columna A (`Fase de Analisis`) pero los sprints en la
  columna B (`Sprint 1`). Hay que fijar una convención antes de parsear.
