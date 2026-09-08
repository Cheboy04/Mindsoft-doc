# Handoff - mindsoftdoc

Ultima actualizacion: 2026-09-04, rama `main`. Documento para retomar el
proyecto en una sesion futura sin volver a leerlo todo. Si algo aca no coincide
con el codigo, gana el codigo: corre `python3 test_mindsoftdoc.py` primero.

## Que es esto

Un generador de documentos Word con el formato de la casa Mindsoft a partir de
markdown. Nacio para automatizar las propuestas comerciales, pero el mismo motor
sirve para cualquier documento de la empresa: reportes, manuales, guias de
usuario, guias de carga de contenido, especificaciones funcionales. Lo unico que
cambia entre uno y otro es el campo `rotulo` de la portada y el contenido.

    python3 mindsoftdoc.py --nuevo                   # menu: arma el .md de arranque
    python3 mindsoftdoc.py documento.md              # deja documento.docx al lado
    python3 mindsoftdoc.py documento.md -o final.docx
    python3 test_mindsoftdoc.py                      # verificacion

Python 3 puro, sin dependencias. Nada que instalar, nada que compilar.

## La decision de diseno que sostiene todo

**El formato no se reinventa: se copia el XML literal de `plantilla.docx`.**

`plantilla.docx` es un reporte real ya aprobado (el de Siegfried Ecuador). El
generador lo abre como ZIP, se queda con `word/document.xml`, recorta la portada
y la barra de seccion, y reconstruye un documento nuevo pegando ese XML con
bloques generados. Las otras 24 partes del paquete (estilos, numeracion, fuentes
Nunito embebidas, tema, pie de pagina, media) se copian byte a byte sin tocar.

Consecuencia practica: fuentes, colores, vinetas, bordes de tabla y pie de pagina
salen identicos siempre, porque nunca se escriben, se heredan. Un test verifica
que esas 24 partes queden intactas respecto de la plantilla.

Consecuencia igual de practica: **cambiar el formato de todos los documentos =
editar `plantilla.docx` en Word**, no tocar Python. Y si se edita la plantilla,
hay que revisar los indices duros del recorte (ver "Trampas").

## La segunda decision de diseno: el menu no vive en la generacion

`--nuevo` abre un menu (tipo de documento, bloques reutilizables, datos) y
escribe un `.md`. Generar ese `.md` no pregunta nada.

Esto es deliberado y conviene no revertirlo: si los bloques se eligieran al
generar, el mismo `.md` daria documentos distintos segun lo que se clickeo, se
perderia la regeneracion (que es el flujo principal), los bloques solo podrian
ir al final, habria que pedir las variables en cada corrida, y los tests no
podrian llamar a `generar()` sin colgarse.

Con el menu en `--nuevo` el `.md` queda como fuente de verdad unica:
versionable, diffeable, revisable por una persona. Y es el artefacto que la
futura automatizacion de redaccion va a producir o pre-llenar.

## Arquitectura de `mindsoftdoc.py` (un solo archivo)

Cinco secciones, en orden en el archivo:

| Lineas | Seccion | Que hace |
|---|---|---|
| Reglas de la casa | `limpiar()` normaliza caracteres prohibidos y lleva la cuenta |
| Bloques reutilizables | `incluir()` resuelve `@incluir` recursivo con tope de 10; `sustituir()` cambia `{{var}}` por el campo del front-matter |
| Plantilla | `_bloques()` parsea `<w:p>`/`<w:tbl>` de primer nivel; `Plantilla` recorta portada, barra, cabeza y cola |
| Piezas de XML | `runs()` para marcado en linea, y un `p_*()` por cada tipo de bloque |
| Lectura del markdown | `leer_markdown()` devuelve `(meta, bloques)`, cada bloque un `(tipo, dato)` |
| Imagenes | `medir()` lee dimensiones de PNG/JPEG a mano; `escalar()` ajusta al ancho util |
| Armado | `generar()` recorre bloques y arma el XML; `_escribir()` rearma el ZIP |
| `--nuevo` | `TIPOS` y `BLOQUES` (listas literales), `armar_nuevo()` puro y testeable, `menu_nuevo()` que es solo entrada/salida |

El flujo completo es lineal y cabe en una frase: markdown -> `leer_markdown` ->
lista de bloques `(tipo, dato)` -> un `p_*()` por bloque -> concatenar con
`tpl.cabeza` y `tpl.cola` -> `_escribir()` rearma el ZIP cambiando cinco partes.

### Los tipos de bloque

`h2`, `sub`, `p`, `ul`, `ol`, `tabla`, `salto`, `img`. Agregar un elemento nuevo
al markdown son tres pasos: reconocerlo en `leer_markdown`, escribir su `p_*()`
copiando el XML de un documento real, y despacharlo en el `for` de `generar()`.

### Las cinco partes del ZIP que se reescriben

Todo lo demas se copia tal cual (`_escribir`, lineas 439-474):

- `word/document.xml` - el documento nuevo
- `word/header1.xml` - el primer `<w:t>` recibe el `encabezado`; el segundo se vacia
- `word/_rels/document.xml.rels` - relaciones de las imagenes agregadas (`rId1001+`)
- `word/settings.xml` - se inyecta `<w:updateFields>` si el documento lleva indice
- `[Content_Types].xml` - se agregan los MIME de jpeg/jpg/gif si faltan

## Reglas de la casa

Antes de que cualquier texto llegue al documento pasa por `limpiar()`: guiones
largos y medios a `-`, comillas tipograficas y angulares a `"` y `'`, puntos
suspensivos de un caracter a `...`, espacios duros a espacio normal. Y en
`runs()`: la `**negrita**` sale como texto normal en parrafos, vinetas y celdas
(se cuenta y se reporta igual); sigue viva solo en titulos, subtitulos y la
cabecera azul de las tablas, que son estilo, no marcado.

No es cosmetica: es un estandar de la empresa, y aplica sin importar quien
escriba el markdown (persona o modelo). Al terminar, el CLI informa que corrigio
(`Reglas de la casa: 28 guiones largos`). Nunca cambia nada en silencio.

Este mismo criterio deberia gobernar cualquier funcion nueva: **corregir siempre,
avisar siempre, nunca preguntar.**

## Campos del front-matter

| Campo | Para que sirve |
|---|---|
| `rotulo` | Palabra sobre el titulo en la portada. Por defecto `Reporte` |
| `titulo` | Titulo grande. Si falta, se toma el primer `#` del documento |
| `subtitulo` | Linea gris bajo el titulo |
| `encabezado` | Linea superior de cada pagina. Si falta, se usa el titulo |
| `indice` | `si` inserta un indice automatico despues de la portada |

El pie con la direccion y los telefonos de Mindsoft esta fijo en
`word/footer1.xml` de la plantilla y no se configura desde el markdown.

## Trampas conocidas (lo que rompe si se toca sin mirar)

1. **Indices duros en `Plantilla.__init__`** (lineas 79-84). La portada son los
   primeros 23 bloques del cuerpo, la barra de seccion es el bloque 24, y el
   rotulo/titulo/subtitulo viven en los indices 19/20/21. Hay tres `assert` que
   avisan si la plantilla cambia, pero si se edita `plantilla.docx` y se agrega
   o quita un parrafo antes de la portada, estos numeros se corren.

2. **`poner()` reemplaza el primer `<w:t>` del bloque** (lineas 386-388). Si en
   Word se parte el texto de la portada en varios runs (pasa al editar y
   corregir ortografia), solo se reemplaza el primero y queda texto viejo
   pegado. Sintoma: la portada dice "ReporteAnalisis de Seguridad - Siegfried".

3. **ONLYOFFICE reescribe estilos al guardar.** Por eso se quito el sombreado
   zebra de las tablas (commit `e0851cb`): no sobrevivia. Si hay que editar a
   mano, hacerlo en Word; mejor, editar el markdown y regenerar.

   Guardar `plantilla.docx` desde un editor tambien **renumera los ids**: los
   `styleId` (`Heading2` paso a `788`, `Normal` a `786`) y las listas de
   `numbering.xml` (la vineta paso de `numId 90` a `5`). Un `pStyle` o un
   `numId` que no resuelve no da error: el titulo sale sin negrita, sin color y
   sin tamano, y la vineta sale sin punto. Por eso el generador ya no los cita
   por id: `leer_estilos()` los busca por `w:name` y `leer_numeracion()` por su
   definicion (bullet / decimal). Del mismo guardado salio que faltara el
   `xmlns:pic` en `<w:document>` (`espacios()` lo repone) y que se perdieran las
   Nunito embebidas, que siguen citadas en `styles.xml` pero ya no viajan en el
   archivo.

4. **El indice lo arma Word al abrir.** Se inserta un campo TOC mas
   `updateFields`. En LibreOffice u ONLYOFFICE hay que refrescar con F9.

5. **La numeracion de titulos la pone Word, no el markdown.** `p_titulo` y
   `p_subtitulo` cuelgan de la lista multinivel `numId 92`, que `_escribir`
   inyecta en `word/numbering.xml` al vuelo (no vive en `plantilla.docx`: un
   guardado desde Word la borraria sin que nadie se entere). En el `.md` los
   titulos van sin numero; Word renumera solo al mover una seccion. Los `###`
   ahora si entran al indice (`TOC \o "1-3"`, con `outlineLvl` explicito en
   cada titulo, que es lo que hace que LibreOffice y ONLYOFFICE tambien puedan
   armarlo). El titulo del propio indice va sin numerar y sin estilo `Heading2`
   (lleva el aspecto copiado a mano en `COMO_TITULO`): con el estilo puesto,
   Word mete el indice como primera entrada de si mismo.

   Cada lista numerada del markdown cuelga de su propio `w:num` (100, 101, ...)
   sobre el `abstractNum 91` de la plantilla, con `startOverride`. Sin eso Word
   las encadena y la segunda lista del documento arranca donde termino la
   primera. Los numIds del documento y los inyectados en `numbering.xml` tienen
   que coincidir: si no, la lista sale sin numeros y en silencio.

6. **Los bloques de hosting y soporte llevan precios y specs reales.** Salen de
   propuestas de 2026. Hay que revisarlos antes de mandar cada propuesta; el
   generador no sabe si estan vigentes.

7. **Hay dos ofertas de hosting distintas en los ejemplos, y el bloque solo trae
   una.** BYD usa Debian 13 + CrowdSec + nftables; Inmoweb usa AlmaLinux +
   WHM/cPanel + Imunify360. No son versiones de lo mismo, son productos
   distintos. `bloques/hosting-dedicado.md` trae la de BYD, por ser la mas
   reciente. Si Mindsoft vende las dos, falta un segundo bloque.

8. **Hay dos formas de pago y el bloque solo cubre una.** Siegfried PMC y Eliana
   usan "50% anticipo / 50% al finalizar" (global); InmoWeb y BYD usan "50% al
   inicio de cada fase / 50% al fin de cada fase". `bloques/forma-de-pago-50-50.md`
   trae la segunda. Falta `forma-de-pago-anticipo-final.md`.

## Limitaciones actuales

- Bloques de codigo con ``` salen como parrafos normales, sin recuadro.
- Sin listas anidadas ni notas al pie.
- Sin tabla de figuras, sin numeracion automatica de figuras (el pie de figura
  es texto libre que escribe el autor).
- El menu de `--nuevo` necesita una terminal interactiva. No hay forma no
  interactiva de armar un `.md`; para eso se copia un esqueleto a mano.
- `{{variables}}` funciona en el cuerpo, no dentro del front-matter.

## Verificacion

`python3 test_mindsoftdoc.py` corre nueve pruebas: reglas de la casa, marcado en
linea (incluido el caso de texto que parece marcado sin serlo, tipo
`[DEPOSIT][CONTROLLER]` o `SQLSTATE[42S22]`), reparto de columnas, medicion de
imagenes, lectura del markdown, `@incluir` y variables (con el bloque que no
existe y el ciclo), comentarios que no llegan al documento, los 5 tipos
generando `.docx` valido con todos sus bloques, y documento generado (XML
valido, cero caracteres prohibidos, conteos de titulos/tablas/listas/saltos,
imagen embebida, y las 24 partes del formato intactas).

Al 2026-09-04 pasan las nueve.

Se valido ademas contra dos documentos reales ya aprobados:

- **Reporte de la cafeteria de Academia Cotopaxi.** Regenerado desde markdown da
  el mismo texto caracter por caracter una vez aplicadas las reglas de la casa, y
  la misma estructura (13 titulos, 14 tablas, 145 negritas).
- **Propuesta de Siegfried PMC** (2026-09-04, con el sistema de bloques ya
  puesto). Reconstruida en 45 lineas de markdown mas un `@incluir` del
  portafolio. Queda en `prueba/`, sin versionar, para comparar contra el
  original en `ejemplos/`.

## Archivos

| Archivo | Que es |
|---|---|
| `mindsoftdoc.py` | El generador. Un solo archivo, sin dependencias |
| `plantilla.docx` | El formato de la casa. Editar solo para cambiar el formato de todos los documentos |
| `bloques/` | Bloques reutilizables, extraidos de las propuestas reales. Texto, sin codigo |
| `esqueletos/` | Estructura de cada tipo de documento. Los `<!-- -->` guian a quien escribe y no llegan al `.docx` |
| `ejemplos/` | Propuestas y el estimation table reales, de referencia. **Sin versionar: llevan precios de clientes** |
| `prueba/` | Reconstrucciones para comparar contra los originales. Sin versionar, descartable |
| `ejemplo.md` / `ejemplo.docx` | Muestra de cada elemento soportado. Sirve de fixture del test |
| `ejemplo-captura.png` | Imagen del ejemplo, medida por el test (600x180) |
| `test_mindsoftdoc.py` | Verificacion |
| `README.md` | Manual de uso para quien escribe markdown |
| `HANDOFF.md` | Este documento, para quien retoma el proyecto |

## Contexto del negocio

El cuello de botella no es el formato, ya esta resuelto. El cuello de botella es
**escribir el contenido del markdown**. Hay un archivo de propuestas historicas
ya hechas que el dueno del proyecto puede aportar como material de referencia
para lo que venga.

## Que sigue

1. **Estandarizar la tabla de estimacion que ve el cliente.** Decision del dueno
   del proyecto, en conversacion con su equipo. Hoy los 4 ejemplos con costo usan
   4 formatos distintos: `Detalle/Tiempo/Complejidad`, `#/Descripcion/Pagina/Horas/Costo`,
   `Detalle/Descripcion/Tiempo/Costo` y `Fase/Tiempo/Inversion`. Los esqueletos
   ya proponen `Fase | Detalle | Tiempo | Inversion`, que es la de BYD.

   Al reconstruir Siegfried PMC aparecio que **probablemente hagan falta dos
   formatos, no uno**: los proyectos grandes (BYD, InmoWeb) estiman por fase y
   llevan costo por fila; los trabajos chicos (Siegfried PMC, Eliana) estiman por
   tarea y solo llevan costo en el total. Forzar los dos a una sola tabla
   deforma uno de los dos.
2. **`estimacion.py`**: leer `Estimation Table Projects 2026.xlsx` y escribir
   `bloques/estimacion-<cliente>.md`. Script aparte a proposito, para no meter un
   lector de Excel dentro del generador. Bloqueado por el punto 1.
3. **Automatizar la redaccion del `.md`.** El dueno del proyecto lo quiere pero no
   todavia. `--nuevo` ya deja el artefacto donde esa automatizacion va a escribir.

Sobre el estimation table, lo que ya se sabe de mirar las 5 hojas:

- El layout es estable: fila 7 es la cabecera (`Category | Task/Description |
  Best Case | Worst Case | Most likely | Cost | Requerido | MVP`), las tarifas
  por rol van en la fila 2, y `TOTAL VALUE` suma la columna MVP.
- En BYD el total del xlsx ($7.700) coincide exacto con el de la propuesta. El
  costo se puede automatizar.
- El tiempo **no** coincide: 3,5 semanas internas contra 16 semanas al cliente.
  El calendario es una decision comercial, no un calculo. No automatizarlo.
- Las fases se marcan en la columna A (`Fase de Analisis`) pero los sprints en la
  columna B (`Sprint 1`). Hay que fijar una convencion antes de parsear.
