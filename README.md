# mindsoftdoc

Genera documentos Word con el formato de la casa Mindsoft a partir de un
archivo markdown. El formato no se reinventa: sale del XML literal de
`plantilla.docx`, asi que fuentes, colores, encabezado, pie, vinetas y tablas
salen identicos siempre.

    python3 mindsoftdoc.py --nuevo                   # menu: arma el .md de arranque
    python3 mindsoftdoc.py documento.md              # deja documento.docx al lado
    python3 mindsoftdoc.py documento.md -o final.docx

Solo necesita Python 3. No hay dependencias que instalar.

## Reglas de la casa

Antes de que cualquier texto llegue al documento pasa por un saneador que
reemplaza:

| Entra | Sale |
|---|---|
| guiones largos y medios | `-` |
| comillas tipograficas y angulares | `"` y `'` |
| puntos suspensivos de un caracter | `...` |
| espacios duros | espacio normal |
| `**negrita**` en cualquier texto | texto normal (la negrita solo vive en titulos y cabeceras de tabla) |

Al terminar informa que corrigio, por ejemplo `Reglas de la casa: 28 guiones
largos`. Nunca cambia nada en silencio, y aplica sin importar quien escriba el
markdown.

## Empezar un documento: `--nuevo`

El menu pregunta dos cosas y escribe el `.md` de arranque:

    $ python3 mindsoftdoc.py --nuevo

    1. Tipo de documento
      1) Propuesta con costos       3) Documento de analisis    5) Cotizacion
      2) Propuesta sin costos       4) Guia / informativo
    > 1

    2. Bloques reutilizables
      1) Forma de pago 50/50           [x]    4) Manejo de hosting externo
      2) SEM / Google Ads                     5) Plan de soporte mensual
      3) Hosting dedicado                     6) Portafolio de clientes  [x]
    > 1 2 6

    3. Datos
       titulo: Lanzamiento + Agente AI
       cliente: InmoWeb

    Escrito: inmoweb.md

Los bloques que se ofrecen dependen del tipo elegido, y las variables que se
preguntan salen de los propios bloques. Con `--nuevo propuesta-con-costos` se
salta la primera pregunta.

El menu esta aca a proposito y no en la generacion: **generar un `.md` tiene que
dar siempre el mismo `.docx`**. El `.md` es la fuente de verdad, se versiona y se
regenera cuando haga falta.

## Bloques reutilizables

Una linea sola, en el lugar del documento donde va el bloque:

    @incluir bloques/portafolio-clientes.md

Se busca junto al `.md` y, si no esta, junto al generador; asi `bloques/...`
funciona desde cualquier carpeta. Un bloque que no existe corta con un error, no
se ignora.

Dentro de un bloque, `{{variable}}` se reemplaza con el campo del front-matter
que tenga ese nombre:

    Se propone posicionar estrategicamente a {{cliente}} en Ecuador.

Lo que no tenga valor se deja a la vista en el documento y se avisa al terminar
(`AVISO - variables sin valor: cliente`). Nunca se borra en silencio.

| Bloque | Va en |
|---|---|
| `bloques/portafolio-clientes.md` | Propuestas y cotizaciones |
| `bloques/forma-de-pago-50-50.md` | Costo por fases: 50% al inicio y 50% al fin de cada fase |
| `bloques/forma-de-pago-anticipo-final.md` | Costo global: 50% a la firma y 50% contra entrega |
| `bloques/sem-google-ads.md` | Propuestas con campana |
| `bloques/hosting-dedicado.md` | Infraestructura sobre Debian, sin panel |
| `bloques/hosting-dedicado-cpanel.md` | Infraestructura sobre AlmaLinux con WHM/cPanel |
| `bloques/hosting-externo.md` | Propuestas con infraestructura |
| `bloques/soporte-mensual.md` | Propuestas con soporte |

Las specs y los precios de los bloques de hosting y soporte salen de propuestas
reales: **revisalos en cada propuesta antes de mandarla**.

Hay dos pares de bloques que se excluyen entre si: se elige uno, no los dos.

- Forma de pago: `forma-de-pago-50-50.md` es la version por fases (InmoWeb, BYD);
  `forma-de-pago-anticipo-final.md` es la global (Siegfried PMC, Eliana).
- Hosting: `hosting-dedicado.md` es la oferta sobre Debian, sin panel (BYD);
  `hosting-dedicado-cpanel.md` es la de AlmaLinux con WHM/cPanel (Inmoweb). Este
  ultimo pide `hosting_vcpu`, `hosting_ram`, `hosting_disco`, `hosting_mensual` y
  `hosting_anual` en el front-matter.

Los `<!-- comentarios -->` de los esqueletos no llegan al documento.

## El archivo de entrada

```
---
rotulo: Reporte
titulo: Analisis de Seguridad - Sistema de Cafeteria
subtitulo: Incidente de reembolso · cotopaxi.k12.ec · Septiembre 2026
encabezado: Analisis de Seguridad - Cafeteria AC | Septiembre 2026
indice: si
---

## Primera seccion
...
```

| Campo | Para que sirve |
|---|---|
| `rotulo` | La palabra sobre el titulo en la portada: Reporte, Manual, Especificacion Funcional, Propuesta. Por defecto `Reporte` |
| `titulo` | Titulo grande de la portada. Si falta, se toma el primer `#` del documento |
| `subtitulo` | Linea gris bajo el titulo |
| `encabezado` | Linea superior de cada pagina. Si falta, se usa el titulo |
| `indice` | `si` inserta un indice automatico despues de la portada |

El pie con la direccion y los telefonos de Mindsoft es fijo y no se configura.

## Que reconoce del markdown

| Escribis | Sale como |
|---|---|
| `## Titulo` | Titulo de seccion azul con la barra debajo, numerado 1, 2, 3 |
| `### Subtitulo` | Negrita azul de 11 puntos, numerado 1.1, 1.2 |
| texto suelto | Parrafo Verdana 10, justificado |
| `- item` | Vineta |
| `1. item` | Lista numerada, cada lista arranca de nuevo en 1 |
| tabla con pipes | Cabecera azul y filas en blanco, texto a la izquierda |
| `*cursiva*` `` `codigo` `` | Formato en linea |
| `**negrita**` | Nada: sale como texto normal (regla de la casa) |
| `[texto](url)` | Texto, con la url entre parentesis si es http |
| `![pie](captura.png)` | Imagen centrada con pie de figura |
| `> cita` | Parrafo normal: se le quita el `>` |
| `---` en una linea | Salto de pagina |

Las columnas de las tablas se reparten solas segun cuanto texto lleva cada una.
Las imagenes se escalan al ancho util de la pagina sin deformarse; PNG y JPEG.

La numeracion la pone Word, no el markdown: en el `.md` los titulos van sin
numero y Word los renumera solo cuando se mueve o se agrega una seccion. Los
`###` entran al indice con su numero (`2.1`), debajo de su seccion.

## Verificacion

    python3 test_mindsoftdoc.py

Comprueba las reglas de la casa, el marcado en linea, el reparto de columnas,
la medicion de imagenes, la lectura del markdown, los `@incluir` y las
variables (incluidos el bloque que falta y el ciclo), que los comentarios no
lleguen al documento, que los 5 tipos generen un `.docx` valido con todos sus
bloques, y que el documento generado sea XML valido, con los estilos y las
listas resueltos contra la plantilla y las 21 partes del formato intactas.

Se valido ademas contra dos documentos reales ya aprobados:

- El reporte de la cafeteria de Academia Cotopaxi: regenerado desde markdown da
  el mismo texto caracter por caracter una vez aplicadas las reglas de la casa,
  y la misma estructura (13 titulos, 14 tablas). Las 145 negritas del original
  hoy saldrian como texto normal: la regla de la casa es posterior a esa prueba.
- La propuesta de Siegfried PMC, ya con bloques: 45 lineas de markdown mas un
  `@incluir` del portafolio.

## Limitaciones conocidas

- Los bloques de codigo con ``` salen como parrafos normales. Si los necesitan
  en recuadro, se agrega.
- Sin listas anidadas ni notas al pie.
- `--nuevo` necesita una terminal interactiva. Para armar un `.md` sin menu, se
  copia a mano el esqueleto de `esqueletos/`.
- `{{variables}}` funciona en el cuerpo del documento, no dentro del
  front-matter.
- El indice lo arma Word al abrir el archivo. En LibreOffice u ONLYOFFICE hay
  que refrescarlo a mano con F9.
- Guardar el documento desde ONLYOFFICE reescribe los estilos. Si hay que
  editar a mano, conviene hacerlo en Word, o mejor editar el markdown y
  regenerar.
- Guardar `plantilla.docx` desde un editor renumera los ids internos de estilos
  y listas. El generador los resuelve por nombre, asi que no se rompe, pero ese
  guardado se lleva las tipografias Nunito embebidas: si el documento tiene que
  verse igual en una maquina sin Nunito instalada, hay que volver a embeberlas.

## Archivos

| Archivo | Que es |
|---|---|
| `mindsoftdoc.py` | El generador |
| `plantilla.docx` | El formato de la casa. No editar salvo para cambiar el formato de todos los documentos |
| `bloques/` | Bloques reutilizables. Texto, sin codigo |
| `esqueletos/` | Estructura de cada tipo de documento |
| `ejemplo.md` / `ejemplo.docx` | Muestra de cada elemento soportado |
| `test_mindsoftdoc.py` | Verificacion |
