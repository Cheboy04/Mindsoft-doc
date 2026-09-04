# mindsoftdoc

Genera documentos Word con el formato de la casa Mindsoft a partir de un
archivo markdown. El formato no se reinventa: sale del XML literal de
`plantilla.docx`, asi que fuentes, colores, encabezado, pie, vinetas y tablas
salen identicos siempre.

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

Al terminar informa que corrigio, por ejemplo `Reglas de la casa: 28 guiones
largos`. Nunca cambia nada en silencio, y aplica sin importar quien escriba el
markdown.

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
| `## Titulo` | Titulo de seccion azul con la barra debajo |
| `### Subtitulo` | Negrita de 11 puntos |
| texto suelto | Parrafo Verdana 10, justificado |
| `- item` | Vineta |
| `1. item` | Lista numerada |
| tabla con pipes | Cabecera azul y filas en blanco |
| `**negrita**` `*cursiva*` `` `codigo` `` | Formato en linea |
| `[texto](url)` | Texto, con la url entre parentesis si es http |
| `![pie](captura.png)` | Imagen centrada con pie de figura |
| `---` en una linea | Salto de pagina |

Las columnas de las tablas se reparten solas segun cuanto texto lleva cada una.
Las imagenes se escalan al ancho util de la pagina sin deformarse; PNG y JPEG.

Los `###` no entran al indice, a proposito: el indice lista solo las secciones.

## Verificacion

    python3 test_mindsoftdoc.py

Comprueba las reglas de la casa, el marcado en linea, el reparto de columnas,
la medicion de imagenes, la lectura del markdown, y que el documento generado
sea XML valido con las 24 partes del formato intactas respecto de la plantilla.

Se valido ademas contra un documento real ya aprobado, el reporte de la
cafeteria de Academia Cotopaxi: regenerado desde markdown da el mismo texto
caracter por caracter una vez aplicadas las reglas de la casa, y la misma
estructura (13 titulos, 14 tablas, 145 negritas).

## Limitaciones conocidas

- Los bloques de codigo con ``` salen como parrafos normales. Si los necesitan
  en recuadro, se agrega.
- Sin listas anidadas ni notas al pie.
- El indice lo arma Word al abrir el archivo. En LibreOffice u ONLYOFFICE hay
  que refrescarlo a mano con F9.
- Guardar el documento desde ONLYOFFICE reescribe los estilos. Si hay que
  editar a mano, conviene hacerlo en Word, o mejor editar el markdown y
  regenerar.

## Archivos

| Archivo | Que es |
|---|---|
| `mindsoftdoc.py` | El generador |
| `plantilla.docx` | El formato de la casa. No editar salvo para cambiar el formato de todos los documentos |
| `ejemplo.md` / `ejemplo.docx` | Muestra de cada elemento soportado |
| `test_mindsoftdoc.py` | Verificacion |
