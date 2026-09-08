#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""mindsoftdoc - genera documentos Word con el formato de la casa Mindsoft.

    python3 mindsoftdoc.py documento.md [-o salida.docx]

El formato sale de plantilla.docx: se reutiliza su XML literal (fuentes,
colores, encabezado, pie, vinetas, tablas), no se reinventa.
"""
import os, re, struct, zipfile, shutil, argparse

AQUI = os.path.dirname(os.path.abspath(__file__))
PLANTILLA = os.path.join(AQUI, 'plantilla.docx')

ANCHO_TABLA = 8770           # dxa, ancho util de la plantilla
ANCHO_EMU = 5573430          # ancho util de pagina en EMU
EMU_POR_PX = 9525            # a 96 dpi

# --------------------------------------------------------------------------
# Reglas de la casa: nada de guiones largos, nada de comillas tipograficas
# --------------------------------------------------------------------------
REEMPLAZOS = [
    ('guiones largos',      {'—': '-', '–': '-', '‒': '-', '―': '-', '−': '-'}),
    ('comillas',            {'‘': "'", '’': "'", '‚': "'", '‛': "'",
                             '“': '"', '”': '"', '„': '"', '‟': '"',
                             '′': "'", '″': '"', '«': '"', '»': '"'}),
    ('puntos suspensivos',  {'…': '...'}),
    ('espacios duros',      {' ': ' ', ' ': ' ', ' ': ' '}),
]
_conteo = {}
_faltantes = set()

def limpiar(texto):
    """Aplica las reglas de la casa y lleva la cuenta de lo reemplazado."""
    for etiqueta, mapa in REEMPLAZOS:
        for malo, bueno in mapa.items():
            n = texto.count(malo)
            if n:
                _conteo[etiqueta] = _conteo.get(etiqueta, 0) + n
                texto = texto.replace(malo, bueno)
    return texto

def resumen_limpieza():
    return ', '.join('%d %s' % (n, k) for k, n in sorted(_conteo.items())) or 'nada que corregir'

# --------------------------------------------------------------------------
# Bloques reutilizables: @incluir y {{variables}}
# --------------------------------------------------------------------------
INCLUIR = re.compile(r'^@incluir\s+(\S+)\s*$')
VARIABLE = re.compile(r'\{\{\s*(\w+)\s*\}\}')

def _ruta_bloque(nombre, base):
    """Busca el bloque junto al .md y, si no, junto al generador."""
    for cand in (os.path.join(base, nombre), os.path.join(AQUI, nombre)):
        if os.path.exists(cand):
            return cand
    raise SystemExit('No encuentro el bloque: %s (buscado en %s y %s)'
                     % (nombre, base, AQUI))

def incluir(lineas, base, profundidad=0):
    """Reemplaza cada linea '@incluir ruta' por el contenido de ese archivo."""
    if profundidad > 10:
        raise SystemExit('@incluir anidado demasiado profundo: hay un ciclo?')
    out = []
    for l in lineas:
        m = INCLUIR.match(l.strip())
        if not m:
            out.append(l); continue
        ruta = _ruta_bloque(m.group(1), base)
        with open(ruta, encoding='utf8') as f:
            dentro = f.read().replace('\r\n', '\n').split('\n')
        # las lineas en blanco evitan que el bloque se pegue al texto vecino
        out.append('')
        out += incluir(dentro, os.path.dirname(ruta), profundidad + 1)
        out.append('')
    return out

def sustituir(lineas, meta):
    """Cambia {{variable}} por su valor del front-matter. Lo que falta se
    deja a la vista y se reporta: nunca se borra en silencio."""
    def cambiar(m):
        if m.group(1) in meta:
            return meta[m.group(1)]
        _faltantes.add(m.group(1))
        return m.group(0)
    return [VARIABLE.sub(cambiar, l) for l in lineas]

def resumen_variables():
    return ', '.join(sorted(_faltantes))

# --------------------------------------------------------------------------
# Plantilla: de ahi salen la portada, la barra de seccion y el paquete
# --------------------------------------------------------------------------
def _bloques(xml):
    """Devuelve los <w:p>/<w:tbl> de primer nivel, respetando anidamiento."""
    out, i = [], 0
    tok = re.compile(r'<w:(p|tbl)(?:\s[^>]*)?(/?)>|</w:(p|tbl)>')
    while True:
        m = tok.search(xml, i)
        if not m:
            return out
        if m.group(3):
            i = m.end(); continue
        if m.group(2) == '/':
            out.append(xml[m.start():m.end()]); i = m.end(); continue
        nombre, prof, j = m.group(1), 1, m.end()
        while prof:
            m2 = tok.search(xml, j)
            if not m2:
                break
            if m2.group(3) == nombre:
                prof -= 1
            elif m2.group(1) == nombre and m2.group(2) != '/':
                prof += 1
            j = m2.end()
        out.append(xml[m.start():j]); i = j

# Un guardado de la plantilla desde Word o ONLYOFFICE renumera los styleId
# ("Heading2" paso a llamarse "788"), y un pStyle que no resuelve deja el titulo
# sin negrita, sin color y sin tamano. El w:name si es estable, asi que los
# estilos se buscan por nombre y estos valores son solo el punto de partida.
ESTILOS = {'normal': 'Normal', 'heading 2': 'Heading2', 'heading 3': 'Heading3'}

def leer_estilos(xml):
    for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>', xml, re.S):
        nombre = re.search(r'<w:name w:val="([^"]+)"', m.group(2))
        if nombre and nombre.group(1).lower() in ESTILOS:
            ESTILOS[nombre.group(1).lower()] = m.group(1)

class Plantilla:
    def __init__(self, ruta=PLANTILLA):
        self.ruta = ruta
        with zipfile.ZipFile(ruta) as z:
            self.doc = z.read('word/document.xml').decode('utf8')
            leer_estilos(z.read('word/styles.xml').decode('utf8'))
            leer_numeracion(z.read('word/numbering.xml').decode('utf8'))
        cuerpo = self.doc[self.doc.index('<w:body>') + 8:self.doc.rindex('<w:sectPr')]
        els = _bloques(cuerpo)
        # Los 23 primeros bloques son la portada (el ultimo es su salto de pagina).
        # El 23 es el primer titulo del reporte original y se descarta; el 24 es
        # su barra de seccion, que se reutiliza debajo de cada titulo generado.
        assert len(els) > 24, 'plantilla.docx no tiene la estructura esperada'
        self.portada = list(els[:23])
        self.barra = els[24]
        assert 'Reporte' in self.portada[19], 'no encuentro el rotulo en la portada'
        assert '<w:drawing>' in self.barra, 'no encuentro la barra de seccion'
        self.cabeza = self.doc[:self.doc.index('<w:body>') + 8]
        self.cola = self.doc[self.doc.rindex('<w:sectPr'):]

# --------------------------------------------------------------------------
# Piezas de XML, copiadas literalmente de la plantilla
# --------------------------------------------------------------------------
def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

_R = '<w:r><w:rPr>%s<w:rtl w:val="0"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
NORMAL = '<w:sz w:val="20"/><w:szCs w:val="20"/>'
CURSIVA = '<w:i w:val="1"/><w:iCs w:val="1"/><w:sz w:val="20"/><w:szCs w:val="20"/>'
CODIGO = '<w:rFonts w:ascii="Consolas" w:cs="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
C_NORMAL = '<w:sz w:val="18"/><w:szCs w:val="18"/>'
C_CURSIVA = '<w:i w:val="1"/><w:iCs w:val="1"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
C_CODIGO = '<w:rFonts w:ascii="Consolas" w:cs="Consolas" w:hAnsi="Consolas"/><w:sz w:val="16"/><w:szCs w:val="16"/>'
C_TITULO = '<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="FFFFFF"/><w:sz w:val="18"/><w:szCs w:val="18"/>'

TOKEN = re.compile(r'(\*\*.+?\*\*|`[^`]+`|\*[^*]+\*|\[[^\]]+\]\([^)]+\))')
ENLACE = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')

def runs(texto, celda=False):
    """Convierte marcado en linea (**negrita**, `codigo`, *cursiva*, enlaces)."""
    normal, cursiva, codigo = ((C_NORMAL, C_CURSIVA, C_CODIGO)
                               if celda else (NORMAL, CURSIVA, CODIGO))
    out = []
    # re.split con un grupo devuelve texto suelto en los indices pares y los
    # tokens capturados en los impares. Hay que mirar la posicion, no el primer
    # caracter: un texto normal como "[DEPOSIT][CONTROLLER]" empieza con "[" y
    # no es un enlace.
    for i, parte in enumerate(TOKEN.split(texto)):
        if not parte:
            continue
        if i % 2 == 0:
            out.append(_R % (normal, esc(parte)))
        elif parte.startswith('**'):
            # regla de la casa: la negrita solo vive en los titulos
            _conteo['negritas'] = _conteo.get('negritas', 0) + 1
            out.append(_R % (normal, esc(parte[2:-2])))
        elif parte.startswith('`'):
            out.append(_R % (codigo, esc(parte[1:-1])))
        elif parte.startswith('*'):
            out.append(_R % (cursiva, esc(parte[1:-1])))
        else:
            m = ENLACE.match(parte)
            texto_enlace, url = m.group(1), m.group(2)
            if url.startswith('http'):
                texto_enlace = '%s (%s)' % (texto_enlace, url)
            out.append(_R % (normal, esc(texto_enlace)))
    return ''.join(out) or _R % (normal, '')

# La numeracion 1, 1.1, 1.2 la pone Word, no el markdown: asi renumera solo
# cuando se mueve una seccion y los subtitulos entran al indice con su numero.
# Estos tres ids son los que el generador *agrega* a numbering.xml. No se
# fijan a mano: leer_numeracion() los corre por encima de los que ya trae la
# plantilla. Pisar un id existente no da error, solo saca la numeracion.
NUM_TITULOS = 92       # numId de la lista multinivel de los titulos
ABS_TITULOS = 92       # su abstractNumId
BASE_LISTAS = 100      # numId de la primera lista numerada del documento
_LVL = ('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/><w:suff w:val="space"/>'
        '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>')

def numeracion_titulos():
    return ('<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="multilevel"/>%s'
            '</w:abstractNum><w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
            % (ABS_TITULOS,
               ''.join(_LVL % (i, t) for i, t in enumerate(('%1.', '%1.%2', '%1.%2.%3'))),
               NUM_TITULOS, ABS_TITULOS))

# Mismo cuento que los estilos: un guardado de la plantilla renumera las listas.
# Se buscan por su definicion (vineta / decimal), no por el id, que no sobrevive.
NUM_VINETA = 90        # numId de la vineta, se usa tal cual
ABS_NUMERADA = 91      # abstractNumId de la numerada, cada lista cuelga uno propio

def leer_numeracion(xml):
    global NUM_VINETA, ABS_NUMERADA, NUM_TITULOS, ABS_TITULOS, BASE_LISTAS
    formato = {}
    for m in re.finditer(r'<w:abstractNum w:abstractNumId="(\d+)"[^>]*>(.*?)</w:abstractNum>',
                         xml, re.S):
        lvl = re.search(r'<w:lvl w:ilvl="0".*?</w:lvl>', m.group(2), re.S)
        fmt = re.search(r'<w:numFmt w:val="(\w+)"', lvl.group(0)) if lvl else None
        if fmt:
            formato[m.group(1)] = (fmt.group(1), '<w:lvlText w:val="\u2022"' in lvl.group(0))
    nums = re.findall(r'<w:num w:numId="(\d+)"\s*>\s*<w:abstractNumId w:val="(\d+)"\s*/>', xml)
    # la vineta de la casa es el punto redondo; si no esta, sirve cualquier vineta
    vinetas = [(n, formato[a][1]) for n, a in nums if formato.get(a, ('',))[0] == 'bullet']
    if vinetas:
        NUM_VINETA = int(sorted(vinetas, key=lambda v: not v[1])[0][0])
    decimales = [a for a, (f, _) in formato.items() if f == 'decimal']
    if decimales:
        ABS_NUMERADA = int(decimales[0])
    # lo que se inyecta va por encima de todo lo que ya existe, para no pisarlo
    NUM_TITULOS = max([int(n) for n, _ in nums] + [0]) + 1
    BASE_LISTAS = NUM_TITULOS + 1
    ABS_TITULOS = max([int(a) for a in re.findall(r'<w:abstractNum w:abstractNumId="(\d+)"', xml)]
                      + [0]) + 1

# Las listas numeradas del markdown son independientes entre si: cada una
# cuelga de su propio w:num sobre el mismo abstractNum de la plantilla, con
# startOverride, o Word las encadena y la segunda arranca donde termino la primera.
def id_lista(n):
    return BASE_LISTAS + n   # el numId de la lista numerada n del documento

def num_lista(n):
    return ('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/><w:lvlOverride w:ilvl="0">'
            '<w:startOverride w:val="1"/></w:lvlOverride></w:num>' % (id_lista(n), ABS_NUMERADA))

def numpr(nivel):
    return ('<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>' % (nivel, NUM_TITULOS))

# El aspecto de Heading2, copiado a mano: lo usa el titulo del indice, que no
# puede llevar el estilo o Word lo mete como primera entrada del propio indice.
COMO_TITULO = ('<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="001689"/>'
               '<w:sz w:val="28"/><w:szCs w:val="28"/>')

def p_titulo(texto, barra, numerado=True):
    pPr = (('<w:pStyle w:val="%s"/>' % ESTILOS['heading 2'] +
            '<w:pageBreakBefore w:val="0"/>' + numpr(0) +
            '<w:spacing w:before="240" w:line="276" w:lineRule="auto"/>'
            '<w:outlineLvl w:val="1"/><w:rPr/>') if numerado else
           ('<w:pageBreakBefore w:val="0"/>'
            '<w:spacing w:before="240" w:line="276" w:lineRule="auto"/>'))
    return ('<w:p><w:pPr>%s</w:pPr><w:r><w:rPr>%s<w:rtl w:val="0"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (pPr, '' if numerado else COMO_TITULO, esc(texto))) + barra

def p_parrafo(texto):
    return ('<w:p><w:pPr><w:pageBreakBefore w:val="0"/>'
            '<w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr>%s</w:p>' % runs(texto))

# El mismo azul de los titulos, en 11 puntos. Va dos veces: en el run del texto
# y en el rPr de la marca de parrafo, que es de donde Word saca el formato del
# numero de la lista (sin eso el 2.1 sale en negro y en 9 puntos).
SUBTITULO = ('<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="001689"/>'
             '<w:sz w:val="22"/><w:szCs w:val="22"/>')

def p_subtitulo(texto):
    return ('<w:p><w:pPr><w:pStyle w:val="%s"/>' % ESTILOS['heading 3'] +
            '<w:pageBreakBefore w:val="0"/>' + numpr(1) +
            '<w:spacing w:after="60" w:line="264" w:lineRule="auto"/>'
            '<w:outlineLvl w:val="2"/><w:rPr>' + SUBTITULO + '</w:rPr></w:pPr>'
            '<w:r><w:rPr>' + SUBTITULO + '<w:rtl w:val="0"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % esc(texto))

def p_item(texto, num_id, despues):
    return ('<w:p><w:pPr><w:pStyle w:val="%s"/>' % ESTILOS['normal'] +
            '<w:numPr><w:ilvl w:val="0"/>'
            '<w:numId w:val="%d"/></w:numPr><w:spacing w:after="%d" w:line="264" w:lineRule="auto"/>'
            '</w:pPr>%s</w:p>' % (num_id, despues, runs(texto)))

ESPACIADOR = ('<w:p><w:pPr><w:pageBreakBefore w:val="0"/>'
              '<w:spacing w:after="40" w:line="264" w:lineRule="auto"/></w:pPr>'
              '<w:r><w:rPr><w:sz w:val="8"/><w:szCs w:val="8"/><w:rtl w:val="0"/></w:rPr>'
              '<w:t xml:space="preserve"></w:t></w:r></w:p>')

SALTO = ('<w:p><w:pPr><w:spacing w:line="264" w:lineRule="auto"/></w:pPr>'
         '<w:r><w:br w:type="page"/></w:r></w:p>')

def _celda(ancho, fondo, texto, titulo=False):
    cuerpo = _R % (C_TITULO, esc(texto)) if titulo else runs(texto, celda=True)
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:shd w:val="clear" w:fill="%s"/>'
            '<w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/>'
            '<w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>'
            '<w:vAlign w:val="center"/></w:tcPr>'
            '<w:p><w:pPr><w:spacing w:after="20" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="left"/></w:pPr>%s</w:p>'
            '</w:tc>' % (ancho, fondo, cuerpo))

def anchos(cabeceras, filas, total=ANCHO_TABLA, minimo=700):
    """Reparte el ancho segun cuanto texto lleva cada columna."""
    n = len(cabeceras)
    peso = [max([len(cabeceras[i])] + [len(f[i]) for f in filas]) or 1 for i in range(n)]
    libre = total - minimo * n
    if libre < 0:
        return [total // n] * (n - 1) + [total - (total // n) * (n - 1)]
    suma = float(sum(peso))
    w = [minimo + int(libre * p / suma) for p in peso]
    w[peso.index(max(peso))] += total - sum(w)
    return w

def p_tabla(cabeceras, filas):
    w = anchos(cabeceras, filas)
    assert sum(w) == ANCHO_TABLA
    x = ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>'
         '<w:top w:val="single" w:sz="4" w:color="BFC7DA"/><w:left w:val="single" w:sz="4" w:color="BFC7DA"/>'
         '<w:bottom w:val="single" w:sz="4" w:color="BFC7DA"/><w:right w:val="single" w:sz="4" w:color="BFC7DA"/>'
         '<w:insideH w:val="single" w:sz="4" w:color="BFC7DA"/><w:insideV w:val="single" w:sz="4" w:color="BFC7DA"/>'
         '</w:tblBorders></w:tblPr><w:tblGrid>%s</w:tblGrid>'
         % (ANCHO_TABLA, ''.join('<w:gridCol w:w="%d"/>' % a for a in w)))
    x += ('<w:tr><w:trPr><w:tblHeader/></w:trPr>%s</w:tr>'
          % ''.join(_celda(a, '001689', c, True) for a, c in zip(w, cabeceras)))
    for fila in filas:
        # Estandar de la casa: cabecera azul y filas blancas, sin sombreado
        # alternado. El zebra no sobrevive a un guardado desde ONLYOFFICE.
        x += '<w:tr>%s</w:tr>' % ''.join(_celda(a, 'FFFFFF', c) for a, c in zip(w, fila))
    return x + '</w:tbl>' + ESPACIADOR

def p_imagen(rid, ancho_emu, alto_emu, indice):
    return ('<w:p><w:pPr><w:pageBreakBefore w:val="0"/><w:spacing w:after="60" w:line="264" '
            'w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr/><w:drawing>'
            '<wp:inline distB="114300" distT="114300" distL="114300" distR="114300">'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent b="0" l="0" r="0" t="0"/>'
            '<wp:docPr id="%d" name="imagen%d"/><a:graphic><a:graphicData '
            'uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
            '<pic:cNvPr id="0" name="imagen%d"/><pic:cNvPicPr preferRelativeResize="0"/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch>'
            '</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"/><a:ln/></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing></w:r></w:p>'
            % (ancho_emu, alto_emu, 100 + indice, indice, indice, rid, ancho_emu, alto_emu))

def p_pie_figura(texto):
    return ('<w:p><w:pPr><w:pageBreakBefore w:val="0"/><w:spacing w:after="120" w:line="264" '
            'w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
            '<w:r><w:rPr><w:i w:val="1"/><w:iCs w:val="1"/><w:color w:val="666666"/>'
            '<w:sz w:val="16"/><w:szCs w:val="16"/><w:rtl w:val="0"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % esc(texto))

def p_indice():
    return (p_titulo('Indice', '', numerado=False)
            + '<w:p><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr>'
              '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
              '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
              '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
              '<w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
              '<w:t xml:space="preserve">Abra el documento en Word para armar el indice, '
              'o pulse F9 sobre esta linea.</w:t></w:r>'
              '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + SALTO)

# --------------------------------------------------------------------------
# Lectura del markdown
# --------------------------------------------------------------------------
SEP_TABLA = re.compile(r'^\|[\s:|-]+\|$')
IMAGEN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
VINETA = re.compile(r'^\s*[-*+]\s+(.*)$')
NUMERADA = re.compile(r'^\s*\d+[.)]\s+(.*)$')

def leer_markdown(texto, base=None):
    """Devuelve (metadatos, bloques). Cada bloque es (tipo, dato)."""
    base = base or AQUI
    lineas = texto.replace('\r\n', '\n').split('\n')
    meta = {}
    if lineas and lineas[0].strip() == '---':
        fin = next((i for i in range(1, len(lineas)) if lineas[i].strip() == '---'), None)
        if fin:
            for l in lineas[1:fin]:
                if ':' in l:
                    k, v = l.split(':', 1)
                    meta[k.strip().lower()] = v.strip()
            lineas = lineas[fin + 1:]

    # el front-matter va primero porque de ahi salen las variables
    texto = '\n'.join(sustituir(incluir(lineas, base), meta))
    texto = re.sub(r'<!--.*?-->', '', texto, flags=re.S)   # notas del esqueleto
    lineas = limpiar(texto).split('\n')

    bloques, i, parrafo = [], 0, []

    def cerrar():
        if parrafo:
            bloques.append(('p', ' '.join(parrafo)))
            del parrafo[:]

    while i < len(lineas):
        l = lineas[i]
        cruda = l.rstrip()
        s = cruda.strip()

        if not s:
            cerrar(); i += 1; continue

        if s == '---' or s == '***':
            cerrar(); bloques.append(('salto', None)); i += 1; continue

        m = IMAGEN.match(s)
        if m:
            cerrar(); bloques.append(('img', (m.group(2), m.group(1)))); i += 1; continue

        if s.startswith('#'):
            cerrar()
            nivel = len(s) - len(s.lstrip('#'))
            titulo = s.lstrip('#').strip()
            if nivel == 1 and 'titulo' not in meta:
                meta['titulo'] = titulo
            elif nivel <= 2:
                bloques.append(('h2', titulo))
            else:
                bloques.append(('sub', titulo))
            i += 1; continue

        if s.startswith('>'):
            cerrar(); bloques.append(('p', s.lstrip('>').strip())); i += 1; continue

        if s.startswith('|') and i + 1 < len(lineas) and SEP_TABLA.match(lineas[i + 1].strip()):
            cerrar()
            def partir(fila):
                return [c.strip() for c in fila.strip().strip('|').split('|')]
            cabeceras = partir(s)
            filas, i = [], i + 2
            while i < len(lineas) and lineas[i].strip().startswith('|'):
                fila = partir(lineas[i])
                fila += [''] * (len(cabeceras) - len(fila))
                filas.append(fila[:len(cabeceras)])
                i += 1
            bloques.append(('tabla', (cabeceras, filas)))
            continue

        m = VINETA.match(cruda)
        if m:
            cerrar()
            items = []
            while i < len(lineas):
                mm = VINETA.match(lineas[i].rstrip())
                if mm:
                    items.append(mm.group(1).strip()); i += 1
                elif lineas[i].strip() and lineas[i].startswith((' ', '\t')):
                    items[-1] += ' ' + lineas[i].strip(); i += 1
                else:
                    break
            bloques.append(('ul', items)); continue

        m = NUMERADA.match(cruda)
        if m:
            cerrar()
            items = []
            while i < len(lineas):
                mm = NUMERADA.match(lineas[i].rstrip())
                if mm:
                    items.append(mm.group(1).strip()); i += 1
                elif lineas[i].strip() and lineas[i].startswith((' ', '\t')):
                    items[-1] += ' ' + lineas[i].strip(); i += 1
                else:
                    break
            bloques.append(('ol', items)); continue

        parrafo.append(s); i += 1

    cerrar()
    return meta, bloques

# --------------------------------------------------------------------------
# Imagenes
# --------------------------------------------------------------------------
def medir(datos):
    if datos[:8] == b'\x89PNG\r\n\x1a\n':
        return struct.unpack('>II', datos[16:24])
    if datos[:2] == b'\xff\xd8':
        i = 2
        while i < len(datos) - 9:
            if datos[i] != 0xFF:
                i += 1; continue
            marca = datos[i + 1]
            if marca in (0xD8, 0x01) or 0xD0 <= marca <= 0xD7:
                i += 2; continue
            largo = struct.unpack('>H', datos[i + 2:i + 4])[0]
            if marca in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                         0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                alto, ancho = struct.unpack('>HH', datos[i + 5:i + 9])
                return ancho, alto
            i += 2 + largo
    raise ValueError('formato de imagen no reconocido (use PNG o JPEG)')

def escalar(px_ancho, px_alto):
    ancho, alto = px_ancho * EMU_POR_PX, px_alto * EMU_POR_PX
    if ancho > ANCHO_EMU:
        alto = int(alto * ANCHO_EMU / float(ancho))
        ancho = ANCHO_EMU
    return ancho, alto

# --------------------------------------------------------------------------
# Armado del .docx
# --------------------------------------------------------------------------
def generar(ruta_md, ruta_salida=None, plantilla=PLANTILLA):
    base = os.path.dirname(os.path.abspath(ruta_md))
    with open(ruta_md, encoding='utf8') as f:
        meta, bloques = leer_markdown(f.read(), base)
    ruta_salida = ruta_salida or os.path.splitext(ruta_md)[0] + '.docx'

    tpl = Plantilla(plantilla)
    rotulo = limpiar(meta.get('rotulo', 'Reporte'))
    titulo = limpiar(meta.get('titulo', os.path.basename(ruta_md)))
    subtitulo = limpiar(meta.get('subtitulo', ''))
    encabezado = limpiar(meta.get('encabezado', titulo))
    quiere_indice = meta.get('indice', '').lower() in ('si', 'sí', 'yes', 'true', '1')

    def poner(xml, texto):
        return re.sub(r'(<w:t xml:space="preserve">)[^<]*(</w:t>)',
                      lambda m: m.group(1) + esc(texto) + m.group(2), xml, count=1)

    portada = list(tpl.portada)
    portada[19] = poner(portada[19], rotulo)
    portada[20] = poner(portada[20], titulo)
    portada[21] = poner(portada[21], subtitulo)

    partes = [''.join(portada)]
    if quiere_indice:
        partes.append(p_indice())

    imagenes, rels_extra, listas = [], [], 0
    for tipo, dato in bloques:
        if tipo == 'h2':
            partes.append(p_titulo(dato, tpl.barra))
        elif tipo == 'sub':
            partes.append(p_subtitulo(dato))
        elif tipo == 'p':
            partes.append(p_parrafo(dato))
        elif tipo == 'ul':
            partes.append(''.join(p_item(t, NUM_VINETA, 60) for t in dato))
        elif tipo == 'ol':
            partes.append(''.join(p_item(t, id_lista(listas), 80) for t in dato))
            listas += 1
        elif tipo == 'tabla':
            partes.append(p_tabla(*dato))
        elif tipo == 'salto':
            partes.append(SALTO)
        elif tipo == 'img':
            ruta, pie = dato
            entera = ruta if os.path.isabs(ruta) else os.path.join(base, ruta)
            if not os.path.exists(entera):
                raise SystemExit('No encuentro la imagen: %s' % entera)
            with open(entera, 'rb') as f:
                datos = f.read()
            n = len(imagenes) + 1
            ext = os.path.splitext(entera)[1].lower().lstrip('.') or 'png'
            nombre = 'word/media/mdimg%d.%s' % (n, ext)
            rid = 'rId%d' % (1000 + n)
            imagenes.append((nombre, datos))
            rels_extra.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
                              'officeDocument/2006/relationships/image" Target="media/mdimg%d.%s"/>'
                              % (rid, n, ext))
            partes.append(p_imagen(rid, *escalar(*medir(datos)), indice=n))
            if pie:
                partes.append(p_pie_figura(pie))

    documento = espacios(tpl.cabeza + ''.join(partes) + tpl.cola)
    _escribir(plantilla, ruta_salida, documento, encabezado, imagenes,
              rels_extra, quiere_indice, listas)
    return ruta_salida

# Un guardado de la plantilla desde Word o ONLYOFFICE puede dejar fuera algun
# xmlns que el generador si usa (paso con "pic", el de las imagenes: el .docx
# salia con XML invalido). Se vuelven a declarar aca en vez de exigir que la
# plantilla los conserve, que es algo que nadie va a recordar.
ESPACIOS = (
    ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'),
    ('a',   'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('wp',  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('r',   'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
)

def espacios(documento):
    corte = documento.index('>', documento.index('<w:document'))
    cabeza = documento[:corte]
    return cabeza + ''.join(' xmlns:%s="%s"' % (p, u) for p, u in ESPACIOS
                            if 'xmlns:%s=' % p not in cabeza) + documento[corte:]

def _escribir(plantilla, salida, documento, encabezado, imagenes, rels_extra, indice, listas):
    tmp = salida + '.tmp'
    with zipfile.ZipFile(plantilla) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            datos = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                datos = documento.encode('utf8')
            elif it.filename == 'word/header1.xml':
                t = datos.decode('utf8')
                textos = re.findall(r'<w:t xml:space="preserve">[^<]*</w:t>', t)
                if textos:
                    t = t.replace(textos[0], '<w:t xml:space="preserve">%s</w:t>' % esc(encabezado), 1)
                    for extra in textos[1:2]:
                        t = t.replace(extra, '<w:t xml:space="preserve"></w:t>', 1)
                datos = t.encode('utf8')
            elif it.filename == 'word/_rels/document.xml.rels' and rels_extra:
                datos = datos.decode('utf8').replace('</Relationships>',
                                                     ''.join(rels_extra) + '</Relationships>').encode('utf8')
            elif it.filename == 'word/settings.xml' and indice:
                t = datos.decode('utf8')
                if 'updateFields' not in t:
                    corte = t.index('>', t.index('<w:settings')) + 1
                    t = t[:corte] + '<w:updateFields w:val="true"/>' + t[corte:]
                datos = t.encode('utf8')
            elif it.filename == 'word/numbering.xml':
                extra = ''.join(num_lista(i) for i in range(listas))
                datos = datos.decode('utf8').replace(
                    '</w:numbering>', numeracion_titulos() + extra + '</w:numbering>', 1).encode('utf8')
            elif it.filename == '[Content_Types].xml':
                t = datos.decode('utf8')
                for ext, mime in (('jpeg', 'image/jpeg'), ('jpg', 'image/jpeg'), ('gif', 'image/gif')):
                    if 'Extension="%s"' % ext not in t:
                        t = t.replace('</Types>', '<Default Extension="%s" ContentType="%s"/></Types>'
                                      % (ext, mime))
                datos = t.encode('utf8')
            zout.writestr(it, datos)
        for nombre, datos in imagenes:
            zout.writestr(nombre, datos)
    shutil.move(tmp, salida)


# --------------------------------------------------------------------------
# --nuevo: arma el .md de arranque. El menu vive aca, no en la generacion,
# para que generar un .md siga dando siempre el mismo .docx.
# --------------------------------------------------------------------------
# clave, etiqueta, rotulo de portada
TIPOS = [
    ('propuesta-con-costos', 'Propuesta con costos',  'Propuesta'),
    ('propuesta-sin-costos', 'Propuesta sin costos',  'Propuesta'),
    ('analisis',             'Documento de analisis', 'Reporte'),
    ('guia',                 'Guia / informativo',    'Guia'),
    ('cotizacion',           'Cotizacion',            'Cotizacion'),
]

# archivo, etiqueta, tipos donde se ofrece, marcado por defecto.
# El orden de esta lista es el orden en que salen en el documento.
BLOQUES = [
    ('forma-de-pago-50-50.md', 'Forma de pago 50/50 por fase',
     ('propuesta-con-costos', 'cotizacion'), True),
    ('forma-de-pago-anticipo-final.md', 'Forma de pago 50/50 anticipo-final',
     ('propuesta-con-costos', 'cotizacion'), False),
    ('sem-google-ads.md', 'SEM / Google Ads',
     ('propuesta-con-costos', 'propuesta-sin-costos'), False),
    ('hosting-dedicado.md', 'Hosting dedicado (Debian)',
     ('propuesta-con-costos', 'propuesta-sin-costos'), False),
    ('hosting-dedicado-cpanel.md', 'Hosting dedicado (AlmaLinux + cPanel)',
     ('propuesta-con-costos', 'propuesta-sin-costos'), False),
    ('hosting-externo.md', 'Manejo de hosting externo',
     ('propuesta-con-costos', 'propuesta-sin-costos'), False),
    ('soporte-mensual.md', 'Plan de soporte mensual',
     ('propuesta-con-costos', 'propuesta-sin-costos'), False),
    ('portafolio-clientes.md', 'Portafolio de clientes',
     ('propuesta-con-costos', 'propuesta-sin-costos', 'cotizacion'), True),
]

def bloques_de(tipo):
    return [b for b in BLOQUES if tipo in b[2]]

def variables_de(archivos):
    """Las variables no se declaran: se leen de los propios bloques."""
    faltan = []
    for a in archivos:
        with open(os.path.join(AQUI, 'bloques', a), encoding='utf8') as f:
            for v in VARIABLE.findall(f.read()):
                if v not in faltan:
                    faltan.append(v)
    return faltan

def armar_nuevo(tipo, archivos, datos):
    """Devuelve el texto del .md. Sin entrada/salida: por eso se puede probar."""
    rotulo = dict((t[0], t[2]) for t in TIPOS)[tipo]
    with open(os.path.join(AQUI, 'esqueletos', tipo + '.md'), encoding='utf8') as f:
        cuerpo = f.read().rstrip('\n')
    campos = ['rotulo: ' + rotulo]
    campos += ['%s: %s' % (k, datos.get(k, '')) for k in ('titulo', 'subtitulo', 'encabezado')]
    campos += ['indice: si']
    campos += ['%s: %s' % (k, v) for k, v in datos.items()
               if k not in ('titulo', 'subtitulo', 'encabezado')]
    incluidos = ''.join('\n@incluir bloques/%s\n' % a for a in archivos)
    return '---\n%s\n---\n\n%s\n%s' % ('\n'.join(campos), cuerpo, incluidos)

def _elegir(titulo, opciones, multiple=False, marcados=()):
    print('\n%s' % titulo)
    for i, etiqueta in enumerate(opciones, 1):
        print('  %d) %-32s %s' % (i, etiqueta, '[x]' if i - 1 in marcados else ''))
    while True:
        crudo = input('> ' if not multiple else '> (numeros separados por espacio, Enter = los marcados) ')
        crudo = crudo.strip()
        if multiple and not crudo:
            return list(marcados)
        try:
            elegidos = [int(x) - 1 for x in crudo.split()]
        except ValueError:
            print('  Solo numeros.'); continue
        if not elegidos or any(e < 0 or e >= len(opciones) for e in elegidos):
            print('  Fuera de rango.'); continue
        if not multiple and len(elegidos) != 1:
            print('  Elegi uno solo.'); continue
        return elegidos

def menu_nuevo(tipo=None, destino=None):
    if tipo is None:
        i = _elegir('1. Tipo de documento', [t[1] for t in TIPOS])[0]
        tipo = TIPOS[i][0]
    elif tipo not in [t[0] for t in TIPOS]:
        raise SystemExit('Tipo desconocido: %s. Hay: %s'
                         % (tipo, ', '.join(t[0] for t in TIPOS)))

    disponibles = bloques_de(tipo)
    archivos = []
    if disponibles:
        marcados = [i for i, b in enumerate(disponibles) if b[3]]
        elegidos = _elegir('2. Bloques reutilizables', [b[1] for b in disponibles],
                           multiple=True, marcados=marcados)
        archivos = [disponibles[i][0] for i in sorted(elegidos)]

    print('\n3. Datos')
    datos = {}
    for campo in ['titulo', 'subtitulo', 'encabezado'] + variables_de(archivos):
        datos[campo] = input('   %s: ' % campo).strip()

    destino = destino or (re.sub(r'[^a-z0-9]+', '-',
                                 (datos.get('cliente') or datos.get('titulo') or tipo).lower()
                                 ).strip('-') + '.md')
    if os.path.exists(destino):
        raise SystemExit('Ya existe %s; borralo o pasa otro nombre.' % destino)
    with open(destino, 'w', encoding='utf8') as f:
        f.write(armar_nuevo(tipo, archivos, datos))
    print('\nEscrito: %s' % destino)
    print('Editalo y despues: python3 mindsoftdoc.py %s' % destino)
    return destino

def main():
    ap = argparse.ArgumentParser(description='Genera un .docx con el formato de la casa Mindsoft.')
    ap.add_argument('entrada', nargs='?', help='archivo markdown')
    ap.add_argument('-o', '--salida', help='ruta del .docx (por defecto, junto al .md)')
    ap.add_argument('--plantilla', default=PLANTILLA, help='otra plantilla .docx')
    ap.add_argument('--nuevo', nargs='?', const='', metavar='TIPO',
                    help='arma un .md de arranque; sin TIPO abre el menu')
    args = ap.parse_args()

    if args.nuevo is not None:
        menu_nuevo(args.nuevo or None, args.entrada)
        return
    if not args.entrada:
        ap.error('falta el archivo markdown (o usa --nuevo)')

    salida = generar(args.entrada, args.salida, args.plantilla)
    print('Generado: %s' % salida)
    print('Reglas de la casa: %s' % resumen_limpieza())
    if _faltantes:
        print('AVISO - variables sin valor, quedaron a la vista en el documento: %s'
              % resumen_variables())

if __name__ == '__main__':
    main()
