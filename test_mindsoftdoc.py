#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Verificacion de mindsoftdoc: python3 test_mindsoftdoc.py"""
import os, re, sys, zipfile, tempfile
from xml.dom.minidom import parseString

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import mindsoftdoc as md

PROHIBIDOS = '—–‒―−‘’‚‛“”„‟′″«»… '


def test_reglas_de_la_casa():
    md._conteo.clear()
    sucio = 'Un guion —largo—, comillas “dobles” y ‘simples’, puntos…'
    limpio = md.limpiar(sucio)
    assert '—' not in limpio and '–' not in limpio, limpio
    assert '“' not in limpio and '”' not in limpio and '’' not in limpio, limpio
    assert '…' not in limpio, limpio
    assert limpio == 'Un guion -largo-, comillas "dobles" y \'simples\', puntos...', limpio
    assert md._conteo['guiones largos'] == 2, md._conteo
    assert md._conteo['comillas'] == 4, md._conteo
    print('  reglas de la casa: guiones y comillas normalizados')


def test_texto_que_parece_marcado():
    # Un texto normal puede empezar con "[" o "*" sin ser enlace ni cursiva.
    for crudo, esperado in (
        ('[DEPOSIT][CONTROLLER] Request recibido', '[DEPOSIT][CONTROLLER] Request recibido'),
        ('SQLSTATE[42S22] 1054 Unknown column', 'SQLSTATE[42S22] 1054 Unknown column'),
        ('2 * 3 = 6', '2 * 3 = 6'),
        ('ver [la guia](https://mindsoft.biz)', 'ver la guia (https://mindsoft.biz)'),
        ('ver [el anexo](anexo.md)', 'ver el anexo'),
        ('un **dato** y `codigo`', 'un dato y codigo'),
    ):
        salida = md.runs(crudo)
        visible = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', salida))
        assert visible == esperado, (crudo, visible)
    assert '<w:b w:val="1"/>' in md.runs('un **dato**'), 'la negrita no se aplico'
    print('  marcado en linea: negrita, codigo y enlaces sin confundir texto normal')


def test_anchos_de_tabla():
    for cab, filas in (
        (['a', 'b'], [['x', 'y']]),
        (['Metrica', 'Valor', 'Nota muy larga que ocupa bastante'], [['1', '2', '3']]),
        (['a'] * 9, [['x'] * 9]),
    ):
        w = md.anchos(cab, filas)
        assert sum(w) == md.ANCHO_TABLA, (w, sum(w))
        assert len(w) == len(cab)
        assert all(x > 0 for x in w), w
    ancho = md.anchos(['corta', 'una columna con mucho mas texto adentro'], [['a', 'b']])
    assert ancho[1] > ancho[0], ancho
    print('  anchos de tabla: suman %d y la columna con mas texto es mas ancha' % md.ANCHO_TABLA)


def test_medir_imagen():
    ancho, alto = md.medir(open(os.path.join(AQUI, 'ejemplo-captura.png'), 'rb').read())
    assert (ancho, alto) == (600, 180), (ancho, alto)
    a, b = md.escalar(200, 100)            # entra sin reducir
    assert (a, b) == (200 * 9525, 100 * 9525), (a, b)
    a, b = md.escalar(600, 180)            # mas ancha que la pagina: se reduce
    assert a == md.ANCHO_EMU, a
    assert abs(b / float(a) - 180 / 600.0) < 1e-4, 'se deformo la proporcion'
    print('  imagenes: PNG medido en 600x180, escalado al ancho util sin deformar')


def test_lectura_markdown():
    meta, bloques = md.leer_markdown(
        '---\ntitulo: T\nrotulo: Manual\n---\n\n## Uno\n\nParrafo.\n\n- a\n- b\n\n'
        '1. x\n2. y\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n---\n\n### Sub\n')
    assert meta['titulo'] == 'T' and meta['rotulo'] == 'Manual', meta
    tipos = [t for t, _ in bloques]
    assert tipos == ['h2', 'p', 'ul', 'ol', 'tabla', 'salto', 'sub'], tipos
    assert bloques[2][1] == ['a', 'b'], bloques[2]
    assert bloques[4][1] == (['A', 'B'], [['1', '2']]), bloques[4]
    print('  markdown: titulos, listas, tabla y salto reconocidos')


def test_documento_generado():
    salida = os.path.join(tempfile.mkdtemp(), 'ejemplo.docx')
    md._conteo.clear()
    md.generar(os.path.join(AQUI, 'ejemplo.md'), salida)

    z = zipfile.ZipFile(salida)
    assert z.testzip() is None, 'el .docx esta corrupto'
    doc = z.read('word/document.xml').decode('utf8')
    parseString(doc)                                    # revienta si no es XML valido

    visible = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', doc))
    malos = sorted({c for c in visible if c in PROHIBIDOS})
    assert not malos, 'sobrevivieron caracteres prohibidos: %r' % malos
    assert md._conteo, 'el ejemplo traia caracteres prohibidos y no se contaron'

    assert doc.count('w:pStyle w:val="Heading2"') == 7, doc.count('w:pStyle w:val="Heading2"')
    assert doc.count('<w:tbl>') == 1, doc.count('<w:tbl>')
    assert doc.count('w:fill="001689"') == 3, 'cabecera azul en las 3 columnas'
    assert doc.count('w:fill="EEF1F8"') == 0, 'las filas no deben llevar sombreado alternado'
    assert doc.count('w:fill="FFFFFF"') == 9, 'las 3 filas de datos van en blanco'
    assert doc.count('w:numId w:val="90"') == 3, 'vinetas'
    assert doc.count('w:numId w:val="91"') == 3, 'numeradas'
    assert doc.count('<w:br w:type="page"/>') == 3, 'saltos de pagina'
    assert 'TOC \\o' in doc, 'falta el campo de indice'
    assert 'word/media/mdimg1.png' in z.namelist(), 'falta la imagen embebida'
    assert 'rId1001' in doc and 'rId1001' in z.read('word/_rels/document.xml.rels').decode('utf8')
    assert 'updateFields' in z.read('word/settings.xml').decode('utf8')

    cab = z.read('word/header1.xml').decode('utf8')
    assert 'Documento de ejemplo | Septiembre 2026' in cab, 'encabezado no aplicado'
    assert 'Siegfried' not in cab, 'quedo el encabezado de la plantilla'
    assert 'Documento de ejemplo - formato Mindsoft' in visible, 'falta el titulo en la portada'

    t = zipfile.ZipFile(md.PLANTILLA)
    intocables = [n for n in t.namelist()
                  if n not in ('word/document.xml', 'word/header1.xml',
                               'word/settings.xml', '[Content_Types].xml',
                               'word/_rels/document.xml.rels')]
    distintas = [n for n in intocables if t.read(n) != z.read(n)]
    assert not distintas, 'se altero el formato de la plantilla en: %s' % distintas
    assert len(intocables) == 24, len(intocables)
    print('  documento: XML valido, %d partes del formato intactas, %s corregidos'
          % (len(intocables), md.resumen_limpieza()))
    return salida


if __name__ == '__main__':
    print('Verificando mindsoftdoc...')
    test_reglas_de_la_casa()
    test_texto_que_parece_marcado()
    test_anchos_de_tabla()
    test_medir_imagen()
    test_lectura_markdown()
    ruta = test_documento_generado()
    print('TODO OK -> %s' % ruta)
