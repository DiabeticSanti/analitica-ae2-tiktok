# -*- coding: utf-8 -*-
"""Actualiza el Word de AE2 con las secciones de AE3.

Lee los archivos que exporta el notebook (datos/tablas, datos/calidad, datos/preparado, resultados)
y agrega, antes de «Referencias», las secciones de metodología, diccionario, resultados y conclusiones.
Así las cifras del documento coinciden con las del notebook.

Uso:  python generar_documento.py      (desde la carpeta AE3/)
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

FONT = "Arial"
BASE_CANDIDATAS = [
    "../AE2/documento/Documento_AE2_Necesidad_y_Diseno.docx",
    "../documento/Documento_AE2_Necesidad_y_Diseno.docx",
    "../Documento_AE2_Necesidad_y_Diseno.docx",
    "documento/Documento_AE2_Necesidad_y_Diseno.docx",
]
SALIDA = "documento/Documento_AE3_Avance.docx"

# ----------------------------------------------------------------- formato de números
def _nan(x):
    return x is None or (isinstance(x, float) and x != x) or pd.isna(x)
def fint(x):
    return "—" if _nan(x) else f"{int(round(float(x))):,}".replace(",", ".")
def fdec(x, d=2):
    return "—" if _nan(x) else f"{float(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ----------------------------------------------------------------- utilidades de python-docx
def _fuente(run, size=11, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    rpr = run._r.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(a), FONT)

def _estilo(p, style_id):
    ppr = p._p.get_or_add_pPr()
    ps = ppr.find(qn("w:pStyle"))
    if ps is None:
        ps = OxmlElement("w:pStyle")
        ppr.insert(0, ps)
    ps.set(qn("w:val"), style_id)

def _campo_seq(par, numero):
    """Agrega un campo SEQ Tabla (con el número ya calculado) como campo complejo, igual que el resto del documento."""
    def run_nuevo():
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        for a in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(a), FONT)
        rpr.append(rf)
        rpr.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "20"); rpr.append(sz)
        r.append(rpr)
        par._p.append(r)
        return r
    for tipo in ("begin", "instr", "separate", "texto", "end"):
        r = run_nuevo()
        if tipo == "instr":
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = " SEQ Tabla \\* ARABIC "
            r.append(it)
        elif tipo == "texto":
            t = OxmlElement("w:t"); t.text = str(numero); r.append(t)
        else:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), tipo); r.append(fc)

class Insertor:
    """Inserta contenido justo antes de un párrafo ancla."""
    def __init__(self, doc, ancla):
        self.doc, self.ancla = doc, ancla
        self.n_tabla = len(re.findall(r"SEQ Tabla", doc.element.xml))

    def h1(self, texto, salto=False):
        p = self.ancla.insert_paragraph_before()
        _estilo(p, "Heading1")
        if salto:
            p.paragraph_format.page_break_before = True
        _fuente(p.add_run(texto), size=15, bold=True)
        return p

    def h2(self, texto):
        p = self.ancla.insert_paragraph_before()
        _estilo(p, "Heading2")
        p.paragraph_format.keep_with_next = True
        _fuente(p.add_run(texto), size=12.5, bold=True)
        return p

    def p(self, texto, negrita_inicio=None, italica=False, size=11, justificar=True):
        par = self.ancla.insert_paragraph_before()
        par.paragraph_format.space_after = Pt(7)
        par.paragraph_format.line_spacing = 1.25
        if justificar:
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if negrita_inicio:
            _fuente(par.add_run(negrita_inicio), size=size, bold=True)
        _fuente(par.add_run(texto), size=size, italic=italica)
        return par

    def vineta(self, texto, negrita_inicio=None):
        par = self.ancla.insert_paragraph_before()
        par.paragraph_format.left_indent = Twips(720)
        par.paragraph_format.first_line_indent = Twips(-360)
        par.paragraph_format.space_after = Pt(4)
        par.paragraph_format.line_spacing = 1.2
        _fuente(par.add_run("•\t"), size=11)
        if negrita_inicio:
            _fuente(par.add_run(negrita_inicio), size=11, bold=True)
        _fuente(par.add_run(texto), size=11)
        return par

    def fuente(self, texto):
        par = self.ancla.insert_paragraph_before()
        par.paragraph_format.space_before = Pt(3)
        par.paragraph_format.space_after = Pt(10)
        _fuente(par.add_run(texto), size=8.5, italic=True)

    def tabla(self, titulo, encabezado, filas, anchos, alinear=None, fuente_txt=None, size=8):
        # --- leyenda numerada (campo SEQ para el índice de tablas)
        self.n_tabla += 1
        cap = self.ancla.insert_paragraph_before()
        cap.paragraph_format.keep_with_next = True
        cap.paragraph_format.space_before = Pt(8)
        cap.paragraph_format.space_after = Pt(4)
        _fuente(cap.add_run("Tabla "), size=10, bold=True)
        _campo_seq(cap, self.n_tabla)
        _fuente(cap.add_run(". " + titulo), size=10, bold=True)

        # --- tabla
        alinear = alinear or ["l"] * len(encabezado)
        tbl = self.doc.add_table(rows=1 + len(filas), cols=len(encabezado))
        tbl.autofit = False
        tblpr = tbl._tbl.tblPr
        bordes = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{lado}")
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:color"), "999999")
            bordes.append(e)
        mar = OxmlElement("w:tblCellMar")
        for lado, v in (("top", 40), ("left", 70), ("bottom", 40), ("right", 70)):
            e = OxmlElement(f"w:{lado}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); mar.append(e)
        # orden del esquema OOXML: tblW, tblBorders, tblLayout, tblCellMar, tblLook
        layout = tblpr.find(qn("w:tblLayout"))
        look = tblpr.find(qn("w:tblLook"))
        if layout is not None:
            layout.addprevious(bordes)
        elif look is not None:
            look.addprevious(bordes)
        else:
            tblpr.append(bordes)
        if look is not None:
            look.addprevious(mar)
        else:
            tblpr.append(mar)

        juntar = len(filas) <= 12                           # tablas cortas: no partirlas entre páginas

        def celda(c, texto, w, enc, al):
            c.width = Twips(w)
            if enc:
                tcpr = c._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "1F3864")
                tcpr.append(shd)
            par = c.paragraphs[0]
            par.paragraph_format.space_after = Pt(0)
            if juntar:
                par.paragraph_format.keep_with_next = True
            par.alignment = {"l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT,
                             "c": WD_ALIGN_PARAGRAPH.CENTER}[al if not enc else "l"]
            run = par.add_run(str(texto))
            _fuente(run, size=size, bold=enc)
            if enc:
                from docx.shared import RGBColor
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        for j, w in enumerate(anchos):
            tbl.columns[j].width = Twips(w)                 # fija la cuadrícula (gridCol)
        for j, h in enumerate(encabezado):
            celda(tbl.rows[0].cells[j], h, anchos[j], True, "l")
        trpr = tbl.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader"); trpr.append(th)
        for i, fila in enumerate(filas, start=1):
            cs = OxmlElement("w:cantSplit"); tbl.rows[i]._tr.get_or_add_trPr().append(cs)
            for j, v in enumerate(fila):
                celda(tbl.rows[i].cells[j], v, anchos[j], False, alinear[j])
        self.ancla._p.addprevious(tbl._tbl)
        self.fuente(fuente_txt or "Fuente: elaboración propia.")

def _reemplazar(doc, viejo, nuevo):
    """Reemplaza texto en párrafos y celdas; devuelve cuántas veces lo encontró."""
    n = 0
    pars = list(doc.paragraphs)
    for t in doc.tables:
        for fila in t.rows:
            for c in fila.cells:
                pars.extend(c.paragraphs)
    for p in pars:
        if viejo not in p.text:
            continue
        hecho = False
        for r in p.runs:
            if viejo in r.text:
                r.text = r.text.replace(viejo, nuevo); hecho = True
        if not hecho and p.runs:                       # texto repartido entre runs
            texto = p.text.replace(viejo, nuevo)
            p.runs[0].text = texto
            for r in p.runs[1:]:
                r.text = ""
        n += 1
    return n

def _reparar_campos_anidados(doc):
    """Repara runs anidados dentro de campos SEQ (XML inválido que dejó una versión previa del Word de AE2)."""
    n = 0
    for pr in list(doc.element.body.iter(qn("w:r"))):
        anidados = [h for h in pr if h.tag == qn("w:r")]
        if not anidados:
            continue
        nested = anidados[0]
        padre = pr.getparent()
        pos = padre.index(pr)
        tras = OxmlElement("w:r")
        visto = False
        for h in list(pr):
            if h is nested:
                visto = True
                continue
            if visto:
                tras.append(h)
        pr.remove(nested)
        padre.insert(pos + 1, nested)
        if len(tras):
            padre.insert(pos + 2, tras)
        n += 1
    return n

def _fijar_texto(p, texto):
    if p.runs:
        p.runs[0].text = texto
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(texto)

# ----------------------------------------------------------------- contenido
PREGUNTAS = [
    ("P1", "¿Cómo se distribuyen las vistas y la interacción en la muestra?",
     "views, interacciones_totales, tasa_interaccion, rango_vistas", "p1a_resumen_general; p1b_rango_vistas",
     "Mediana, cuartiles y concentración de vistas"),
    ("P2", "¿Cómo varían según el tipo de publicación (video o foto)?",
     "type, views, tasa_interaccion", "p2_tipo_publicacion", "Mediana de vistas e interacción por tipo"),
    ("P3", "¿Cómo varían según el rango de duración?",
     "duration_ms → duracion_s → rango_duracion", "p3_rango_duracion", "Mediana de vistas e interacción por rango"),
    ("P4", "¿Cómo varían según el país?",
     "country → pais_agrupado", "p4_pais; p4b_pais_completo", "Comparación entre los países con más publicaciones"),
    ("P5", "¿Cómo varían según el carácter comercial (publicidad, TikTok Shop, contenido de marca)?",
     "is_ad, is_shop_video, branded_content → categoria_comercial", "p5a_categoria_comercial; p5b_indicadores_comerciales",
     "Mediana de vistas e interacción por categoría"),
    ("P6", "¿Cómo cambian según el año de publicación? (pregunta de contexto)",
     "posted_at → anio_publicacion, edad_dias, vistas_por_dia", "p6_anio_publicacion", "Efecto del tiempo en línea sobre las métricas"),
]

AJUSTES = [
    ("A1. Desglose en P1–P6. ", "La pregunta de AE2 combina cuatro dimensiones y no cabe en una sola tabla; se separa por dimensión y se añade una línea base (P1)."),
    ("A2. P6 es una pregunta nueva. ", "AE2 advirtió que cada registro tiene métricas medidas en fechas distintas; sin el año de publicación y el tiempo en línea, las comparaciones de P2 a P5 podrían confundirse con la antigüedad."),
    ("A3. País agrupado. ", "Se comparan los países con más publicaciones; el resto se agrupa en «Otros países» y los vacíos en «SIN DATO», para que los totales cuadren. La tabla completa se exporta aparte."),
    ("A4. Duración solo en videos. ", "Según la ficha, las fotos tienen duración 0; se clasifican aparte y no entran en los rankings de P3."),
    ("A5. Tasa agregada. ", "Además de la mediana de la tasa por publicación se calcula la tasa agregada (suma de interacciones sobre suma de vistas) para no depender de publicaciones con pocas vistas."),
    ("A6. Alcance de variables. ", "De las columnas cargadas en AE2 se conservan las que responden las preguntas. No se incorpora ninguna fuente externa, porque las variables existentes bastan para la necesidad."),
]

def _leer(ruta, **kw):
    return pd.read_csv(ruta, encoding="utf-8", **kw)

def _filas_grupo(t, extra=None, prefijo=None):
    filas = []
    for _, r in t.iterrows():
        g = str(r["grupo"]) if prefijo is None else f"{r[prefijo]}: {r['grupo']}"
        fila = [g, fint(r["n_publicaciones"]), fdec(r["pct_publicaciones"]), fint(r["vistas_mediana"]),
                f"{fint(r['vistas_p25'])} – {fint(r['vistas_p75'])}", fdec(r["tasa_mediana_pct"]),
                fdec(r["tasa_agregada_pct"]), fdec(r["pct_vistas"])]
        if extra:
            fila = fila[:-1] + [fint(r[extra[0]]), fdec(r[extra[1]], 1), fila[-1]]
        filas.append(fila)
    return filas

ENC_GRUPO = ["Grupo", "Publ. (n)", "% publ.", "Vistas (mediana)", "Vistas P25 – P75",
             "Tasa mediana (%)", "Tasa agregada (%)", "% de vistas"]
AN_GRUPO = [2200, 1000, 800, 1000, 1360, 900, 900, 1200]
AL_GRUPO = ["l", "r", "r", "r", "r", "r", "r", "r"]

def main(base=None, salida=SALIDA):
    raiz = Path(".")
    if base is None:
        base = next((c for c in BASE_CANDIDATAS if Path(c).exists()), None)
    if base is None or not Path(base).exists():
        raise FileNotFoundError("No se encontró el Word de AE2. Indique su ruta: generar_documento.main(base='ruta.docx')")

    meta = json.loads(Path("resultados/metadatos_ejecucion.json").read_text(encoding="utf-8"))
    interp = json.loads(Path("resultados/interpretaciones.json").read_text(encoding="utf-8"))
    bit = _leer("datos/calidad/bitacora_transformaciones.csv")
    ctr = _leer("datos/calidad/controles_calidad.csv")
    val = _leer("datos/calidad/validaciones_enriquecimiento.csv")
    cmp_res = _leer("datos/calidad/comparacion_resumen.csv")
    der = _leer("datos/calidad/variables_derivadas.csv")
    ver = _leer("datos/calidad/verificaciones_consistencia.csv")
    dic = _leer("datos/preparado/diccionario_actualizado.csv")
    T = {n: _leer(f"datos/tablas/{n}.csv", dtype={"grupo": str, "indicador": str})
         for n in meta["tablas"]}

    doc = Document(base)
    reparados = _reparar_campos_anidados(doc)
    if reparados:
        print(f"Se repararon {reparados} campos de numeración de tablas del Word de AE2.")

    # ---------- 1) Completar marcadores pendientes de AE2 con datos de la ejecución
    rev = meta.get("revision") or "no disponible"
    _reemplazar(doc, "[Completar con el valor que imprime el notebook]", str(rev))
    _reemplazar(doc, "[Completar con los valores obtenidos.]",
                f"En la ejecución de AE3 la muestra tuvo {fint(meta['n_filas_cargadas'])} filas "
                f"({fdec(100 * meta['n_filas_cargadas'] / meta['total_filas_dataset'], 4)} % de las "
                f"{fint(meta['total_filas_dataset'])} del dataset) y cubrió publicaciones de {meta['anio_min']} a {meta['anio_max']}.")
    _reemplazar(doc, "[Completar con las dimensiones, los tipos, los nulos y el periodo de la muestra obtenidos al ejecutar el notebook.]",
                f"se cargaron {fint(meta['n_filas_cargadas'])} filas y {meta['n_cols_cargadas']} columnas, con "
                f"{fint(meta['n_nulos_cargados'])} valores nulos en total; el periodo de publicación va de "
                f"{meta['periodo_min'][:10]} a {meta['periodo_max'][:10]}. El detalle de tipos y nulos por columna está en "
                "datos/calidad/comparacion_columnas.csv.")

    # ---------- 2) Portada
    for p in doc.paragraphs[:25]:
        if p.text.startswith("Avance AE2"):
            _fijar_texto(p, "Avances AE2 y AE3 – Necesidad, diseño, transformación y enriquecimiento de datos")
            break

    # ---------- 3) Ancla: «Referencias»
    ancla = None
    for p in doc.paragraphs:
        if p.text.strip() == "Referencias":
            ancla = p
    if ancla is None:                                   # sin sección de referencias: se agrega al final
        ancla = doc.add_paragraph("Referencias")
        _estilo(ancla, "Heading1")
    ins = Insertor(doc, ancla)
    b0 = ins.n_tabla                                   # tablas que ya existen (AE2)
    def nt(k):
        return f"Tabla {b0 + k}"

    # =====================================================================
    # 8. Metodología de la etapa 2
    # =====================================================================
    ins.h1("8. Metodología de la etapa 2 (transformación y enriquecimiento)", salto=True)
    ins.p("Esta etapa prepara la muestra de AE2 para responder descriptivamente las preguntas de negocio y deja los resultados "
          "organizados para su presentación gráfica en AE4. Todo el procedimiento está en el notebook "
          "AE3_Transformacion_Enriquecimiento.ipynb, que se reproduce con rutas relativas; este documento se actualizó a partir de los "
          "archivos que el notebook exporta, por lo que las cifras coinciden con él.")

    ins.h2("8.1 Retoma de la necesidad y ajustes respecto a AE2")
    ins.p("La necesidad sigue siendo la planteada en la sección 2: un escenario hipotético en el que el equipo de contenidos de una "
          "agencia necesita conocer cómo se comportan el alcance y la interacción según el formato y el carácter comercial de las "
          "publicaciones. Se mantiene la pregunta de AE2 y se desglosa en las preguntas de la " + nt(1) + ", cada una con las variables que "
          "necesita y la tabla que la responde.")
    ins.tabla("Preguntas de negocio, variables y resultados esperados",
              ["N.º", "Pregunta", "Variables", "Tabla de salida (datos/tablas)", "Resultado esperado"],
              [list(x) for x in PREGUNTAS], [500, 2500, 2100, 1900, 2360],
              fuente_txt="Fuente: elaboración propia a partir de la pregunta de AE2 y del diseño inicial (Tabla 3).")
    ins.p("Los ajustes respecto a AE2 y su justificación son los siguientes:")
    for neg, txt in AJUSTES:
        ins.vineta(txt, negrita_inicio=neg)

    ins.h2("8.2 Datos de partida")
    ins.p(f"Se cargó la misma muestra de AE2 mediante la clase HFDatasetLoader: dataset {meta['dataset_id']}, split {meta['split']}, "
          f"semilla {meta['semilla']} y hasta {fint(meta['n_filas_pedidas'])} filas (revisión {rev}). Resultaron "
          f"{fint(meta['n_filas_cargadas'])} filas y {meta['n_cols_cargadas']} columnas. {meta['nota_muestra']} "
          f"El periodo de publicación cubierto va de {meta['periodo_min'][:10]} a {meta['periodo_max'][:10]}. "
          "El notebook conserva una copia sin modificar (df_original) para comparar los cambios.")

    ins.h2("8.3 Transformaciones y reglas de calidad")
    ins.p("Cada transformación se registró en una bitácora con el número de registros afectados y su justificación (" + nt(2) + "). "
          "Ningún registro se elimina salvo los duplicados exactos por video_id, si existieran; los registros dudosos se marcan "
          "con indicadores f_* y se excluyen únicamente de los cálculos que dependen del campo dudoso.")
    ins.tabla("Bitácora de transformaciones",
              ["Paso", "Operación", "Registros afectados", "Justificación"],
              [[r["paso"], r["operacion"], fint(r["registros_afectados"]), r["justificacion"]] for _, r in bit.iterrows()],
              [600, 1800, 1100, 5860], ["l", "l", "r", "l"],
              fuente_txt="Fuente: datos/calidad/bitacora_transformaciones.csv, generado por el notebook.")
    ins.p("Los controles de calidad se informan aunque no detecten problemas (" + nt(3) + ").")
    ins.tabla("Controles de calidad",
              ["Control", "Criterio", "Detectados", "%", "Tratamiento"],
              [[r["control"], r["criterio"], fint(r["registros_detectados"]), fdec(r["porcentaje"], 2), r["tratamiento"]]
               for _, r in ctr.iterrows()],
              [1700, 2700, 1000, 800, 3160], ["l", "l", "r", "r", "l"],
              fuente_txt="Fuente: datos/calidad/controles_calidad.csv, generado por el notebook.")
    ins.p(f"Como resultado, {fint(meta['n_analisis_vistas'])} de {fint(meta['n_filas_final'])} registros "
          f"({fdec(100 * meta['n_analisis_vistas'] / meta['n_filas_final'], 2)} %) entran en los cálculos de vistas y "
          f"{fint(meta['n_analisis_tasa'])} ({fdec(100 * meta['n_analisis_tasa'] / meta['n_filas_final'], 2)} %) en los de tasa de "
          f"interacción. Se completaron {fint(meta['n_tasas_imputadas'])} tasas nulas con la fórmula de la ficha, marcadas en tasa_imputada.")
    ins.p("Verificaciones del enriquecimiento:", justificar=False)
    for _, r in val.iterrows():
        ins.vineta(f"{r['detalle']} — resultado: {r['resultado']}.", negrita_inicio=f"{r['verificacion']}: ")

    ins.h2("8.4 Comparación antes y después")
    ins.p(nt(4) + " compara la muestra cargada con el DataFrame final. El detalle por columna (tipo y nulos antes y después, y si "
          "la columna se conservó, se derivó o se descartó) está en datos/calidad/comparacion_columnas.csv.")
    ins.tabla("Comparación antes y después de la transformación",
              ["Indicador", "Antes", "Después"],
              [[r["indicador"], fint(r["antes"]), fint(r["despues"])] for _, r in cmp_res.iterrows()],
              [4000, 2680, 2680], ["l", "r", "r"],
              fuente_txt="Fuente: datos/calidad/comparacion_resumen.csv, generado por el notebook.")
    ins.p("El total de valores nulos cambia porque se descartan columnas y porque algunas variables derivadas quedan vacías por definición "
          "(por ejemplo, duracion_s en fotos o edad_dias cuando la fecha es incoherente); esos vacíos no son pérdida de datos.")

    ins.h2("8.5 Enriquecimiento")
    ins.p("Se incorporaron variables calculadas con los datos existentes, cada una con su regla, significado y utilidad para las preguntas (" + nt(5) + ").")
    ins.tabla("Variables derivadas",
              ["Variable", "Regla", "Significado", "Utilidad"],
              [[r["variable"], r["regla"], r["significado"], r["utilidad"]] for _, r in der.iterrows()],
              [1700, 2700, 2400, 2560],
              fuente_txt="Fuente: datos/calidad/variables_derivadas.csv, generado por el notebook.")

    ins.h2("8.6 Procedimiento para responder cada pregunta")
    ins.vineta("Los registros se filtran con analisis_vistas (para vistas) o analisis_tasa (para interacción), de modo que todas las tablas usan la misma base.")
    ins.vineta("Para cada dimensión se agrupa y se calcula: número y porcentaje de publicaciones, vistas totales y porcentaje de vistas, mediana y cuartiles de vistas, mediana de la tasa de interacción y tasa agregada.")
    ins.vineta("Se priorizan medianas y cuartiles porque las métricas están concentradas en pocas publicaciones; la tasa agregada complementa la mediana.")
    ins.vineta(f"Las comparaciones de la interpretación solo incluyen grupos con {meta['min_n_interp']} o más registros analizables.")
    ins.vineta("Cada tabla se exporta a .csv y se verifica que la suma de publicaciones, de porcentajes y de vistas coincida con la base de análisis; luego se relee el .csv para comprobar que coincide con lo calculado.")

    # =====================================================================
    # 9. Diccionario actualizado
    # =====================================================================
    ins.h1("9. Diccionario de datos actualizado", salto=True)
    ins.p(nt(6) + " documenta todas las variables del dataset preparado (datos/preparado/tiktok_preparado.csv): las originales que se "
          "conservan, las derivadas y las marcas de calidad. Los tipos corresponden al DataFrame final del notebook.")
    ins.tabla("Diccionario actualizado de variables",
              ["Variable", "Origen", "Significado", "Tipo", "Unidad / regla"],
              [[r["variable"], r["origen"], r["significado"], r["tipo"], r["unidad_o_regla"]] for _, r in dic.iterrows()],
              [1700, 900, 3300, 1000, 2460],
              fuente_txt="Fuente: datos/preparado/diccionario_actualizado.csv, generado por el notebook.")

    # =====================================================================
    # 10. Resultados y discusión
    # =====================================================================
    ins.h1("10. Resultados y discusión", salto=True)
    ins.p("Las tablas siguientes responden cada pregunta de negocio. Las vistas se resumen con medianas y cuartiles (P25 y P75); la tasa "
          "de interacción se expresa en %, tanto como mediana por publicación como en forma agregada (suma de interacciones sobre suma "
          "de vistas del grupo). Las diferencias son descriptivas: no demuestran que una característica cause un mejor resultado.")

    # P1
    ins.h2("10.1 P1. " + interp["P1"]["titulo"])
    t = T["p1a_resumen_general"]
    ins.tabla("Resumen general de alcance e interacción",
              ["Variable", "Unidad", "n", "Media", "Mediana", "P25", "P75", "P95", "Máximo"],
              [[r["variable"], r["unidad"], fint(r["n"]), fdec(r["media"], 1), fdec(r["mediana"], 1), fdec(r["p25"], 1),
                fdec(r["p75"], 1), fdec(r["p95"], 1), fdec(r["maximo"], 1)] for _, r in t.iterrows()],
              [1700, 1500, 900, 900, 900, 800, 800, 900, 960], ["l", "l", "r", "r", "r", "r", "r", "r", "r"],
              fuente_txt="Fuente: datos/tablas/p1a_resumen_general.csv.")
    ins.tabla("Distribución de publicaciones y vistas por rango de vistas", ENC_GRUPO,
              _filas_grupo(T["p1b_rango_vistas"]), AN_GRUPO, AL_GRUPO,
              fuente_txt="Fuente: datos/tablas/p1b_rango_vistas.csv.")
    ins.p(interp["P1"]["texto"], negrita_inicio="Interpretación. ")
    ins.p(interp["P1"]["limitacion"], negrita_inicio="Limitación. ")

    # P2–P5
    bloques = [
        ("10.2", "P2", "p2_tipo_publicacion", "Publicaciones, vistas e interacción por tipo de publicación"),
        ("10.3", "P3", "p3_rango_duracion", "Publicaciones, vistas e interacción por rango de duración"),
        ("10.4", "P4", "p4_pais", "Publicaciones, vistas e interacción por país"),
        ("10.5", "P5", "p5a_categoria_comercial", "Publicaciones, vistas e interacción por categoría comercial"),
    ]
    for num, pid, nombre, titulo in bloques:
        ins.h2(f"{num} {pid}. " + interp[pid]["titulo"])
        ins.tabla(titulo, ENC_GRUPO, _filas_grupo(T[nombre]), AN_GRUPO, AL_GRUPO, fuente_txt=f"Fuente: datos/tablas/{nombre}.csv.")
        if pid == "P5":
            ins.tabla("Comparación por indicador comercial (sí / no)", ENC_GRUPO,
                      _filas_grupo(T["p5b_indicadores_comerciales"], prefijo="indicador"), AN_GRUPO, AL_GRUPO,
                      fuente_txt="Fuente: datos/tablas/p5b_indicadores_comerciales.csv.")
        ins.p(interp[pid]["texto"], negrita_inicio="Interpretación. ")
        ins.p(interp[pid]["limitacion"], negrita_inicio="Limitación. ")

    # P6
    ins.h2("10.6 P6. " + interp["P6"]["titulo"])
    enc6 = ENC_GRUPO[:-1] + ["Edad mediana (días)", "Vistas/día (mediana)", ENC_GRUPO[-1]]
    ins.tabla("Publicaciones, vistas e interacción por año de publicación", ["Año"] + enc6[1:],
              _filas_grupo(T["p6_anio_publicacion"], extra=("edad_dias_mediana", "vistas_por_dia_mediana")),
              [1000, 1000, 760, 1000, 1360, 800, 800, 900, 900, 840], ["l"] + ["r"] * 9,
              fuente_txt="Fuente: datos/tablas/p6_anio_publicacion.csv.")
    ins.p(interp["P6"]["texto"], negrita_inicio="Interpretación. ")
    ins.p(interp["P6"]["limitacion"], negrita_inicio="Limitación. ")

    # Verificación y archivos
    ins.h2("10.7 Verificación de consistencia y archivos exportados")
    ok = int((ver["resultado"] == "OK").sum())
    ins.p(f"El notebook ejecutó {len(ver)} verificaciones de consistencia (suma de publicaciones, de porcentajes y de vistas de cada tabla "
          f"frente a la base de análisis, y registros finales frente a cargados); {ok} resultaron OK. Además, relee los .csv exportados y "
          "comprueba que coinciden con los datos y las tablas calculadas. Los archivos se listan en la " + nt(15) + ".")
    ins.tabla("Archivos exportados por el notebook",
              ["Ruta (relativa a AE3/)", "Contenido"],
              [["datos/preparado/tiktok_preparado.csv", f"Dataset preparado: {fint(meta['n_filas_final'])} filas y {meta['n_cols_final']} columnas"],
               ["datos/preparado/diccionario_actualizado.csv", "Diccionario de datos actualizado"],
               ["datos/tablas/p1a … p6 (9 archivos)", "Una tabla por pregunta de negocio (listas para graficar en AE4)"],
               ["datos/calidad/ (7 archivos)", "Bitácora, controles, validaciones, comparaciones, variables derivadas y verificaciones"],
               ["resultados/resumen_dataframe.txt", "Resumen descriptivo del DataFrame final, generado con DatasetSummarizer"],
               ["resultados/metadatos_ejecucion.json; interpretaciones.json", "Parámetros de la ejecución y textos interpretativos"]],
              [3800, 5560], fuente_txt="Fuente: elaboración propia.")

    # =====================================================================
    # 11. Conclusiones
    # =====================================================================
    ins.h1("11. Conclusiones del avance", salto=True)
    ins.p("Lo que permite conocer el análisis respecto a la necesidad:", justificar=False)
    for pid in ["P1", "P2", "P3", "P4", "P5", "P6"]:
        h = [x for x in interp[pid]["hallazgos"] if x]
        ins.vineta(" ".join(h), negrita_inicio=f"{pid}. {interp[pid]['titulo']}: ")
    ins.p("Con esto, el equipo de la agencia hipotética cuenta con una descripción ordenada de cómo se distribuyen el alcance y la "
          "interacción en la muestra y de qué grupos concentran más vistas o mayor tasa de interacción, y puede formular hipótesis para "
          "etapas posteriores.")
    ins.p("Aspectos que quedan sin resolver:", justificar=False)
    ins.vineta("Causalidad: las tablas describen diferencias entre grupos, no demuestran que el formato, la duración o el carácter comercial produzcan más o menos interacción.")
    ins.vineta("Rentabilidad: el dataset no incluye costos, inversión ni conversiones; no puede afirmarse qué formato comercial es más rentable.")
    ins.vineta("Representatividad: la muestra proviene de un archivo concreto del dataset y no es una muestra aleatoria simple; los resultados no se extrapolan a toda la plataforma.")
    ins.vineta(f"Antigüedad: las métricas se midieron en fechas distintas; P6 contextualiza el efecto del tiempo en línea, pero no lo elimina.")
    ins.vineta(f"País: {fint(meta['n_paises_sin_dato'])} publicaciones no tienen país asignado y no pueden compararse.")
    ins.p("Próximo paso (AE4): graficar las tablas de datos/tablas/ (barras por grupo, distribución por rango de vistas y evolución por "
          "año) a partir del dataset preparado.")

    # ---------- 4) Actualizar campos (índices) al abrir en Word
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true"); settings.append(uf)

    Path(salida).parent.mkdir(parents=True, exist_ok=True)
    doc.save(salida)
    return Path(salida)

if __name__ == "__main__":
    print("Documento generado:", main())
