#!/usr/bin/env python3
"""
Word (.docx) -> PDF con el modelo de los manuales "SPI Swissôtel Lima".

Lee el .docx (python-docx), reconoce el cuadro de cabecera (N°, Página,
Área, Campo, Norma, fechas...) y el cuadro de contenido (OBJETIVO,
RESPONSABLES, PASOS A SEGUIR, viñetas, numerados...), y los vuelve a dibujar
en un PDF A4 (reportlab) siempre con la misma letra (Times, como el PDF
original) aunque el Word venga con otra.

Es el equivalente de word_a_pdf.js (la versión para navegador); ambos deben
producir el mismo diseño.

Uso:
    python3 word_a_pdf.py archivo.docx [-o salida.pdf]
    python3 word_a_pdf.py carpeta_con_docx/
"""

import argparse
import io
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as rl_canvas

# --- Diseño de página (puntos PostScript; 1 pt = 20 twips) ----------
PDF_PAGE_W, PDF_PAGE_H = A4            # 595.28 x 841.89
PDF_MARGIN = 35                         # 700 twips
PDF_BOX_W = 495                         # 9900 twips: ancho de los cuadros
PDF_HDR_LEFT_W = 340                    # 6800 twips
PDF_BOX_X = (PDF_PAGE_W - PDF_BOX_W) / 2
PDF_PAD = 6                             # relleno interior de las celdas
PDF_LINE_GAP = 1.25                     # interlineado (x tamaño de letra)
PDF_PARA_GAP = 5                        # 100 twips de "after" en el Word
PDF_INDENT_LIST = 18                    # 360 twips
PDF_INDENT_LETTER = 25                  # 500 twips
FONT_REG, FONT_BOLD = "Times-Roman", "Times-Bold"

# Etiquetas del cuadro de cabecera tal como las escribe el .docx (o como
# las escribiría una persona a mano).
DOCX_HEADER_LABELS = [
    ("numero", re.compile(r"^n[°º*o]?\s*[:.]", re.I)),
    ("pagina", re.compile(r"^p[aá]gina\s*:?", re.I)),
    ("area", re.compile(r"^[aá]rea\s*:?", re.I)),
    ("campo", re.compile(r"^campo\s*:?", re.I)),
    ("norma_titulo", re.compile(r"^norma\s*:?", re.I)),
    ("fecha_elaboracion", re.compile(r"^fecha de elaboraci[oó]n\s*:?", re.I)),
    ("reemplaza", re.compile(r"^reemplaza a la norma del\s*:?", re.I)),
    ("elaborada_por", re.compile(r"^elaborada? por\s*:?", re.I)),
    ("revisado_por", re.compile(r"^revisado por\s*:?", re.I)),
    ("referencia_corporativa", re.compile(r"^referencia corporativa(?: sl)?\s*:?", re.I)),
    ("revisado_aprobado", re.compile(r"^revisado y aprobado por\s*:?", re.I)),
    ("autorizado_por", re.compile(r"^autorizado por el gerente general\s*:?", re.I)),
]


# ------------------------------------------------------------------ #
# 1. Lectura del .docx
# ------------------------------------------------------------------ #

def _leer_parrafo(p_el, part):
    """Devuelve (runs, indent, text, page_break, image) de un <w:p>."""
    runs, indent, page_break, image = [], 0, False, None
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is not None and ind.get(qn("w:left")):
            indent = int(ind.get(qn("w:left")))
    for r in p_el.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        bold = False
        if rpr is not None:
            b = rpr.find(qn("w:b"))
            if b is not None:
                val = b.get(qn("w:val"))
                bold = val in (None, "", "1", "true", "on")
        for child in r:
            tag = child.tag
            if tag == qn("w:t"):
                runs.append((child.text or "", bold))
            elif tag == qn("w:tab"):
                runs.append(("\t", bold))
            elif tag == qn("w:br"):
                if child.get(qn("w:type")) == "page":
                    page_break = True
                else:
                    runs.append(("\n", bold))
            elif tag == qn("w:drawing"):
                for blip in child.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
                    rid = blip.get(qn("r:embed"))
                    if rid and rid in part.rels:
                        rel = part.rels[rid]
                        image = {"bytes": rel.target_part.blob,
                                 "ext": str(rel.target_ref).split(".")[-1].lower()}
                        break
    text = "".join(t for t, _ in runs).replace("\t", " ").strip()
    return runs, indent, text, page_break, image


def leer_docx(path):
    """Lista de elementos en el orden del documento (ver word_a_pdf.js)."""
    doc = Document(str(path))
    part = doc.part
    elementos = []

    def parrafos_de(container_el):
        pars = []
        for p in container_el.iter(qn("w:p")):
            runs, indent, text, page_break, image = _leer_parrafo(p, part)
            if image:
                pars.append({"type": "image", **image})
            if text:
                pars.append({"type": "paragraph", "runs": runs, "indent": indent, "text": text})
        return pars

    for node in doc.element.body:
        if node.tag == qn("w:p"):
            runs, indent, text, page_break, image = _leer_parrafo(node, part)
            if image:
                elementos.append({"type": "image", **image})
            if text:
                elementos.append({"type": "paragraph", "runs": runs, "indent": indent, "text": text})
            if page_break:
                elementos.append({"type": "pagebreak"})
        elif node.tag == qn("w:tbl"):
            rows = []
            for tr in node.findall(qn("w:tr")):
                cells = [parrafos_de(tc) for tc in tr.findall(qn("w:tc"))]
                rows.append(cells)
            elementos.append({"type": "table", "rows": rows})
    return elementos


# ------------------------------------------------------------------ #
# 2. Interpretación: cabecera y cuerpo
# ------------------------------------------------------------------ #

def limpiar_valor(v):
    v = re.sub(r"^\s*[:.\-–—]+\s*", "", v or "")
    return re.sub(r"\s+", " ", v).strip()


def parse_header_lineas(lineas):
    f, encontrado = {}, 0
    for raw in lineas:
        ln = raw.strip()
        if not ln:
            continue
        if re.match(r"^SPI\s+SWISS", ln, re.I) or re.match(r"^Manual de Procedimientos", ln, re.I):
            encontrado += 1
            continue
        matched = False
        for key, rx in DOCX_HEADER_LABELS:
            m = rx.match(ln)
            if m:
                f[key] = limpiar_valor(ln[m.end():])
                encontrado += 1
                matched = True
                break
        if not matched and "norma_titulo" in f and not f.get("fecha_elaboracion"):
            f["norma_titulo"] = (f["norma_titulo"] + " " + ln).strip()
    if "numero" not in f and "pagina" not in f:
        return None
    return f if encontrado >= 3 else None


def _es_negrita(p):
    return bool(p["runs"]) and all(b or not t.strip() for t, b in p["runs"])


def es_titulo_parrafo(p):
    if p["type"] != "paragraph":
        return False
    t = p["text"]
    if _es_negrita(p):
        return True
    if len(t) <= 45 and len(t.split()) <= 4 and t.endswith(":") and t[:1].isupper():
        return True
    return len(t) <= 45 and t.isupper() and not t.endswith((".", ",", ";"))


def bloques_desde_parrafos(pars):
    blocks, i = [], 0
    while i < len(pars):
        p = pars[i]
        if p["type"] == "image":
            blocks.append({"type": "image", "bytes": p["bytes"], "ext": p["ext"]})
            i += 1
            continue
        text, bold = p["text"], _es_negrita(p)
        if re.match(r"^OBJETIVOS?\s*:?\s*$", text, re.I):
            partes, i = [], i + 1
            while i < len(pars) and pars[i]["type"] == "paragraph" and not es_titulo_parrafo(pars[i]):
                partes.append(pars[i]["text"])
                i += 1
            blocks.append({"type": "objetivo", "text": " ".join(partes)})
            continue
        if re.match(r"^RESPONSABLES?\s*:?\s*$", text, re.I):
            names, i = [], i + 1
            while i < len(pars) and pars[i]["type"] == "paragraph" and not es_titulo_parrafo(pars[i]):
                names.append(pars[i]["text"])
                i += 1
            blocks.append({"type": "responsables", "names": names})
            continue
        m_num = re.match(r"^(\d{1,2})[.)]\s+(.*)$", text)
        m_let = re.match(r"^([a-h])\)\s+(.*)$", text)
        m_bul = re.match(r"^[•·\-*]\s+(.*)$", text)
        if m_num and not bold:
            blocks.append({"type": "numbered", "n": m_num.group(1), "text": m_num.group(2)})
        elif m_let and not bold:
            blocks.append({"type": "letter", "letter": m_let.group(1), "text": m_let.group(2)})
        elif m_bul:
            blocks.append({"type": "bullet", "text": m_bul.group(1)})
        elif bold or es_titulo_parrafo(p):
            blocks.append({"type": "subheading", "text": text})
        else:
            blocks.append({"type": "paragraph", "text": text})
        i += 1
    return blocks


def paginas_desde_docx(elementos):
    paginas, actual = [], None

    def cerrar():
        nonlocal actual
        if actual:
            paginas.append(actual)
            actual = None

    for el in elementos:
        if el["type"] == "pagebreak":
            cerrar()
        elif el["type"] == "table":
            celdas = [p for row in el["rows"] for cell in row for p in cell]
            lineas = [p["text"] for p in celdas if p["type"] == "paragraph"]
            fields = parse_header_lineas(lineas) if len(el["rows"]) >= 2 else None
            if fields:
                cerrar()
                actual = {"fields": fields, "blocks": []}
                continue
            if actual is None:
                actual = {"fields": None, "blocks": []}
            actual["blocks"].extend(bloques_desde_parrafos(celdas))
        elif el["type"] == "image":
            if actual and actual["fields"] and not actual["blocks"]:
                actual["blocks"].append({"type": "image", "bytes": el["bytes"], "ext": el["ext"]})
            else:
                cerrar()
                paginas.append({"fullImage": {"bytes": el["bytes"], "ext": el["ext"]}})
        elif el["type"] == "paragraph":
            if actual is None:
                actual = {"fields": None, "blocks": []}
            actual["blocks"].extend(bloques_desde_parrafos([el]))
    cerrar()

    # Word "plano" (sin tablas): cabecera desde las primeras líneas.
    for pg in paginas:
        if "fullImage" in pg or pg.get("fields"):
            continue
        textos = [b["text"] for b in pg["blocks"] if b.get("text")]
        fields = parse_header_lineas(textos[:16])
        if fields:
            pg["fields"] = fields
            vistos = 0
            for k, b in enumerate(pg["blocks"][:16]):
                t = b.get("text", "")
                if (re.match(r"^SPI\s+SWISS", t, re.I) or re.match(r"^Manual de Procedimientos", t, re.I)
                        or any(rx.match(t) for _, rx in DOCX_HEADER_LABELS)):
                    vistos = k + 1
            pg["blocks"] = pg["blocks"][vistos:]
    return paginas


# ------------------------------------------------------------------ #
# 3. Dibujo del PDF (reportlab)
# ------------------------------------------------------------------ #

def texto_winansi(s):
    s = re.sub(r"[‘’‚]", "'", s)
    s = re.sub(r"[“”„]", '"', s)
    s = s.replace("…", "...").replace("–", "-").replace("—", "-").replace(" ", " ")
    return "".join(c if (c in "\t\n\r" or 0x20 <= ord(c) <= 0x7E or 0xA0 <= ord(c) <= 0xFF or c in "€•") else "?" for c in s)


def envolver_texto(texto, font, size, max_w):
    palabras = texto_winansi(texto).split()
    lineas, actual = [], ""
    for w in palabras:
        prueba = (actual + " " + w) if actual else w
        if stringWidth(prueba, font, size) <= max_w or not actual:
            actual = prueba
        else:
            lineas.append(actual)
            actual = w
    if actual:
        lineas.append(actual)
    return lineas or [""]


class PdfEscritor:
    def __init__(self, c):
        self.c = c
        self.y = 0
        self.primera = True

    def nueva_pagina(self):
        if not self.primera:
            self.c.showPage()
        self.primera = False
        self.y = PDF_PAGE_H - PDF_MARGIN

    def _texto(self, s, x, y, font, size):
        self.c.setFont(font, size)
        self.c.drawString(x, y, texto_winansi(s))

    def _rect(self, x, y, w, h):
        self.c.setLineWidth(0.75)
        self.c.rect(x, y, w, h)

    def dibujar_cabecera(self, f):
        x0, x_mid, top = PDF_BOX_X, PDF_BOX_X + PDF_HDR_LEFT_W, self.y
        izq1 = [("SPI SWISSÔTEL LIMA", FONT_BOLD, 11), ("Manual de Procedimientos", FONT_REG, 10)]
        der1 = [(f"N°  : {f.get('numero', '')}", FONT_BOLD, 10), (f"Página: {f.get('pagina', '')}", FONT_REG, 10)]
        izq2 = [(f"Area  : {f.get('area', '')}", FONT_REG, 9),
                (f"Campo : {f.get('campo', '')}", FONT_REG, 9),
                (f"Norma : {f.get('norma_titulo', '')}", FONT_REG, 9)]
        der2 = []
        for key, label in (("fecha_elaboracion", "Fecha de elaboración"),
                           ("reemplaza", "Reemplaza a la norma del"),
                           ("elaborada_por", "Elaborada por"),
                           ("revisado_por", "Revisado por"),
                           ("referencia_corporativa", "Referencia Corporativa SL"),
                           ("revisado_aprobado", "Revisado y aprobado por")):
            if f.get(key):
                der2.append((f"{label} : {f[key]}", FONT_REG, 9))
        der2.append((f"Autorizado por el Gerente General : {f.get('autorizado_por', '')}".strip(), FONT_REG, 9))

        w_izq = PDF_HDR_LEFT_W - 2 * PDF_PAD
        w_der = PDF_BOX_W - PDF_HDR_LEFT_W - 2 * PDF_PAD

        def alto(items, max_w):
            h = PDF_PAD
            for t, font, size in items:
                h += len(envolver_texto(t, font, size, max_w)) * size * PDF_LINE_GAP + PDF_PARA_GAP
            return h + PDF_PAD - PDF_PARA_GAP

        h1 = max(alto(izq1, w_izq), alto(der1, w_der))
        h2 = max(alto(izq2, w_izq), alto(der2, w_der))

        def celda(items, x, y_top, max_w):
            y = y_top - PDF_PAD
            for t, font, size in items:
                for ln in envolver_texto(t, font, size, max_w):
                    y -= size * PDF_LINE_GAP
                    self._texto(ln, x + PDF_PAD, y + size * 0.25, font, size)
                y -= PDF_PARA_GAP

        celda(izq1, x0, top, w_izq)
        celda(der1, x_mid, top, w_der)
        celda(izq2, x0, top - h1, w_izq)
        celda(der2, x_mid, top - h1, w_der)
        self._rect(x0, top - h1, PDF_HDR_LEFT_W, h1)
        self._rect(x_mid, top - h1, PDF_BOX_W - PDF_HDR_LEFT_W, h1)
        self._rect(x0, top - h1 - h2, PDF_HDR_LEFT_W, h2)
        self._rect(x_mid, top - h1 - h2, PDF_BOX_W - PDF_HDR_LEFT_W, h2)
        self.y = top - h1 - h2 - 12

    def _imagen(self, img):
        try:
            im = Image.open(io.BytesIO(img["bytes"]))
            if im.mode not in ("RGB", "L"):
                im = im.convert("RGB")
            return im
        except Exception as e:  # noqa: BLE001
            print(f"  ! No se pudo leer una imagen del Word: {e}", file=sys.stderr)
            return None

    def dibujar_contenido(self, blocks, fields):
        from reportlab.lib.utils import ImageReader

        x0 = PDF_BOX_X
        max_w = PDF_BOX_W - 2 * PDF_PAD
        state = {"box_top": self.y, "y": self.y - PDF_PAD}
        bottom = PDF_MARGIN + PDF_PAD
        solo_imagenes = bool(blocks) and all(b["type"] == "image" for b in blocks)

        def cerrar():
            if solo_imagenes:
                return
            h = state["box_top"] - (state["y"] - PDF_PAD)
            self._rect(x0, state["y"] - PDF_PAD, PDF_BOX_W, h)

        def saltar():
            cerrar()
            self.nueva_pagina()
            if fields:
                self.dibujar_cabecera(fields)
            state["box_top"] = self.y
            state["y"] = self.y - PDF_PAD

        items = []
        for b in blocks:
            t = b["type"]
            if t == "objetivo":
                items += [(FONT_BOLD, 10, "OBJETIVO:", 0), (FONT_REG, 10, b["text"], 0)]
            elif t == "responsables":
                items.append((FONT_BOLD, 10, "RESPONSABLES:", 0))
                items += [(FONT_REG, 10, n, 0) for n in b["names"]]
            elif t == "subheading":
                items.append((FONT_BOLD, 10, b["text"], 0))
            elif t == "numbered":
                items.append((FONT_REG, 10, f"{b['n']}.  {b['text']}", PDF_INDENT_LIST))
            elif t == "letter":
                items.append((FONT_REG, 10, f"{b['letter']})  {b['text']}", PDF_INDENT_LETTER))
            elif t == "bullet":
                items.append((FONT_REG, 10, f"•  {b['text']}", PDF_INDENT_LIST))
            elif t == "image":
                items.append(("IMG", b))
            else:
                items.append((FONT_REG, 10, b["text"], 0))

        for it in items:
            if it[0] == "IMG":
                im = self._imagen(it[1])
                if im is None:
                    continue
                w = max_w
                h = im.height / im.width * w
                if state["y"] - h < bottom:
                    saltar()
                self.c.drawImage(ImageReader(im), x0 + PDF_PAD, state["y"] - h, width=w, height=h)
                state["y"] -= h + PDF_PARA_GAP
                continue
            font, size, text, indent = it
            for ln in envolver_texto(text, font, size, max_w - indent):
                if state["y"] - size * PDF_LINE_GAP < bottom:
                    saltar()
                state["y"] -= size * PDF_LINE_GAP
                self._texto(ln, x0 + PDF_PAD + indent, state["y"] + size * 0.25, font, size)
            state["y"] -= PDF_PARA_GAP
        cerrar()
        self.y = state["y"] - PDF_PAD - 12

    def dibujar_imagen_completa(self, img):
        from reportlab.lib.utils import ImageReader

        im = self._imagen(img)
        if im is None:
            return
        w = PDF_PAGE_W - 2 * PDF_MARGIN
        h = min(im.height / im.width * w, PDF_PAGE_H - 2 * PDF_MARGIN)
        self.c.drawImage(ImageReader(im), PDF_MARGIN, self.y - h, width=w, height=h)
        self.y -= h


# ------------------------------------------------------------------ #
# 4. Orquestación
# ------------------------------------------------------------------ #

def convert(docx_path: Path, out_path: Path, verbose=True):
    elementos = leer_docx(docx_path)
    paginas = paginas_desde_docx(elementos)
    if not paginas:
        print(f"  ! {docx_path.name}: el Word no tiene contenido reconocible.", file=sys.stderr)
        return
    c = rl_canvas.Canvas(str(out_path), pagesize=A4)
    c.setTitle(docx_path.stem)
    esc = PdfEscritor(c)
    for n, pg in enumerate(paginas, start=1):
        esc.nueva_pagina()
        if "fullImage" in pg:
            if verbose:
                print(f"  · Página {n}: imagen completa (se copia tal cual).")
            esc.dibujar_imagen_completa(pg["fullImage"])
            continue
        if pg["fields"]:
            if verbose:
                print(f"  ✓ Página {n}: cabecera (N° {pg['fields'].get('numero', '?')}, "
                      f"página {pg['fields'].get('pagina', '?')}).")
            esc.dibujar_cabecera(pg["fields"])
        elif verbose:
            print(f"  ⚠ Página {n}: sin cuadro de cabecera reconocible; se escribe el texto con el mismo formato.")
        if pg["blocks"]:
            esc.dibujar_contenido(pg["blocks"], pg["fields"])
    c.save()
    if verbose:
        print(f"  -> {out_path}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("entrada", help="Archivo .docx o carpeta con varios .docx")
    ap.add_argument("-o", "--salida", help="Ruta del .pdf de salida (solo si 'entrada' es un archivo)")
    args = ap.parse_args()

    entrada = Path(args.entrada)
    if entrada.is_dir():
        docs = sorted(entrada.glob("*.docx"))
        if not docs:
            print("No hay archivos .docx en esa carpeta.")
            return
        for d in docs:
            print(f"Procesando {d.name}...")
            convert(d, d.with_suffix(".pdf"))
    else:
        out = Path(args.salida) if args.salida else entrada.with_suffix(".pdf")
        print(f"Procesando {entrada.name}...")
        convert(entrada, out)
    print("Listo.")


if __name__ == "__main__":
    main()
