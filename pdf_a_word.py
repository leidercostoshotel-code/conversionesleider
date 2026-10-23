#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_a_word.py
=============
Convierte los manuales de procedimientos "SPI Swissôtel Lima" (PDF escaneado,
sin capa de texto) a un Word (.docx) EDITABLE que replica el diseño original:
la tabla de cabecera (Área / Campo / Norma / N° / fechas) y el recuadro con
Objetivo / Responsables / Pasos a Seguir.

------------------------------------------------------------------------
INSTALACIÓN (una sola vez)
------------------------------------------------------------------------
Windows (con Chocolatey) / Mac (con Homebrew) / Linux (con apt):

    # Motor de OCR + idioma español
    apt install tesseract-ocr tesseract-ocr-spa poppler-utils      # Linux
    brew install tesseract tesseract-lang poppler                  # Mac
    choco install tesseract poppler                                 # Windows
                                                                     (en Windows,
                                                                     agrega también el
                                                                     paquete de idioma
                                                                     "spa" desde
                                                                     github.com/tesseract-ocr/tessdata)

    # Librerías de Python
    pip install python-docx pytesseract pdf2image pillow

------------------------------------------------------------------------
USO
------------------------------------------------------------------------
    python pdf_a_word.py archivo.pdf
    python pdf_a_word.py archivo.pdf -o salida.docx
    python pdf_a_word.py archivo.pdf --paginas 2-3        # solo esas páginas del PDF
    python pdf_a_word.py carpeta_con_pdfs/                # convierte todos los .pdf de la carpeta

------------------------------------------------------------------------
CÓMO FUNCIONA (y sus límites)
------------------------------------------------------------------------
1. Cada página del PDF se rasteriza a imagen (300 dpi).
2. Se le aplica OCR en español (Tesseract).
3. Se busca en el texto el patrón de cabecera fijo de estos manuales
   (SPI SWISSÔTEL LIMA / N° / Área / Campo / Norma / fechas). Si aparece,
   se reconstruye esa página como texto real: título en negrita, listas
   numeradas o con viñetas, subtítulos ("Recepción", "Puntos de Venta", etc.).
4. Si una página NO tiene ese patrón (formularios, organigramas, cuadros
   financieros de Excel, portadas), se inserta la página TAL CUAL como
   imagen, en vez de forzar un OCR que saldría mal. Así no se pierde
   información aunque no quede editable.
5. El OCR nunca es 100% exacto (fechas, siglas y nombres propios son los
   que más fallan). SIEMPRE revisa el .docx resultante antes de usarlo;
   piensa en este programa como quien te hace el primer borrador rápido,
   no como un reemplazo de la revisión final.

Este programa fue hecho a la medida del formato de estos manuales
("SPI SWISSÔTEL LIMA — Manual de Procedimientos"). Con otro tipo de PDF
probablemente no reconozca la cabecera y meta todo como imagen — lo cual
sigue siendo seguro (no pierde contenido), solo que no queda editable.
"""

import argparse
import re
from datetime import date
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.shared import Pt, Twips, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

FONT = "Arial"
DPI = 300

# Valores fijos que se escriben SIEMPRE en el .docx de salida, en vez de lo
# que diga el PDF original (pedido del usuario): la norma nueva reemplaza a
# la anterior en la fecha en que se hace la conversión, la elabora GPL/JACH,
# y el Sub Contralor es siempre el primer responsable.
ELABORADA_POR_FIJO = "GPL/JACH"
RESPONSABLE_FIJO = "Sub Contralor"


def fecha_de_hoy() -> str:
    return date.today().strftime("%d.%m.%Y")

# ------------------------------------------------------------------ #
# 1. PDF -> imágenes, imagen -> texto (OCR)
# ------------------------------------------------------------------ #

def pdf_to_images(pdf_path: Path, workdir: Path, dpi: int = DPI, page_range=None):
    """Rasteriza el PDF a una imagen PNG por página con pdftoppm (poppler)."""
    prefix = workdir / "page"
    cmd = ["pdftoppm", "-png", "-r", str(dpi)]
    if page_range:
        cmd += ["-f", str(page_range[0]), "-l", str(page_range[1])]
    cmd += [str(pdf_path), str(prefix)]
    subprocess.run(cmd, check=True, capture_output=True)
    images = sorted(workdir.glob("page-*.png"))
    if not images:
        # pdftoppm no rellena ceros si hay una sola página
        images = sorted(workdir.glob("page*.png"))
    return images


def ocr_page(image_path: Path) -> str:
    """OCR en español sobre una imagen de página. Devuelve el texto plano.

    Se usa --psm 4 ("una sola columna de texto de tamaños variables") en
    vez de --psm 6 ("bloque uniforme de texto): con --psm 6, Tesseract
    tiende a fusionar la tabla de cabecera (dos columnas) en una sola
    línea por fila y así se pierde por completo la segunda línea de esa
    tabla (donde está "Página : X/Y"), lo que hacía que la página cayera
    al modo imagen aunque el formato fuera reconocible.
    """
    result = subprocess.run(
        ["tesseract", str(image_path), "-", "-l", "spa", "--psm", "4"],
        check=True, capture_output=True, text=True,
    )
    return result.stdout


# ------------------------------------------------------------------ #
# 2. Reconocer la cabecera fija del manual
# ------------------------------------------------------------------ #

RIGHT_LABELS = [
    ("fecha_elaboracion", r"Fecha de elaboraci[oó0]n"),
    ("reemplaza", r"Reemplaza a la norma del"),
    ("referencia_corporativa", r"Referencia Corporativa"),
    ("elaborada_por", r"Elaborada por"),
    ("revisado_aprobado", r"Revisado y aprobado por"),
    # Sin "General" al final: ese campo siempre queda en blanco en estos
    # manuales, y el OCR a veces lee mal esa última palabra (p. ej. "General"
    # -> "cabal"). "Autorizado por el Gerente" ya es una frase lo bastante
    # específica para no matchear nada por accidente.
    ("autorizado_por", r"Autorizado por el Gerente"),
]
LEFT_LABELS = [
    ("area", r"[AÁ]rea"),
    ("campo", r"Campo"),
    ("norma_titulo", r"Norma"),
]
ALL_LABELS = LEFT_LABELS + RIGHT_LABELS
LABEL_RE = re.compile("|".join(f"(?P<{k}>{p})" for k, p in ALL_LABELS))
LEFT_KEYS = {k for k, _ in LEFT_LABELS}
# Fila de la tabla de cabecera a la que pertenece cada etiqueta de la
# derecha -> campo de la izquierda de esa misma fila. Sirve cuando el OCR se
# come la etiqueta izquierda ("Area :", "Campo :", "Norma :") y solo queda su
# valor suelto delante de la etiqueta derecha.
ROW_LEFT_KEY = {
    "fecha_elaboracion": "area",
    "reemplaza": "campo",
    "elaborada_por": "norma_titulo",
    "autorizado_por": "norma_titulo",
}


def clean_value(s: str) -> str:
    """Quita el '|' que a veces deja el OCR como ruido del borde de la
    columna de la tabla (p. ej. 'Bono por Fallecimiento de |'), y normaliza
    espacios repetidos que puedan quedar al quitarlo."""
    s = re.sub(r"[|\[\]]", " ", s)
    s = re.sub(r"[\s\-–—]+$", "", s)  # guiones sueltos del borde de columna
    return re.sub(r"\s+", " ", s).strip()

NUMERO_RE = re.compile(r"\b(\d{2}\.\d{2}(?:\.\d{2}){0,2})\b")
# El N° también viene en el nombre del archivo ("SPI 04.01.13.02 Despacho...").
# Ese valor es exacto, mientras que el OCR a veces cambia el último dígito,
# así que cuando está en el nombre manda sobre lo leído de la imagen.
FILENAME_NUM_RE = re.compile(r"(?<!\d)(\d{2}\.\d{2}(?:\.\d{2}){0,2})(?!\d)")


def numero_desde_nombre(nombre):
    m = FILENAME_NUM_RE.search(nombre or "")
    return m.group(1) if m else None
# La barra de "2/3" es opcional: Tesseract.js (el motor del navegador) a
# veces se la come y lee "Página : 23". Como va justo detrás de la etiqueta
# "Página", esos dígitos pegados son casi seguro "página/total"; ver
# normalizar_pagina().
# "P?" porque con el render de pdf.js Tesseract a veces se come la "P"
# inicial y lee "agina: 1/1".
PAGINA_RE = re.compile(r"P?[aáé]gina[^\d]{0,40}(\d{1,2}\s*[/\\|]?\s*\d{1,2})")


def normalizar_pagina(raw: str) -> str:
    s = re.sub(r"\s+", "", raw).replace("\\", "/").replace("|", "/")
    if "/" in s:
        return s
    # Sin separador: "23" -> "2/3", "110" -> "1/10", "1013" -> "10/13".
    if len(s) == 2:
        return f"{s[0]}/{s[1]}"
    if len(s) == 3:
        return f"{s[0]}/{s[1:]}"
    if len(s) == 4:
        return f"{s[:2]}/{s[2:]}"
    return s


def parse_header(text: str):
    """
    Busca el bloque de cabecera (primeras líneas, antes de 'OBJETIVO' /
    'PASOS A SEGUIR') y extrae sus campos. Devuelve (dict_o_None, resto_texto).
    """
    lines = text.splitlines()

    # Marcador de inicio de cuerpo ("OBJETIVO(S):" / "PASOS A SEGUIR"),
    # solo en MAYÚSCULAS (así están siempre los títulos de sección; buscarlo
    # sin distinguir mayúsculas confundía "los pasos a seguir" dentro de una
    # oración normal) y solo dentro de las primeras líneas: en una página de
    # continuación puede aparecer un "OBJETIVOS:" a mitad de página (empieza
    # un sub-procedimiento nuevo) y NO marca el fin de la cabecera.
    marker_idx = None
    for i, ln in enumerate(lines[:15]):
        if re.search(r"OBJETIVOS?\s*:", ln) or re.search(r"PASOS A SEGUIR", ln):
            marker_idx = i
            break

    # El fin real de la cabecera lo deciden las etiquetas fijas: la última
    # línea con alguna etiqueta, más hasta 2 líneas cortas de continuación
    # (valores envueltos a varias líneas, como "Norma" o "Autorizado por").
    # Así, un título del cuerpo que venga justo después de la cabecera y
    # antes de "OBJETIVOS:" (p. ej. "GASTOS DE CAPITAL") no se pega al
    # campo "Norma". El marcador, si existe, solo actúa como tope.
    last_label_idx = None
    for i, ln in enumerate(lines[:15]):
        if LABEL_RE.search(ln):
            last_label_idx = i
    if last_label_idx is None:
        body_start_idx = marker_idx if marker_idx is not None else min(10, len(lines))
    else:
        idx = last_label_idx + 1
        while (
            idx < len(lines)
            and idx <= last_label_idx + 2
            and lines[idx].strip()
            and len(lines[idx].strip()) <= 50
            and not NUM_ITEM_RE.match(lines[idx])
        ):
            idx += 1
        body_start_idx = idx
        if marker_idx is not None:
            body_start_idx = min(body_start_idx, marker_idx)

    header_lines = lines[:body_start_idx]
    header_text = "\n".join(header_lines)

    numero_match = NUMERO_RE.search(header_text)
    pagina_match = PAGINA_RE.search(header_text)
    # En algunos PDF el OCR se come la línea "Página : 1/1" entera. Si aun
    # así hay N° y alguna etiqueta inequívoca del cuadro (Fecha de
    # elaboración / Elaborada por / Manual de Procedimientos), la cabecera
    # se acepta igual y el número de página se completa después contando
    # las páginas reconocidas (ver convert()).
    otras_etiquetas = re.search(
        r"Fecha de elaboraci|Elaborada por|Manual de Procedimientos", header_text, re.I
    )
    if not numero_match or not (pagina_match or otras_etiquetas):
        return None, text  # no se reconoce cabecera -> esta página va como imagen

    fields = {
        "numero": numero_match.group(1),
        "pagina": normalizar_pagina(pagina_match.group(1)) if pagina_match else "",
    }
    # Se rastrea la última etiqueta activa de cada columna por separado
    # (izquierda: Área/Campo/Norma: derecha: Fecha/Reemplaza/Elaborada
    # por/Autorizado...). El OCR a veces fusiona ambas columnas en una
    # misma línea física (p. ej. "Norma : ... | Elaborada por : ..."), y
    # cuando eso pasa, el texto de continuación de una columna (el título
    # de "Norma" envuelto a varias líneas) puede terminar en una línea
    # cuya única etiqueta reconocible es de la OTRA columna. En estos
    # manuales, el único campo que se envuelve a varias líneas es "Norma"
    # (columna izquierda), así que ese texto suelto se prioriza para la
    # columna izquierda en vez de perderse o pegarse a la derecha.
    last_left_key = None
    last_right_key = None
    for ln in header_lines:
        matches = list(LABEL_RE.finditer(ln))
        if not matches:
            target = last_left_key or last_right_key
            if target:
                fields[target] = (fields.get(target, "") + " " + ln.strip()).strip()
            continue

        first_start = matches[0].start()
        if first_start > 0:
            pre_text = clean_value(re.sub(r"^[\s:\-–—]+", "", ln[:first_start]))
            row_left = ROW_LEFT_KEY.get(matches[0].lastgroup)
            if pre_text and row_left and not fields.get(row_left):
                # La etiqueta izquierda se perdió en el OCR (con el render de
                # pdf.js, Tesseract omite "Area :" y deja solo
                # "- Contraloría Fecha de elaboración - 14.03.15"): el texto
                # previo a la etiqueta derecha es el valor izquierdo de esa
                # misma fila de la tabla.
                fields[row_left] = pre_text
                last_left_key = row_left
            else:
                target = last_left_key or last_right_key
                if pre_text and target:
                    fields[target] = (fields.get(target, "") + " " + pre_text).strip()

        for idx, m in enumerate(matches):
            key = m.lastgroup
            val_start = m.end()
            val_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(ln)
            # También se quita un guión inicial: el OCR del navegador a veces
            # lee el ":" de la tabla como "-" ("Elaborada por - GPL").
            value = clean_value(ln[val_start:val_end].lstrip(" :-–—"))
            fields[key] = (fields.get(key, "") + " " + value).strip() if key in fields and key not in ("area", "campo") else value
            if key in LEFT_KEYS:
                last_left_key = key
            else:
                last_right_key = key

    # Fecha sin puntos ("160412"): el OCR a veces se come los separadores.
    # Fecha sin puntos o con puntos perdidos ("160412", "2302.15", "23.0910"):
    # si quedan exactamente 6 dígitos, se reescribe como DD.MM.AA.
    fe = fields.get("fecha_elaboracion", "")
    fe_digits = re.sub(r"\D", "", fe)
    if len(fe_digits) == 6 and re.fullmatch(r"[\d.\s]+", fe):
        fields["fecha_elaboracion"] = f"{fe_digits[:2]}.{fe_digits[2:4]}.{fe_digits[4:]}"

    body_text = "\n".join(lines[body_start_idx:])
    return fields, body_text


# ------------------------------------------------------------------ #
# 3. Reconocer el cuerpo: Objetivo / Responsables / Pasos a Seguir
# ------------------------------------------------------------------ #

# Paso numerado: "1. ", "1) ", y también las variantes que trae el OCR o el
# propio PDF: "1.- ", "1- ", "7.-Estos" (sin espacio). Lo que sigue al
# marcador debe empezar con letra o comilla, para no confundirlo con una
# cifra como "75.00".
NUM_ITEM_RE = re.compile(r"^\s*(\d{1,2})\s*(?:\.-|[.\)]|-)\s*(?=[A-Za-zÁÉÍÓÚÑÜáéíóúñü“\"(¿¡])(.*)")
LETTER_ITEM_RE = re.compile(r"^\s*([a-h])\)\s+(.*)")
# Además de los glifos reales de viñeta, se aceptan "+" y "=": así lee el
# OCR, respectivamente, al "•" pequeño y al "▪" cuadrado de algunos PDF.
# El "." al inicio también cuenta: es como el navegador suele leer la viñeta.
BULLET_ITEM_RE = re.compile(r"^\s*[•·\-\*\+=.]\s+(.*)")
# El glifo de viñeta (•), en algunos PDF/fuentes, el OCR lo confunde con una
# "e" suelta (ambigüedad visual con el punto de la viñeta). Se trata como
# viñeta solo cuando es una "e" aislada seguida de mayúscula (inicio real de
# oración), para no confundir la conjunción "e" ("madre e hija").
BULLET_OCR_E_RE = re.compile(r"^\s*e\s+(?=[A-ZÁÉÍÓÚÑ])(.*)")
# Título de sub-sección numerado sin ":" (p. ej. "1.- RUTINA", "2.- NO
# RUTINA"). Se distingue de un paso numerado real (NUM_ITEM_RE) porque acá
# el punto va seguido de un guión, no de un espacio.
SECTION_NUM_HEADING_RE = re.compile(r"^\s*\d{1,2}\.-\s+.{1,60}$")

# Ruido típico de capturas de un visor web (fecha/hora, migas de pan,
# barra de usuario, pie con URL y contador de página) que se mete en el
# OCR cuando el PDF es en realidad una captura de pantalla de un visor.
NOISE_SUB_PATTERNS = [
    r"https?\s*:\s*//\S*[\w/=?.]*",
    r"\d{1,2}/\d{1,2}/\d{2,4},?\s*\d{1,2}:\d{2}\s*[ap]\.?\s*m\.?\.?",
    r"\b\d{1,2}\s*/\s*\d{1,2}\b(?=\s*$)",
]
NOISE_LINE_PATTERNS = [
    r"Intranet corporativa",
    r"Cerrar sesi[oó]n",
    r"—\s*Intranet\s*$",
    r"^\s*[←<]\s*\w+\s*[-–—]?\s*\d{0,3}%?\s*[-+]?\s*$",
    r"^\s*P[aá]g\.?\s*\d+\s*$",
    r"Documento de uso interno",
]


def is_garbage_line(line: str) -> bool:
    """Línea de "basura" del OCR: el navegador convierte las franjas grises
    de la intranet en ristras como "ITTiaIl ol iTíimig1aíditGi". En ellas
    más de la mitad de las letras son palos verticales (I, l, i, T, 1, í, Í)."""
    letters = [c for c in line if c.isalnum()]
    if len(letters) < 8:
        return False
    palos = sum(1 for c in letters if c in "IliTt1íÍ|")
    return palos / len(letters) > 0.5


def strip_noise(text: str) -> str:
    for pat in NOISE_SUB_PATTERNS:
        text = re.sub(pat, " ", text)
    lines = text.splitlines()
    lines = [ln for ln in lines if not any(re.search(p, ln, re.I) for p in NOISE_LINE_PATTERNS)]
    lines = [ln for ln in lines if not is_garbage_line(ln)]
    return "\n".join(lines)


def is_heading_line(line: str) -> bool:
    """Título de sección corto, con mayúscula inicial y terminado en ':'
    (p. ej. "RESPONSABLES:", "BASE LEGAL:", "CARACTERISTICAS:").
    Se exige además que sean pocas palabras (<=4): los títulos reales de
    este tipo de manuales son de 1-3 palabras, a diferencia de una
    oración normal que también puede terminar en ':' antes de una lista
    (p. ej. "El formato se distribuye como sigue:"), que no es un
    título aunque cumpla lo demás."""
    line = line.strip()
    return (
        bool(line)
        and len(line) <= 45
        and len(line.split()) <= 4
        and line.endswith(":")
        and line[:1].isupper()
    )


def split_section(text, start_pat, end_pats):
    """Devuelve (prefijo_antes_de_la_marca, valor_de_la_sección, resto).
    `prefijo` es lo que había ANTES de encontrar `start_pat` -- normalmente
    vacío, salvo cuando hay una sección intermedia no prevista (p. ej.
    "BASE LEGAL:" entre "OBJETIVO:" y "RESPONSABLES:") que de otro modo se
    perdería en silencio. Si no se encuentra `start_pat`, devuelve
    (None, None, text) sin consumir nada."""
    # Los marcadores de sección se buscan anclados al inicio de línea y en
    # MAYÚSCULAS (así aparecen siempre en estos manuales): buscarlos sin
    # distinguir mayúsculas hacía que "los pasos a seguir" dentro de una
    # oración normal cortara la sección "OBJETIVO" por la mitad.
    m = re.search(start_pat, text, re.M)
    if not m:
        return None, None, text, None
    prefix = text[:m.start()]
    rest = text[m.end():]
    end_pos = len(rest)
    for ep in end_pats:
        em = re.search(ep, rest, re.M)
        if em and em.start() < end_pos:
            end_pos = em.start()
    # Además de los patrones explícitos (que solo cubren los nombres de
    # sección esperados), cualquier línea que en sí misma parezca un
    # título de sección marca el final de este bloque. Esto generaliza a
    # secciones no previstas (p. ej. "BASE LEGAL:", "CARACTERISTICAS:")
    # que de otro modo quedarían arrastradas dentro de "OBJETIVO" o
    # "RESPONSABLES" hasta la siguiente sección conocida.
    pos = 0
    for line in rest.splitlines(keepends=True):
        if pos >= end_pos:
            break
        if pos > 0 and is_heading_line(line):
            end_pos = pos
            break
        pos += len(line)
    return prefix, rest[:end_pos].strip(), rest[end_pos:], m.group(0).strip()


def has_list_items(text: str) -> bool:
    return any(
        NUM_ITEM_RE.match(ln) or LETTER_ITEM_RE.match(ln)
        or BULLET_ITEM_RE.match(ln) or BULLET_OCR_E_RE.match(ln)
        for ln in text.splitlines()
    )


def parse_body(body_text: str):
    """
    Devuelve una lista de bloques tipados:
      {'type': 'objetivo'|'responsables'|'subheading'|'numbered'|'bullet'|'letter'|'paragraph', ...}
    """
    blocks = []
    remaining = body_text

    pre_objetivo, objetivo, remaining, objetivo_marker = split_section(
        remaining, r"^\s*OBJETIVOS?\s*:", [r"^\s*RESPONSABLES?\s*:", r"^\s*PASOS A SEGUIR", r"^\s*POL[IÍ]TICAS"]
    )
    if pre_objetivo and pre_objetivo.strip():
        blocks.extend(parse_steps(pre_objetivo))
    if objetivo_marker and not objetivo:
        # El marcador está al final de la página y su contenido sigue en la
        # siguiente: se conserva el título igual.
        blocks.append({"type": "subheading", "text": objetivo_marker})
    elif objetivo:
        if has_list_items(objetivo):
            # Algunos manuales ponen el objetivo como lista de viñetas (p. ej.
            # bajo "OBJETIVOS:"): se conserva esa estructura en vez de
            # aplastarla en un solo párrafo.
            blocks.append({"type": "subheading", "text": objetivo_marker})
            blocks.extend(parse_steps(objetivo))
        else:
            blocks.append({"type": "objetivo", "text": clean_paragraph(objetivo)})

    pre_responsables, responsables, remaining, _ = split_section(
        remaining, r"^\s*RESPONSABLES?\s*:", [r"^\s*PASOS A SEGUIR", r"^\s*POL[IÍ]TICAS"]
    )
    if pre_responsables and pre_responsables.strip():
        # Sección intermedia no prevista (p. ej. "BASE LEGAL:") entre
        # "OBJETIVO:" y "RESPONSABLES:": se reconstruye con el parser
        # genérico en vez de perderla.
        blocks.extend(parse_steps(pre_responsables))
    if responsables:
        # La lista de nombres termina en el primer título que aparezca
        # (p. ej. "COSTO POR SERVICIO SEGÚN CONTRATO:"); lo que sigue se
        # reconstruye con el parser genérico.
        lineas = [ln.strip() for ln in responsables.splitlines() if ln.strip()]
        corte = len(lineas)
        for k, ln in enumerate(lineas):
            es_titulo = is_heading_line(ln) or (len(ln) <= 45 and ln.isupper() and not ln.endswith((".", ",", ";")))
            if es_titulo and k > 0 and len(ln) >= 3:
                corte = k
                break
        names = [ln.strip(" .") for ln in lineas[:corte]]
        blocks.append({"type": "responsables", "names": names})
        if corte < len(lineas):
            blocks.extend(parse_steps("\n".join(lineas[corte:])))

    # Lo que sigue (o todo el texto, si esta página es una continuación
    # sin encabezados) se interpreta como la lista de pasos.
    steps_text = remaining
    m = re.search(r"^\s*PASOS A SEGUIR\s*:?", steps_text, re.M)
    if m:
        # Lo que haya entre la última sección reconocida y "PASOS A SEGUIR"
        # (p. ej. un "RESPONSABLE:" con un nombre que no coincidió con el
        # patrón esperado) se reconstruye con el parser genérico en vez de
        # descartarse.
        before = steps_text[:m.start()]
        if before.strip():
            blocks.extend(parse_steps(before))
        # El título "PASOS A SEGUIR:" también va al Word, como en el PDF.
        blocks.append({"type": "subheading", "text": "PASOS A SEGUIR:"})
        steps_text = steps_text[m.end():]

    blocks.extend(parse_steps(steps_text))
    return blocks


def clean_paragraph(text):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return " ".join(lines)


def parse_steps(text):
    # Se conserva si cada línea venía precedida de una línea en blanco:
    # el OCR separa los párrafos/ítems con un renglón vacío, y esa es la
    # única pista que queda cuando el navegador se come las viñetas.
    lines, blank_before = [], []
    pending_blank = False
    for raw in text.splitlines():
        ln = raw.strip()
        if not ln:
            pending_blank = True
            continue
        lines.append(ln)
        blank_before.append(pending_blank)
        pending_blank = False
    blocks = []
    current = None
    prev_len = 0  # largo de la última línea física procesada
    i = 0
    while i < len(lines):
        line = lines[i]
        this_len = len(line)

        m = NUM_ITEM_RE.match(line)
        # "1.- RUTINA" (texto corto en mayúsculas tras el "N.-") es un
        # subtítulo numerado, no un paso: se deja para la regla de títulos.
        titulo_numerado = bool(m and SECTION_NUM_HEADING_RE.match(line) and m.group(2).strip().isupper())
        if m and not titulo_numerado:
            if current:
                blocks.append(current)
            current = {"type": "numbered", "n": m.group(1), "text": m.group(2).strip()}
            prev_len = this_len
            i += 1
            continue

        m = LETTER_ITEM_RE.match(line)
        if m:
            if current:
                blocks.append(current)
            current = {"type": "letter", "letter": m.group(1), "text": m.group(2).strip()}
            prev_len = this_len
            i += 1
            continue

        m = BULLET_ITEM_RE.match(line) or BULLET_OCR_E_RE.match(line)
        if m:
            if current:
                blocks.append(current)
            current = {"type": "bullet", "text": m.group(1).strip()}
            prev_len = this_len
            i += 1
            continue

        # Un título de sección corto, en mayúscula inicial y terminado en
        # ":" (p. ej. "CONTROLES:", "REMUNERACIONES:", "PAGOS:") es en sí
        # mismo una señal inequívoca de subtítulo en este tipo de manuales
        # -- no hace falta que además reinicie una lista en "1.".
        # Una línea corta TODA en mayúsculas (p. ej. "DISPOSICIÓN Y CASTIGO DE
        # LOS ACTIVOS") es título de sección aunque pase de 4 palabras.
        all_caps_heading = len(line) <= 45 and line.isupper() and not line.endswith((".", ",", ";"))
        if is_heading_line(line) or SECTION_NUM_HEADING_RE.match(line) or all_caps_heading:
            if current:
                blocks.append(current)
                current = None
            blocks.append({"type": "subheading", "text": line})
            prev_len = this_len
            i += 1
            continue

        looks_like_heading = (
            len(line) <= 45
            and len(line.split()) <= 4
            and not line.endswith((".", ",", ";"))
            and line[:1].isupper()
        )
        # Un subtítulo sin ":" casi siempre viene seguido de un ítem "1."
        # (la lista de esa sub-sección vuelve a empezar). Usamos esa
        # señal aunque haya un ítem numerado abierto en ese momento.
        next_starts_list_at_1 = False
        if i + 1 < len(lines):
            nm = NUM_ITEM_RE.match(lines[i + 1])
            next_starts_list_at_1 = bool(nm and nm.group(1) == "1")

        # Un renglón corto aislado (línea en blanco antes y después) es un
        # subtítulo aunque haya un párrafo abierto (p. ej. "General").
        # Se exige además que el párrafo anterior haya terminado en ".": si
        # terminó en ":" o sin puntuación, lo que sigue es una lista.
        aislado = (
            blank_before[i]
            and i + 1 < len(lines)
            and blank_before[i + 1]
            and (current is None or current.get("text", "").endswith("."))
        )
        if looks_like_heading and (current is None or next_starts_list_at_1 or aislado):
            if current:
                blocks.append(current)
                current = None
            blocks.append({"type": "subheading", "text": line})
            prev_len = this_len
            i += 1
            continue

        # Con el render del navegador, el OCR a veces se come por completo
        # los glifos de viñeta y los ítems llegan sin marcador. Dentro de
        # un párrafo genérico, una línea que empieza en mayúscula justo
        # después de una que terminó en ".", ":" o ";" es un ítem nuevo,
        # no la continuación envuelta del anterior.
        # También es ítem nuevo si venía separado por una línea en blanco
        # o si la línea anterior era corta (una línea envuelta ocupa casi
        # todo el ancho; una corta es el final de su párrafo).
        # (Una línea que empieza con cifra, p. ej. "24 Horas 30.00", también
        # es ítem nuevo si viene tras línea en blanco o tras línea corta.)
        empieza_mayus_o_cifra = line[:1].isupper() or line[:1].isdigit()
        nuevo_item = (
            current is not None
            and (
                (current["type"] == "paragraph" and line[:1].isupper()
                 and re.search(r"[.:;]$", current["text"]) is not None)
                or (current["type"] in ("paragraph", "bullet") and empieza_mayus_o_cifra
                    and (blank_before[i] or prev_len < 50)
                    and not (line[:1].isdigit() and current["text"].endswith(":")))
            )
        )
        if current and not nuevo_item:
            current["text"] = (current["text"] + " " + line).strip()
        else:
            if current:
                blocks.append(current)
            # OJO: se deja `current` apuntando a este párrafo (no se hace
            # append todavía) para que la siguiente línea, si es una
            # continuación envuelta del mismo párrafo, se una aquí en vez
            # de crear un bloque nuevo por cada línea física del OCR.
            current = {"type": "paragraph", "text": line}
        prev_len = this_len
        i += 1

    if current:
        blocks.append(current)
    return blocks


# ------------------------------------------------------------------ #
# 4. Construcción del .docx (mismo diseño que el original)
# ------------------------------------------------------------------ #

def set_cell_border(cell, size=4):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tc_pr.append(borders)


def add_run(paragraph, text, bold=False, size=10):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)
    paragraph.paragraph_format.space_after = Pt(5)
    paragraph.paragraph_format.space_before = Pt(0)
    return run


def add_header_table(doc, f):
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [Twips(6800), Twips(3100)]
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w
            set_cell_border(cell)

    c = table.cell(0, 0)
    c.paragraphs[0].text = ""
    add_run(c.paragraphs[0], "SPI SWISSÔTEL LIMA", bold=True, size=11)
    p2 = c.add_paragraph()
    add_run(p2, "Manual de Procedimientos", size=10)

    c = table.cell(0, 1)
    add_run(c.paragraphs[0], f"N°  : {f.get('numero', '')}", bold=True, size=10)
    p2 = c.add_paragraph()
    add_run(p2, f"Página: {f.get('pagina', '')}", size=10)

    c = table.cell(1, 0)
    add_run(c.paragraphs[0], f"Area  : {f.get('area', '')}", size=9)
    for key, label in (("campo", "Campo"), ("norma_titulo", "Norma")):
        p = c.add_paragraph()
        add_run(p, f"{label} : {f.get(key, '')}", size=9)

    c = table.cell(1, 1)
    f = dict(f, reemplaza=fecha_de_hoy(), elaborada_por=ELABORADA_POR_FIJO)
    right_fields = [
        ("fecha_elaboracion", "Fecha de elaboración"),
        ("reemplaza", "Reemplaza a la norma del"),
        ("elaborada_por", "Elaborada por"),
        ("referencia_corporativa", "Referencia Corporativa SL"),
        ("revisado_aprobado", "Revisado y aprobado por"),
    ]
    first = True
    for key, label in right_fields:
        if not f.get(key):
            continue
        p = c.paragraphs[0] if first else c.add_paragraph()
        add_run(p, f"{label} : {f[key]}", size=9)
        first = False
    p = c.paragraphs[0] if first else c.add_paragraph()
    add_run(p, "Autorizado por el Gerente General :", size=9)
    return table


def add_content_box(doc, blocks):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Twips(9900)
    set_cell_border(cell)
    cell.paragraphs[0].text = ""
    first = True

    def new_para():
        nonlocal first
        if first:
            first = False
            return cell.paragraphs[0]
        return cell.add_paragraph()

    for b in blocks:
        t = b["type"]
        if t == "objetivo":
            add_run(new_para(), "OBJETIVO:", bold=True)
            add_run(new_para(), b["text"])
        elif t == "responsables":
            add_run(new_para(), "RESPONSABLES:", bold=True)
            otros = [n for n in b["names"] if n.strip().lower() != RESPONSABLE_FIJO.lower()]
            for name in [RESPONSABLE_FIJO] + otros:
                add_run(new_para(), name)
        elif t == "subheading":
            add_run(new_para(), b["text"], bold=True)
        elif t == "numbered":
            p = new_para()
            p.paragraph_format.left_indent = Twips(360)
            add_run(p, f"{b['n']}.  {b['text']}")
        elif t == "letter":
            p = new_para()
            p.paragraph_format.left_indent = Twips(500)
            add_run(p, f"{b['letter']})  {b['text']}")
        elif t == "bullet":
            p = new_para()
            p.paragraph_format.left_indent = Twips(360)
            add_run(p, f"•  {b['text']}")
        else:  # paragraph
            add_run(new_para(), b["text"])
    return table


def add_full_page_image(doc, image_path):
    section = doc.sections[-1]
    usable_w = section.page_width - section.left_margin - section.right_margin
    doc.add_picture(str(image_path), width=usable_w)


# Zona del cuerpo (fracciones del alto/ancho de la página) que ocupa el
# cuadro de contenido debajo de la cabecera en estos manuales. Se usa para
# los anexos: formularios con tablas que el OCR no puede reconstruir.
BODY_CROP = (0.08, 0.228, 0.93, 0.86)  # izquierda, arriba, derecha, abajo


def es_pagina_anexo(body_text: str) -> bool:
    """Las páginas "ANEXO N" son formularios (tablas, líneas de puntos)
    que como texto quedan vacíos: se insertan como imagen bajo la cabecera."""
    primeras = [ln.strip() for ln in body_text.splitlines() if ln.strip()][:3]
    return any(re.match(r"^ANEXO\s*\d*\s*$", ln, re.I) for ln in primeras)


def add_body_image(doc, image_path):
    """Inserta solo la zona del cuerpo de la página (bajo la cabecera)."""
    from PIL import Image
    section = doc.sections[-1]
    usable_w = section.page_width - section.left_margin - section.right_margin
    im = Image.open(image_path)
    w, h = im.size
    l, t, r, b = BODY_CROP
    recorte = im.crop((int(w * l), int(h * t), int(w * r), int(h * b)))
    tmp_path = Path(str(image_path) + ".body.png")
    recorte.save(tmp_path)
    doc.add_picture(str(tmp_path), width=usable_w)


def add_page_break(doc):
    doc.add_page_break()


# ------------------------------------------------------------------ #
# 5. Orquestación
# ------------------------------------------------------------------ #

def parse_page_range(spec, total_pages):
    if not spec:
        return None
    if "-" in spec:
        a, b = spec.split("-", 1)
        return (int(a), int(b))
    n = int(spec)
    return (n, n)


def completar_numeros_de_pagina(lista_fields):
    """Renumera 'pagina' ("k/N") en las cabeceras reconocidas: k es el
    orden de la página entre las reconocidas (el OCR confunde "7/7" con
    "1/7" o se come la línea entera) y N el total que más se repite en lo
    leído, o el número de páginas reconocidas si no se leyó ninguno."""
    reconocidas = [f for f in lista_fields if f is not None]
    totales = []
    for f in reconocidas:
        m = re.fullmatch(r"\s*\d{1,2}\s*/\s*(\d{1,2})\s*", f.get("pagina") or "")
        if m:
            totales.append(m.group(1))
    total = max(set(totales), key=totales.count) if totales else str(len(reconocidas))
    for k, f in enumerate(reconocidas, start=1):
        f["pagina"] = f"{k}/{total}"


def convert(pdf_path: Path, out_path: Path, page_range_spec=None, keep_tmp=False, verbose=True):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        page_range = parse_page_range(page_range_spec, None)
        images = pdf_to_images(pdf_path, tmp, page_range=page_range)
        if not images:
            print(f"  ! No se pudo rasterizar {pdf_path.name}", file=sys.stderr)
            return

        numero_archivo = numero_desde_nombre(pdf_path.name)
        doc = Document()
        section = doc.sections[0]
        section.page_width, section.page_height = Twips(11906), Twips(16838)  # A4
        section.top_margin = section.bottom_margin = Twips(700)
        section.left_margin = section.right_margin = Twips(700)

        # Primero se hace el OCR de todas las páginas, para poder numerar
        # ("k/N") las que llegaron sin "Página : ..." legible.
        paginas = []
        for i, img_path in enumerate(images):
            if verbose:
                print(f"  · OCR página {i + 1}/{len(images)}...")
            text = ocr_page(img_path)
            text = strip_noise(text)
            fields, body_text = parse_header(text)
            if fields is not None and numero_archivo:
                fields["numero"] = numero_archivo
            paginas.append((img_path, fields, body_text))
        completar_numeros_de_pagina([f for _, f, _ in paginas])

        for i, (img_path, fields, body_text) in enumerate(paginas):
            if i > 0:
                add_page_break(doc)

            if fields is None:
                # No se reconoció la cabecera -> se inserta la página tal cual
                add_full_page_image(doc, img_path)
                continue

            add_header_table(doc, fields)
            doc.add_paragraph()  # espaciador
            if es_pagina_anexo(body_text):
                add_body_image(doc, img_path)
                continue
            blocks = parse_body(body_text)
            add_content_box(doc, blocks)

        doc.save(out_path)
        if verbose:
            print(f"  -> {out_path}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("entrada", help="Archivo .pdf o carpeta con varios .pdf")
    ap.add_argument("-o", "--salida", help="Ruta del .docx de salida (solo si 'entrada' es un archivo)")
    ap.add_argument("--paginas", help="Rango de páginas del PDF a convertir, ej: 2-3")
    args = ap.parse_args()

    entrada = Path(args.entrada)
    if entrada.is_dir():
        pdfs = sorted(entrada.glob("*.pdf"))
        if not pdfs:
            print("No hay archivos .pdf en esa carpeta.")
            return
        for pdf in pdfs:
            print(f"Procesando {pdf.name}...")
            out = pdf.with_suffix(".docx")
            convert(pdf, out, args.paginas)
    else:
        out = Path(args.salida) if args.salida else entrada.with_suffix(".docx")
        print(f"Procesando {entrada.name}...")
        convert(entrada, out, args.paginas)

    print("Listo. Revisa el/los .docx generados antes de usarlos (el OCR puede tener errores).")


if __name__ == "__main__":
    main()
