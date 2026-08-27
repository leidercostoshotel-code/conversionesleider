/*
 * app.js
 * ======
 * Versión para navegador de pdf_a_word.py: convierte los manuales de
 * procedimientos "SPI Swissôtel Lima" (PDF escaneado) a un .docx editable,
 * enteramente en el cliente (pdf.js + Tesseract.js + docx.js).
 *
 * Reproduce, casi línea por línea, la lógica del script Python original:
 * reconocimiento de la cabecera fija (Área / Campo / Norma / N° / fechas),
 * reconstrucción del cuerpo (Objetivo / Responsables / Pasos a Seguir) y,
 * si una página no encaja en ese patrón, inserción tal cual como imagen.
 */

/* global pdfjsLib, Tesseract, docx */

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, PageBreak, ImageRun,
} = docx;

const FONT = "Arial";
const DPI = 300;
const A4_WIDTH_TWIPS = 11906;
const A4_HEIGHT_TWIPS = 16838;
const MARGIN_TWIPS = 700;
const USABLE_WIDTH_TWIPS = A4_WIDTH_TWIPS - 2 * MARGIN_TWIPS;

// ------------------------------------------------------------------ //
// 2. Reconocer la cabecera fija del manual
// ------------------------------------------------------------------ //

const LEFT_LABELS = [
  ["area", "[AÁ]rea"],
  ["campo", "Campo"],
  ["norma_titulo", "Norma"],
];
const RIGHT_LABELS = [
  ["fecha_elaboracion", "Fecha de elaboraci[oó0]n"],
  ["reemplaza", "Reemplaza a la norma del"],
  ["referencia_corporativa", "Referencia Corporativa"],
  ["elaborada_por", "Elaborada por"],
  ["revisado_aprobado", "Revisado y aprobado por"],
  // Sin "General" al final: ese campo siempre queda en blanco en estos
  // manuales, y el OCR a veces lee mal esa última palabra (p. ej.
  // "General" -> "cabal"). "Autorizado por el Gerente" ya es lo bastante
  // específico para no matchear nada por accidente.
  ["autorizado_por", "Autorizado por el Gerente"],
];
const ALL_LABELS = LEFT_LABELS.concat(RIGHT_LABELS);
const LABEL_RE_SOURCE = ALL_LABELS.map(([k, p]) => `(?<${k}>${p})`).join("|");
const LEFT_KEYS = new Set(LEFT_LABELS.map(([k]) => k));

function cleanValue(s) {
  // Quita el "|" que a veces deja el OCR como ruido del borde de la
  // columna de la tabla (p. ej. "Bono por Fallecimiento de |"), y
  // normaliza espacios repetidos que puedan quedar al quitarlo.
  return s.replace(/\|/g, " ").replace(/\s+/g, " ").trim();
}

const NUMERO_RE = /\b(\d{2}\.\d{2}(?:\.\d{2}){0,2})\b/;
const PAGINA_RE = /P[aá]gina[^\d]{0,40}(\d+\s*\/\s*\d+)/;

function parseHeader(text) {
  const lines = text.split("\n");
  let bodyStartIdx = null;
  for (let i = 0; i < lines.length; i++) {
    if (/OBJETIVO\s*:/i.test(lines[i]) || /PASOS A SEGUIR/i.test(lines[i])) {
      bodyStartIdx = i;
      break;
    }
  }

  if (bodyStartIdx === null) {
    // Página de continuación de un procedimiento de varias páginas: la
    // cabecera se repite en cada página, pero "OBJETIVO:"/"PASOS A SEGUIR"
    // solo aparecen en la primera. Antes, al no encontrar esos marcadores,
    // se descartaba TODO el cuerpo de la página (quedaba solo la cabecera
    // y el resto en blanco). En vez de eso, se ubica el final real de la
    // cabecera: la última línea con alguna etiqueta fija, más alguna línea
    // corta de continuación (valores envueltos a varias líneas, como
    // "Norma" o "Autorizado por...").
    let lastLabelIdx = null;
    for (let i = 0; i < Math.min(lines.length, 15); i++) {
      if (new RegExp(LABEL_RE_SOURCE).test(lines[i])) lastLabelIdx = i;
    }
    if (lastLabelIdx === null) {
      bodyStartIdx = Math.min(10, lines.length);
    } else {
      let idx = lastLabelIdx + 1;
      while (
        idx < lines.length &&
        idx <= lastLabelIdx + 2 &&
        lines[idx].trim() &&
        lines[idx].trim().length <= 50 &&
        !NUM_ITEM_RE.test(lines[idx])
      ) {
        idx++;
      }
      bodyStartIdx = idx;
    }
  }

  const headerLines = lines.slice(0, bodyStartIdx);
  const headerText = headerLines.join("\n");

  const numeroMatch = NUMERO_RE.exec(headerText);
  const paginaMatch = PAGINA_RE.exec(headerText);
  if (!numeroMatch || !paginaMatch) {
    return [null, text]; // no se reconoce cabecera -> esta página va como imagen
  }

  const fields = {
    numero: numeroMatch[1],
    pagina: paginaMatch[1].replace(/\s+/g, ""),
  };

  // Se rastrea la última etiqueta activa de cada columna por separado
  // (izquierda: Área/Campo/Norma; derecha: Fecha/Reemplaza/Elaborada
  // por/Autorizado...). El OCR a veces fusiona ambas columnas en una
  // misma línea física (p. ej. "Norma : ... | Elaborada por : ..."), y
  // cuando eso pasa, el texto de continuación de una columna (el título
  // de "Norma" envuelto a varias líneas) puede terminar en una línea
  // cuya única etiqueta reconocible es de la OTRA columna. En estos
  // manuales, el único campo que se envuelve a varias líneas es "Norma"
  // (columna izquierda), así que ese texto suelto se prioriza para la
  // columna izquierda en vez de perderse o pegarse a la derecha.
  let lastLeftKey = null;
  let lastRightKey = null;
  for (const ln of headerLines) {
    const matches = [...ln.matchAll(new RegExp(LABEL_RE_SOURCE, "g"))];
    if (matches.length === 0) {
      const target = lastLeftKey || lastRightKey;
      if (target) {
        fields[target] = ((fields[target] || "") + " " + ln.trim()).trim();
      }
      continue;
    }

    const firstStart = matches[0].index;
    if (firstStart > 0) {
      const preText = ln.slice(0, firstStart).trim();
      const target = lastLeftKey || lastRightKey;
      if (preText && target) {
        fields[target] = ((fields[target] || "") + " " + preText).trim();
      }
    }

    for (let idx = 0; idx < matches.length; idx++) {
      const m = matches[idx];
      let key = null;
      for (const gk of Object.keys(m.groups || {})) {
        if (m.groups[gk] !== undefined) { key = gk; break; }
      }
      const valStart = m.index + m[0].length;
      const valEnd = idx + 1 < matches.length ? matches[idx + 1].index : ln.length;
      let value = cleanValue(ln.slice(valStart, valEnd).replace(/^[\s:]+/, ""));
      if (key in fields && key !== "area" && key !== "campo") {
        fields[key] = ((fields[key] || "") + " " + value).trim();
      } else {
        fields[key] = value;
      }
      if (LEFT_KEYS.has(key)) {
        lastLeftKey = key;
      } else {
        lastRightKey = key;
      }
    }
  }

  const bodyText = lines.slice(bodyStartIdx).join("\n");
  return [fields, bodyText];
}

// ------------------------------------------------------------------ //
// 3. Reconocer el cuerpo: Objetivo / Responsables / Pasos a Seguir
// ------------------------------------------------------------------ //

const NUM_ITEM_RE = /^\s*(\d{1,2})[.)]\s+(.*)/;
const LETTER_ITEM_RE = /^\s*([a-h])\)\s+(.*)/;
const BULLET_ITEM_RE = /^\s*[•·\-*+]\s+(.*)/;
// El glifo de viñeta (•), en algunos PDF/fuentes, el OCR lo confunde con una
// "e" suelta (ambigüedad visual con el punto de la viñeta). Se trata como
// viñeta solo cuando es una "e" aislada seguida de mayúscula (inicio real de
// oración), para no confundir la conjunción "e" ("madre e hija").
const BULLET_OCR_E_RE = /^\s*e\s+(?=[A-ZÁÉÍÓÚÑ])(.*)/;
// Título de sub-sección numerado sin ":" (p. ej. "1.- RUTINA", "2.- NO
// RUTINA"). Se distingue de un paso numerado real (NUM_ITEM_RE) porque acá
// el punto va seguido de un guión, no de un espacio.
const SECTION_NUM_HEADING_RE = /^\s*\d{1,2}\.-\s+.{1,60}$/;

// Ruido típico de capturas de un visor web (fecha/hora, migas de pan,
// barra de usuario, pie con URL y contador de página) que se mete en el
// OCR cuando el PDF es en realidad una captura de pantalla de un visor.
const NOISE_SUB_PATTERNS = [
  /https?\s*:\s*\/\/\S*[\w/=?.]*/g,
  /\d{1,2}\/\d{1,2}\/\d{2,4},?\s*\d{1,2}:\d{2}\s*[ap]\.?\s*m\.?\.?/gi,
  /\b\d{1,2}\s*\/\s*\d{1,2}\b(?=\s*$)/g,
];
const NOISE_LINE_PATTERNS = [
  /Intranet corporativa/i,
  /Cerrar sesi[oó]n/i,
  /—\s*Intranet\s*$/i,
  /^\s*[←<]\s*\w+\s*[-–—]?\s*\d{0,3}%?\s*[-+]?\s*$/i,
  /^\s*P[aá]g\.?\s*\d+\s*$/i,
];

function stripNoise(text) {
  for (const pat of NOISE_SUB_PATTERNS) {
    text = text.replace(pat, " ");
  }
  const lines = text.split("\n").filter((ln) => !NOISE_LINE_PATTERNS.some((p) => p.test(ln)));
  return lines.join("\n");
}

function isHeadingLine(line) {
  // Título de sección corto, con mayúscula inicial y terminado en ":"
  // (p. ej. "RESPONSABLES:", "BASE LEGAL:", "CARACTERISTICAS:"). Se exige
  // además que sean pocas palabras (<=4): los títulos reales de este tipo
  // de manuales son de 1-3 palabras, a diferencia de una oración normal
  // que también puede terminar en ":" antes de una lista (p. ej. "El
  // formato se distribuye como sigue:"), que no es un título aunque
  // cumpla lo demás.
  const trimmed = line.trim();
  return (
    trimmed.length > 0 &&
    trimmed.length <= 45 &&
    trimmed.split(/\s+/).length <= 4 &&
    trimmed.endsWith(":") &&
    trimmed[0] === trimmed[0].toUpperCase() &&
    trimmed[0] !== trimmed[0].toLowerCase()
  );
}

function splitSection(text, startPat, endPats) {
  // Devuelve [prefijo_antes_de_la_marca, valor_de_la_sección, resto].
  // `prefijo` es lo que había ANTES de encontrar `startPat` -- normalmente
  // vacío, salvo cuando hay una sección intermedia no prevista (p. ej.
  // "BASE LEGAL:" entre "OBJETIVO:" y "RESPONSABLES:") que de otro modo se
  // perdería en silencio. Si no se encuentra `startPat`, devuelve
  // [null, null, text] sin consumir nada.
  const m = new RegExp(startPat, "i").exec(text);
  if (!m) return [null, null, text];
  const prefix = text.slice(0, m.index);
  const rest = text.slice(m.index + m[0].length);
  let endPos = rest.length;
  for (const ep of endPats) {
    const em = new RegExp(ep, "i").exec(rest);
    if (em && em.index < endPos) endPos = em.index;
  }
  // Además de los patrones explícitos (que solo cubren los nombres de
  // sección esperados), cualquier línea que en sí misma parezca un título
  // de sección marca el final de este bloque. Esto generaliza a secciones
  // no previstas (p. ej. "BASE LEGAL:", "CARACTERISTICAS:") que de otro
  // modo quedarían arrastradas dentro de "OBJETIVO" o "RESPONSABLES" hasta
  // la siguiente sección conocida.
  let pos = 0;
  for (const line of rest.split("\n")) {
    if (pos >= endPos) break;
    if (pos > 0 && isHeadingLine(line)) {
      endPos = pos;
      break;
    }
    pos += line.length + 1; // +1 por el "\n" separador
  }
  return [prefix, rest.slice(0, endPos).trim(), rest.slice(endPos)];
}

function cleanParagraph(text) {
  return text.split("\n").map((ln) => ln.trim()).filter(Boolean).join(" ");
}

function stripChars(s, chars) {
  const esc = chars.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`^[${esc}]+|[${esc}]+$`, "g");
  return s.replace(re, "");
}

function parseBody(bodyText) {
  const blocks = [];
  let remaining = bodyText;

  let preObjetivo, objetivo;
  [preObjetivo, objetivo, remaining] = splitSection(
    remaining, "OBJETIVO\\s*:", ["RESPONSABLES\\s*:", "PASOS A SEGUIR", "POL[IÍ]TICAS"]
  );
  if (preObjetivo && preObjetivo.trim()) blocks.push(...parseSteps(preObjetivo));
  if (objetivo) blocks.push({ type: "objetivo", text: cleanParagraph(objetivo) });

  let preResponsables, responsables;
  [preResponsables, responsables, remaining] = splitSection(
    remaining, "RESPONSABLES\\s*:", ["PASOS A SEGUIR", "POL[IÍ]TICAS"]
  );
  if (preResponsables && preResponsables.trim()) {
    // Sección intermedia no prevista (p. ej. "BASE LEGAL:") entre
    // "OBJETIVO:" y "RESPONSABLES:": se reconstruye con el parser
    // genérico en vez de perderla.
    blocks.push(...parseSteps(preResponsables));
  }
  if (responsables) {
    const names = responsables.split("\n")
      .map((ln) => ln.trim())
      .filter(Boolean)
      .map((ln) => stripChars(ln, " ."));
    blocks.push({ type: "responsables", names });
  }

  let stepsText = remaining;
  const m = /PASOS A SEGUIR\s*:?/i.exec(stepsText);
  if (m) stepsText = stepsText.slice(m.index + m[0].length);

  blocks.push(...parseSteps(stepsText));
  return blocks;
}

function parseSteps(text) {
  const lines = text.split("\n").map((ln) => ln.trim()).filter(Boolean);
  const blocks = [];
  let current = null;
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    let m = NUM_ITEM_RE.exec(line);
    if (m) {
      if (current) blocks.push(current);
      current = { type: "numbered", n: m[1], text: m[2].trim() };
      i++; continue;
    }

    m = LETTER_ITEM_RE.exec(line);
    if (m) {
      if (current) blocks.push(current);
      current = { type: "letter", letter: m[1], text: m[2].trim() };
      i++; continue;
    }

    m = BULLET_ITEM_RE.exec(line) || BULLET_OCR_E_RE.exec(line);
    if (m) {
      if (current) blocks.push(current);
      current = { type: "bullet", text: m[1].trim() };
      i++; continue;
    }

    // Un título de sección corto, en mayúscula inicial y terminado en ":"
    // (p. ej. "CONTROLES:", "REMUNERACIONES:", "PAGOS:") es en sí mismo una
    // señal inequívoca de subtítulo en este tipo de manuales -- no hace
    // falta que además reinicie una lista en "1.".
    if (isHeadingLine(line) || SECTION_NUM_HEADING_RE.test(line)) {
      if (current) { blocks.push(current); current = null; }
      blocks.push({ type: "subheading", text: line });
      i++; continue;
    }

    const looksLikeHeading =
      line.length <= 45 &&
      line.split(/\s+/).length <= 4 &&
      !/[.,;]$/.test(line) &&
      line.length > 0 &&
      line[0] === line[0].toUpperCase() &&
      line[0] !== line[0].toLowerCase();

    // Un subtítulo sin ":" casi siempre viene seguido de un ítem "1."
    // (la lista de esa sub-sección vuelve a empezar). Usamos esa
    // señal aunque haya un ítem numerado abierto en ese momento.
    let nextStartsListAt1 = false;
    if (i + 1 < lines.length) {
      const nm = NUM_ITEM_RE.exec(lines[i + 1]);
      nextStartsListAt1 = !!(nm && nm[1] === "1");
    }

    if (looksLikeHeading && (current === null || nextStartsListAt1)) {
      if (current) { blocks.push(current); current = null; }
      blocks.push({ type: "subheading", text: line });
      i++; continue;
    }

    if (current) {
      current.text = (current.text + " " + line).trim();
    } else {
      // OJO: se deja `current` apuntando a este párrafo (no se hace push
      // todavía) para que la siguiente línea, si es una continuación
      // envuelta del mismo párrafo, se una aquí en vez de crear un bloque
      // nuevo por cada línea física del OCR.
      current = { type: "paragraph", text: line };
    }
    i++;
  }

  if (current) blocks.push(current);
  return blocks;
}

// ------------------------------------------------------------------ //
// 4. Construcción del .docx (mismo diseño que el original)
// ------------------------------------------------------------------ //

const CELL_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
};

function makeParagraph(text, { bold = false, size = 10, indent = 0 } = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold, font: FONT, size: size * 2 })],
    spacing: { before: 0, after: 100 },
    indent: indent ? { left: indent } : undefined,
  });
}

function buildHeaderTable(f) {
  const cell00 = new TableCell({
    width: { size: 6800, type: WidthType.DXA },
    borders: CELL_BORDERS,
    children: [
      makeParagraph("SPI SWISSÔTEL LIMA", { bold: true, size: 11 }),
      makeParagraph("Manual de Procedimientos", { size: 10 }),
    ],
  });
  const cell01 = new TableCell({
    width: { size: 3100, type: WidthType.DXA },
    borders: CELL_BORDERS,
    children: [
      makeParagraph(`N°  : ${f.numero || ""}`, { bold: true, size: 10 }),
      makeParagraph(`Página: ${f.pagina || ""}`, { size: 10 }),
    ],
  });

  const cell10Children = [makeParagraph(`Area  : ${f.area || ""}`, { size: 9 })];
  for (const [key, label] of [["campo", "Campo"], ["norma_titulo", "Norma"]]) {
    cell10Children.push(makeParagraph(`${label} : ${f[key] || ""}`, { size: 9 }));
  }
  const cell10 = new TableCell({
    width: { size: 6800, type: WidthType.DXA },
    borders: CELL_BORDERS,
    children: cell10Children,
  });

  const rightFieldsDef = [
    ["fecha_elaboracion", "Fecha de elaboración"],
    ["reemplaza", "Reemplaza a la norma del"],
    ["elaborada_por", "Elaborada por"],
    ["referencia_corporativa", "Referencia Corporativa SL"],
    ["revisado_aprobado", "Revisado y aprobado por"],
  ];
  const cell11Children = [];
  for (const [key, label] of rightFieldsDef) {
    if (!f[key]) continue;
    cell11Children.push(makeParagraph(`${label} : ${f[key]}`, { size: 9 }));
  }
  cell11Children.push(makeParagraph("Autorizado por el Gerente General :", { size: 9 }));
  const cell11 = new TableCell({
    width: { size: 3100, type: WidthType.DXA },
    borders: CELL_BORDERS,
    children: cell11Children,
  });

  return new Table({
    width: { size: 9900, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ children: [cell00, cell01] }),
      new TableRow({ children: [cell10, cell11] }),
    ],
  });
}

function buildContentBox(blocks) {
  const children = [];

  for (const b of blocks) {
    switch (b.type) {
      case "objetivo":
        children.push(makeParagraph("OBJETIVO:", { bold: true }));
        children.push(makeParagraph(b.text));
        break;
      case "responsables":
        children.push(makeParagraph("RESPONSABLES:", { bold: true }));
        for (const name of b.names) children.push(makeParagraph(name));
        break;
      case "subheading":
        children.push(makeParagraph(b.text, { bold: true }));
        break;
      case "numbered":
        children.push(makeParagraph(`${b.n}.  ${b.text}`, { indent: 360 }));
        break;
      case "letter":
        children.push(makeParagraph(`${b.letter})  ${b.text}`, { indent: 500 }));
        break;
      case "bullet":
        children.push(makeParagraph(`•  ${b.text}`, { indent: 360 }));
        break;
      default: // paragraph
        children.push(makeParagraph(b.text));
    }
  }
  if (children.length === 0) children.push(makeParagraph(""));

  const cell = new TableCell({
    width: { size: 9900, type: WidthType.DXA },
    borders: CELL_BORDERS,
    children,
  });

  return new Table({
    width: { size: 9900, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({ children: [cell] })],
  });
}

async function buildFullPageImage(canvas) {
  const dataUrl = canvas.toDataURL("image/png");
  const res = await fetch(dataUrl);
  const arrayBuffer = await res.arrayBuffer();
  const widthPx = Math.round((USABLE_WIDTH_TWIPS / 1440) * 96);
  const heightPx = Math.round(widthPx * (canvas.height / canvas.width));
  return new Paragraph({
    children: [new ImageRun({ data: arrayBuffer, transformation: { width: widthPx, height: heightPx } })],
  });
}

// ------------------------------------------------------------------ //
// 1. PDF -> imágenes (pdf.js), imagen -> texto (Tesseract.js)
// ------------------------------------------------------------------ //

async function renderPageToCanvas(pdf, pageNumber, dpi = DPI) {
  const page = await pdf.getPage(pageNumber);
  const scale = dpi / 72;
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext("2d");
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas;
}

function parsePageRange(spec) {
  if (!spec) return null;
  const trimmed = spec.trim();
  if (!trimmed) return null;
  if (trimmed.includes("-")) {
    const [a, b] = trimmed.split("-", 2);
    return [parseInt(a, 10), parseInt(b, 10)];
  }
  const n = parseInt(trimmed, 10);
  return [n, n];
}

// ------------------------------------------------------------------ //
// 5. Orquestación
// ------------------------------------------------------------------ //

let ocrWorker = null;
async function getOcrWorker(onLog) {
  if (!ocrWorker) {
    onLog("Cargando motor de OCR (español)...");
    ocrWorker = await Tesseract.createWorker("spa");
    // "single column" en vez de "single block": con SINGLE_BLOCK, Tesseract
    // tiende a fusionar las dos columnas de la tabla de cabecera en una sola
    // línea por fila, perdiendo la línea donde está "Página : X/Y" y haciendo
    // que la página caiga al modo imagen aunque el formato sea reconocible.
    await ocrWorker.setParameters({ tessedit_pageseg_mode: Tesseract.PSM.SINGLE_COLUMN });
  }
  return ocrWorker;
}

async function convertPdf(file, pageRangeSpec, onLog) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const range = parsePageRange(pageRangeSpec);
  const firstPage = range ? Math.max(1, range[0]) : 1;
  const lastPage = range ? Math.min(pdf.numPages, range[1]) : pdf.numPages;

  const worker = await getOcrWorker(onLog);

  const allChildren = [];
  const totalPages = lastPage - firstPage + 1;

  for (let p = firstPage; p <= lastPage; p++) {
    const idx = p - firstPage;
    onLog(`OCR página ${idx + 1}/${totalPages} (página ${p} del PDF)...`);

    const canvas = await renderPageToCanvas(pdf, p);
    const { data: { text: rawText } } = await worker.recognize(canvas);
    const text = stripNoise(rawText);
    const [fields, bodyText] = parseHeader(text);

    if (idx > 0) {
      allChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }

    if (fields === null) {
      onLog(`  ⚠ No se reconoció la cabecera en la página ${p} -> se insertó como imagen.`);
      onLog(`  Texto OCR detectado en esa página:`);
      const preview = text.trim() || "(vacío)";
      const truncated = preview.length > 1000 ? preview.slice(0, 1000) + " …(truncado)" : preview;
      for (const line of truncated.split("\n")) {
        onLog(`    │ ${line}`);
      }
      allChildren.push(await buildFullPageImage(canvas));
      continue;
    }

    onLog(`  ✓ Cabecera reconocida (N° ${fields.numero}, página ${fields.pagina}).`);

    allChildren.push(buildHeaderTable(fields));
    allChildren.push(new Paragraph({ text: "" })); // espaciador
    const blocks = parseBody(bodyText);
    allChildren.push(buildContentBox(blocks));
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: A4_WIDTH_TWIPS, height: A4_HEIGHT_TWIPS },
          margin: { top: MARGIN_TWIPS, bottom: MARGIN_TWIPS, left: MARGIN_TWIPS, right: MARGIN_TWIPS },
        },
      },
      children: allChildren,
    }],
  });

  return Packer.toBlob(doc);
}

// ------------------------------------------------------------------ //
// UI
// ------------------------------------------------------------------ //

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const dropLabel = document.getElementById("dropLabel");
const convertBtn = document.getElementById("convertBtn");
const pageRangeInput = document.getElementById("pageRange");
const logEl = document.getElementById("log");
const downloadsEl = document.getElementById("downloads");
const downloadListEl = document.getElementById("downloadList");

let selectedFiles = [];

function updateFileLabel() {
  if (selectedFiles.length === 0) {
    dropLabel.textContent = "Selecciona uno o varios PDF, o arrástralos aquí";
    convertBtn.disabled = true;
  } else if (selectedFiles.length === 1) {
    dropLabel.textContent = selectedFiles[0].name;
    convertBtn.disabled = false;
  } else {
    dropLabel.textContent = `${selectedFiles.length} archivos seleccionados`;
    convertBtn.disabled = false;
  }
}

fileInput.addEventListener("change", () => {
  selectedFiles = Array.from(fileInput.files || []);
  updateFileLabel();
});

["dragover", "dragenter"].forEach((evt) => {
  dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach((evt) => {
  dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
  });
});
dropZone.addEventListener("drop", (e) => {
  const files = Array.from(e.dataTransfer.files || []).filter((f) => f.type === "application/pdf");
  if (files.length) {
    selectedFiles = files;
    updateFileLabel();
  }
});

function fileLogSection(name) {
  const wrap = document.createElement("div");
  wrap.className = "file-log";
  const h3 = document.createElement("h3");
  h3.textContent = name;
  wrap.appendChild(h3);
  logEl.appendChild(wrap);
  return {
    line(text, isError = false) {
      const p = document.createElement("div");
      p.className = "line" + (isError ? " error" : "");
      p.textContent = text;
      wrap.appendChild(p);
    },
  };
}

convertBtn.addEventListener("click", async () => {
  if (selectedFiles.length === 0) return;

  convertBtn.disabled = true;
  logEl.innerHTML = "";
  downloadListEl.innerHTML = "";
  downloadsEl.hidden = true;

  const pageRangeSpec = pageRangeInput.value.trim();

  for (const file of selectedFiles) {
    const log = fileLogSection(file.name);
    try {
      log.line("Procesando...");
      const blob = await convertPdf(file, pageRangeSpec, (msg) => log.line(msg));
      const outName = file.name.replace(/\.pdf$/i, "") + ".docx";
      const url = URL.createObjectURL(blob);

      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = url;
      a.download = outName;
      a.textContent = `Descargar ${outName}`;
      li.appendChild(a);
      downloadListEl.appendChild(li);
      downloadsEl.hidden = false;

      log.line(`Listo -> ${outName}`);
    } catch (err) {
      console.error(err);
      log.line(`Error: ${err.message || err}`, true);
    }
  }

  const finalLine = document.createElement("div");
  finalLine.className = "line";
  finalLine.textContent = "Listo. Revisa el/los .docx generados antes de usarlos (el OCR puede tener errores).";
  logEl.appendChild(finalLine);

  convertBtn.disabled = false;
});
