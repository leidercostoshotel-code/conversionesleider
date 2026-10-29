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

// Valores fijos que se escriben SIEMPRE en el .docx de salida, en vez de lo
// que diga el PDF original (pedido del usuario):
// - "Fecha de elaboración" es la fecha en que se hace la conversión.
// - "Reemplaza a la norma del" baja la fecha de elaboración del PDF original
//   (o, si no se leyó, la fecha de "Reemplaza" original).
// - La elabora JACH y la revisa GPL (línea "Revisado por" nueva).
// - "Contralor" y "Sub Contralor" van siempre primero en RESPONSABLES.
const ELABORADA_POR_FIJO = "JACH";
const REVISADO_POR_FIJO = "GPL";
const RESPONSABLES_FIJOS = ["Contralor", "Sub Contralor"];
// Cargos que van juntos en la lista de responsables (pedido del usuario):
// si aparece uno de ellos, aparecen todos, en este orden y en el lugar del
// primero que estaba. Se comparan sin mayúsculas ni guiones.
const PARES_RESPONSABLES = [
  ["Contralor", "Sub Contralor"],
  ["Director de Operaciones", "Gerente de Operaciones"],
];

function claveResponsable(n) {
  return n.trim().toLowerCase().replace(/-/g, " ").replace(/\s+/g, " ");
}

// Normaliza la lista de responsables: completa los cargos que van juntos,
// quita duplicados y, si `fijosPrimero`, antepone siempre Contralor y
// Sub Contralor.
function normalizarResponsables(names, fijosPrimero) {
  let lista = names.map((n) => n.trim()).filter(Boolean);
  for (const grupo of PARES_RESPONSABLES) {
    const claves = grupo.map(claveResponsable);
    const idx = lista.findIndex((n) => claves.includes(claveResponsable(n)));
    if (idx < 0) continue;
    lista = [...lista.slice(0, idx), ...grupo, ...lista.slice(idx).filter((n) => !claves.includes(claveResponsable(n)))];
  }
  if (fijosPrimero) {
    const fijos = RESPONSABLES_FIJOS.map(claveResponsable);
    lista = [...RESPONSABLES_FIJOS, ...lista.filter((n) => !fijos.includes(claveResponsable(n)))];
  }
  const vistos = new Set();
  return lista.filter((n) => !vistos.has(claveResponsable(n)) && vistos.add(claveResponsable(n)));
}

function fechaDeHoy() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getFullYear()}`;
}
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
// Fila de la tabla de cabecera a la que pertenece cada etiqueta de la
// derecha -> campo de la izquierda de esa misma fila. Sirve cuando el OCR se
// come la etiqueta izquierda ("Area :", "Campo :", "Norma :") y solo queda su
// valor suelto delante de la etiqueta derecha.
const ROW_LEFT_KEY = {
  fecha_elaboracion: "area",
  reemplaza: "campo",
  elaborada_por: "norma_titulo",
  autorizado_por: "norma_titulo",
};

function matchKey(m) {
  for (const gk of Object.keys(m.groups || {})) {
    if (m.groups[gk] !== undefined) return gk;
  }
  return null;
}

function cleanValue(s) {
  // Quita el "|" que a veces deja el OCR como ruido del borde de la
  // columna de la tabla (p. ej. "Bono por Fallecimiento de |"), y
  // normaliza espacios repetidos que puedan quedar al quitarlo.
  return s.replace(/[|\[\]]/g, " ").replace(/[\s\-–—]+$/, "").replace(/\s+/g, " ").trim();
}

const NUMERO_RE = /\b(\d{2}\.\d{2}(?:\.\d{2}){0,2})\b/;
// El N° también viene en el nombre del archivo ("SPI 04.01.13.02 Despacho...").
// Ese valor es exacto, mientras que el OCR a veces cambia el último dígito,
// así que cuando está en el nombre manda sobre lo leído de la imagen.
const FILENAME_NUM_RE = /(?<!\d)(\d{2}\.\d{2}(?:\.\d{2}){0,2})(?!\d)/;

function numeroDesdeNombre(nombre) {
  const m = FILENAME_NUM_RE.exec(nombre || "");
  return m ? m[1] : null;
}
// La barra de "2/3" es opcional: Tesseract.js a veces se la come y lee
// "Página : 23". Como va justo detrás de la etiqueta "Página", esos dígitos
// pegados son casi seguro "página/total"; ver normalizarPagina().
// "P?" porque con el render de pdf.js Tesseract a veces se come la "P"
// inicial y lee "agina: 1/1".
const PAGINA_RE = /P?[aáé]gina[^\d]{0,40}(\d{1,2}\s*[/\\|]?\s*\d{1,2})/;

function normalizarPagina(raw) {
  const s = raw.replace(/\s+/g, "").replace(/[\\|]/g, "/");
  if (s.includes("/")) return s;
  // Sin separador: "23" -> "2/3", "110" -> "1/10", "1013" -> "10/13".
  if (s.length === 2) return `${s[0]}/${s[1]}`;
  if (s.length === 3) return `${s[0]}/${s.slice(1)}`;
  if (s.length === 4) return `${s.slice(0, 2)}/${s.slice(2)}`;
  return s;
}

function parseHeader(text) {
  const lines = text.split("\n");

  // Marcador de inicio de cuerpo ("OBJETIVO(S):" / "PASOS A SEGUIR"), solo
  // en MAYÚSCULAS (así están siempre los títulos de sección; buscarlo sin
  // distinguir mayúsculas confundía "los pasos a seguir" dentro de una
  // oración normal) y solo dentro de las primeras líneas: en una página de
  // continuación puede aparecer un "OBJETIVOS:" a mitad de página (empieza
  // un sub-procedimiento nuevo) y NO marca el fin de la cabecera.
  let markerIdx = null;
  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    if (/OBJETIVOS?\s*:/.test(lines[i]) || /PASOS A SEGUIR/.test(lines[i])) {
      markerIdx = i;
      break;
    }
  }

  // El fin real de la cabecera lo deciden las etiquetas fijas: la última
  // línea con alguna etiqueta, más hasta 2 líneas cortas de continuación
  // (valores envueltos a varias líneas, como "Norma" o "Autorizado por").
  // Así, un título del cuerpo que venga justo después de la cabecera y
  // antes de "OBJETIVOS:" (p. ej. "GASTOS DE CAPITAL") no se pega al campo
  // "Norma". El marcador, si existe, solo actúa como tope.
  let lastLabelIdx = null;
  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    if (new RegExp(LABEL_RE_SOURCE).test(lines[i])) lastLabelIdx = i;
  }
  let bodyStartIdx;
  if (lastLabelIdx === null) {
    bodyStartIdx = markerIdx !== null ? markerIdx : Math.min(10, lines.length);
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
    if (markerIdx !== null) bodyStartIdx = Math.min(bodyStartIdx, markerIdx);
  }

  const headerLines = lines.slice(0, bodyStartIdx);
  const headerText = headerLines.join("\n");

  const numeroMatch = NUMERO_RE.exec(headerText);
  const paginaMatch = PAGINA_RE.exec(headerText);
  // En algunos PDF el OCR se come la línea "Página : 1/1" entera. Si aun así
  // hay N° y alguna etiqueta inequívoca del cuadro (Fecha de elaboración /
  // Elaborada por / Manual de Procedimientos), la cabecera se acepta igual y
  // el número de página se completa después contando las páginas
  // reconocidas (ver convertPdf()).
  const otrasEtiquetas = /Fecha de elaboraci|Elaborada por|Manual de Procedimientos/i.test(headerText);
  if (!numeroMatch || !(paginaMatch || otrasEtiquetas)) {
    return [null, text]; // no se reconoce cabecera -> esta página va como imagen
  }

  const fields = {
    numero: numeroMatch[1],
    pagina: paginaMatch ? normalizarPagina(paginaMatch[1]) : "",
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
      const preText = cleanValue(ln.slice(0, firstStart).replace(/^[\s:\-–—]+/, ""));
      const rowLeft = ROW_LEFT_KEY[matchKey(matches[0])];
      if (preText && rowLeft && !fields[rowLeft]) {
        // La etiqueta izquierda se perdió en el OCR (con el render de pdf.js,
        // Tesseract omite "Area :" y deja solo
        // "- Contraloría Fecha de elaboración - 14.03.15"): el texto previo a
        // la etiqueta derecha es el valor izquierdo de esa misma fila.
        fields[rowLeft] = preText;
        lastLeftKey = rowLeft;
      } else {
        const target = lastLeftKey || lastRightKey;
        if (preText && target) {
          fields[target] = ((fields[target] || "") + " " + preText).trim();
        }
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
      // También se quita un guión inicial: el OCR del navegador a veces lee
      // el ":" de la tabla como "-" ("Elaborada por - GPL").
      let value = cleanValue(ln.slice(valStart, valEnd).replace(/^[\s:\-–—]+/, ""));
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

  // Fecha sin puntos ("160412"): el OCR a veces se come los separadores.
  // Fecha sin puntos o con puntos perdidos ("160412", "2302.15", "23.0910"):
  // si quedan exactamente 6 dígitos, se reescribe como DD.MM.AA.
  const fe = fields.fecha_elaboracion || "";
  const feDigits = fe.replace(/\D/g, "");
  if (feDigits.length === 6 && /^[\d.\s]+$/.test(fe)) {
    fields.fecha_elaboracion = `${feDigits.slice(0, 2)}.${feDigits.slice(2, 4)}.${feDigits.slice(4)}`;
  }

  const bodyText = lines.slice(bodyStartIdx).join("\n");
  return [fields, bodyText];
}

// ------------------------------------------------------------------ //
// 3. Reconocer el cuerpo: Objetivo / Responsables / Pasos a Seguir
// ------------------------------------------------------------------ //

// Paso numerado: "1. ", "1) ", y también las variantes que trae el OCR o el
// propio PDF: "1.- ", "1- ", "7.-Estos" (sin espacio). Lo que sigue al
// marcador debe empezar con letra o comilla, para no confundirlo con una
// cifra como "75.00".
const NUM_ITEM_RE = /^\s*(\d{1,2})\s*(?:\.?-+|[.)])\s*(?=[A-Za-zÁÉÍÓÚÑÜáéíóúñü“"(¿¡])(.*)/;
const LETTER_ITEM_RE = /^\s*([a-h])\)\s+(.*)/;
// Además de los glifos reales de viñeta, se aceptan "+" y "=": así lee el
// OCR, respectivamente, al "•" pequeño y al "▪" cuadrado de algunos PDF.
// El "." al inicio también cuenta: es como el navegador suele leer la viñeta.
const BULLET_ITEM_RE = /^\s*[•·\-*+=.]\s+(.*)/;
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
  /Documento de uso interno/i,
];

// Línea de "basura" del OCR: el navegador convierte las franjas grises de
// la intranet en ristras como "ITTiaIl ol iTíimig1aíditGi". En ellas más de
// la mitad de las letras son palos verticales (I, l, i, T, 1, í, Í).
function isGarbageLine(line) {
  const letters = line.match(/[\p{L}\p{N}]/gu) || [];
  if (letters.length < 8) return false;
  const palos = letters.filter((c) => "IliTt1íÍ|".includes(c)).length;
  return palos / letters.length > 0.5;
}

function stripNoise(text) {
  for (const pat of NOISE_SUB_PATTERNS) {
    text = text.replace(pat, " ");
  }
  const lines = text.split("\n")
    .filter((ln) => !NOISE_LINE_PATTERNS.some((p) => p.test(ln)))
    .filter((ln) => !isGarbageLine(ln));
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
  // [null, null, text, null] sin consumir nada. El cuarto valor es el texto
  // del marcador tal como se encontró (p. ej. "OBJETIVOS:").
  //
  // Los marcadores se buscan anclados al inicio de línea y en MAYÚSCULAS
  // (así aparecen siempre en estos manuales): buscarlos sin distinguir
  // mayúsculas hacía que "los pasos a seguir" dentro de una oración normal
  // cortara la sección "OBJETIVO" por la mitad.
  const m = new RegExp(startPat, "m").exec(text);
  if (!m) return [null, null, text, null];
  const prefix = text.slice(0, m.index);
  const rest = text.slice(m.index + m[0].length);
  let endPos = rest.length;
  for (const ep of endPats) {
    const em = new RegExp(ep, "m").exec(rest);
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
  return [prefix, rest.slice(0, endPos).trim(), rest.slice(endPos), m[0].trim()];
}

function hasListItems(text) {
  return text.split("\n").some((ln) =>
    NUM_ITEM_RE.test(ln) || LETTER_ITEM_RE.test(ln) ||
    BULLET_ITEM_RE.test(ln) || BULLET_OCR_E_RE.test(ln)
  );
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

  let preObjetivo, objetivo, objetivoMarker;
  [preObjetivo, objetivo, remaining, objetivoMarker] = splitSection(
    remaining, "^\\s*OBJETIVOS?\\s*:", ["^\\s*RESPONSABLES?\\s*:", "^\\s*PASOS A SEGUIR", "^\\s*POL[IÍ]TICAS"]
  );
  if (preObjetivo && preObjetivo.trim()) blocks.push(...parseSteps(preObjetivo));
  if (objetivoMarker && !objetivo) {
    // El marcador está al final de la página y su contenido sigue en la
    // siguiente: se conserva el título igual.
    blocks.push({ type: "subheading", text: objetivoMarker });
  } else if (objetivo) {
    if (hasListItems(objetivo)) {
      // Algunos manuales ponen el objetivo como lista de viñetas (p. ej.
      // bajo "OBJETIVOS:"): se conserva esa estructura en vez de
      // aplastarla en un solo párrafo.
      blocks.push({ type: "subheading", text: objetivoMarker });
      blocks.push(...parseSteps(objetivo));
    } else {
      blocks.push({ type: "objetivo", text: cleanParagraph(objetivo) });
    }
  }

  let preResponsables, responsables;
  [preResponsables, responsables, remaining] = splitSection(
    remaining, "^\\s*RESPONSABLES?\\s*:", ["^\\s*PASOS A SEGUIR", "^\\s*POL[IÍ]TICAS"]
  );
  if (preResponsables && preResponsables.trim()) {
    // Sección intermedia no prevista (p. ej. "BASE LEGAL:") entre
    // "OBJETIVO:" y "RESPONSABLES:": se reconstruye con el parser
    // genérico en vez de perderla.
    blocks.push(...parseSteps(preResponsables));
  }
  if (responsables) {
    // La lista de nombres termina en el primer título que aparezca (p. ej.
    // "COSTO POR SERVICIO SEGÚN CONTRATO:"); lo que sigue se reconstruye
    // con el parser genérico.
    const lineas = responsables.split("\n").map((ln) => ln.trim()).filter(Boolean);
    let corte = lineas.length;
    for (let k = 0; k < lineas.length; k++) {
      const ln = lineas[k];
      const esTitulo = isHeadingLine(ln) ||
        (ln.length <= 45 && ln === ln.toUpperCase() && /[A-ZÁÉÍÓÚÑ]/.test(ln) && !/[.,;]$/.test(ln));
      if (esTitulo && k > 0 && ln.length >= 3) { corte = k; break; }
    }
    const names = lineas.slice(0, corte).map((ln) => stripChars(ln, " ."));
    blocks.push({ type: "responsables", names });
    if (corte < lineas.length) blocks.push(...parseSteps(lineas.slice(corte).join("\n")));
  }

  let stepsText = remaining;
  const m = /^\s*PASOS A SEGUIR\s*:?/m.exec(stepsText);
  if (m) {
    // Lo que haya entre la última sección reconocida y "PASOS A SEGUIR"
    // (p. ej. un "RESPONSABLE:" con un nombre que no coincidió con el
    // patrón esperado) se reconstruye con el parser genérico en vez de
    // descartarse.
    const before = stepsText.slice(0, m.index);
    if (before.trim()) blocks.push(...parseSteps(before));
    // El título "PASOS A SEGUIR:" también va al Word, como en el PDF.
    blocks.push({ type: "subheading", text: "PASOS A SEGUIR:" });
    stepsText = stepsText.slice(m.index + m[0].length);
  }

  blocks.push(...parseSteps(stepsText));
  return blocks;
}

function parseSteps(text) {
  // Se conserva si cada línea venía precedida de una línea en blanco: el
  // OCR separa los párrafos/ítems con un renglón vacío, y esa es la única
  // pista que queda cuando el navegador se come las viñetas.
  const lines = [];
  const blankBefore = [];
  let pendingBlank = false;
  for (const raw of text.split("\n")) {
    const ln = raw.trim();
    if (!ln) { pendingBlank = true; continue; }
    lines.push(ln);
    blankBefore.push(pendingBlank);
    pendingBlank = false;
  }
  const blocks = [];
  let current = null;
  let prevLen = 0; // largo de la última línea física procesada
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const thisLen = line.length;

    let m = NUM_ITEM_RE.exec(line);
    // "1.- RUTINA" (texto corto en mayúsculas tras el "N.-") es un subtítulo
    // numerado, no un paso: se deja para la regla de títulos.
    const tituloNumerado = Boolean(m && SECTION_NUM_HEADING_RE.test(line) && (m[2].trim() === m[2].trim().toUpperCase() && /[A-ZÁÉÍÓÚÑ]/.test(m[2])));
    if (m && !tituloNumerado) {
      if (current) blocks.push(current);
      current = { type: "numbered", n: m[1], text: m[2].trim() };
      prevLen = thisLen;
      i++; continue;
    }

    m = LETTER_ITEM_RE.exec(line);
    if (m) {
      if (current) blocks.push(current);
      current = { type: "letter", letter: m[1], text: m[2].trim() };
      prevLen = thisLen;
      i++; continue;
    }

    m = BULLET_ITEM_RE.exec(line) || BULLET_OCR_E_RE.exec(line);
    if (m) {
      if (current) blocks.push(current);
      current = { type: "bullet", text: m[1].trim() };
      prevLen = thisLen;
      i++; continue;
    }

    // Un título de sección corto, en mayúscula inicial y terminado en ":"
    // (p. ej. "CONTROLES:", "REMUNERACIONES:", "PAGOS:") es en sí mismo una
    // señal inequívoca de subtítulo en este tipo de manuales -- no hace
    // falta que además reinicie una lista en "1.".
    // Una línea corta TODA en mayúsculas (p. ej. "DISPOSICIÓN Y CASTIGO DE
    // LOS ACTIVOS") es título de sección aunque pase de 4 palabras.
    const allCapsHeading =
      line.length <= 45 &&
      line === line.toUpperCase() &&
      line !== line.toLowerCase() &&
      !/[.,;]$/.test(line);
    if (isHeadingLine(line) || SECTION_NUM_HEADING_RE.test(line) || allCapsHeading) {
      if (current) { blocks.push(current); current = null; }
      blocks.push({ type: "subheading", text: line });
      prevLen = thisLen;
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

    // Un renglón corto aislado (línea en blanco antes y después) es un
    // subtítulo aunque haya un párrafo abierto (p. ej. "General").
    // Se exige además que el párrafo anterior haya terminado en ".": si
    // terminó en ":" o sin puntuación, lo que sigue es una lista.
    const aislado =
      blankBefore[i] && i + 1 < lines.length && blankBefore[i + 1] &&
      (current === null || (current.text || "").endsWith("."));
    if (looksLikeHeading && (current === null || nextStartsListAt1 || aislado)) {
      if (current) { blocks.push(current); current = null; }
      blocks.push({ type: "subheading", text: line });
      prevLen = thisLen;
      i++; continue;
    }

    // Con el render del navegador, el OCR a veces se come por completo los
    // glifos de viñeta y los ítems llegan sin marcador. Dentro de un párrafo
    // genérico, una línea que empieza en mayúscula justo después de una que
    // terminó en ".", ":" o ";" es un ítem nuevo, no la continuación
    // envuelta del anterior.
    // También es ítem nuevo si venía separado por una línea en blanco o si
    // la línea anterior era corta (una línea envuelta ocupa casi todo el
    // ancho; una corta es el final de su párrafo).
    // (Una línea que empieza con cifra, p. ej. "24 Horas 30.00", también es
    // ítem nuevo si viene tras línea en blanco o tras línea corta.)
    const empiezaMayus = line[0] === line[0].toUpperCase() && line[0] !== line[0].toLowerCase();
    const empiezaMayusOCifra = empiezaMayus || /^\d/.test(line);
    const nuevoItem =
      current !== null &&
      ((current.type === "paragraph" && empiezaMayus && /[.:;]$/.test(current.text)) ||
        ((current.type === "paragraph" || current.type === "bullet") && empiezaMayusOCifra &&
          (blankBefore[i] || prevLen < 50) &&
          !(/^\d/.test(line) && current.text.endsWith(":"))));
    if (current && !nuevoItem) {
      current.text = (current.text + " " + line).trim();
    } else {
      if (current) blocks.push(current);
      // OJO: se deja `current` apuntando a este párrafo (no se hace push
      // todavía) para que la siguiente línea, si es una continuación
      // envuelta del mismo párrafo, se una aquí en vez de crear un bloque
      // nuevo por cada línea física del OCR.
      current = { type: "paragraph", text: line };
    }
    prevLen = thisLen;
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

  f = {
    ...f,
    fecha_elaboracion: fechaDeHoy(),
    reemplaza: f.fecha_elaboracion || f.reemplaza || "",
    elaborada_por: ELABORADA_POR_FIJO,
    revisado_por: REVISADO_POR_FIJO,
  };
  const rightFieldsDef = [
    ["fecha_elaboracion", "Fecha de elaboración"],
    ["reemplaza", "Reemplaza a la norma del"],
    ["elaborada_por", "Elaborada por"],
    ["revisado_por", "Revisado por"],
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
        for (const name of normalizarResponsables(b.names, true)) children.push(makeParagraph(name));
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

// Zona del cuerpo (fracciones del alto/ancho de la página) que ocupa el
// cuadro de contenido debajo de la cabecera en estos manuales. Se usa para
// los anexos: formularios con tablas que el OCR no puede reconstruir.
const BODY_CROP = [0.08, 0.228, 0.93, 0.86]; // izquierda, arriba, derecha, abajo

// Quita del texto OCR de una página las líneas del cuadro de cabecera que no
// se pudo interpretar: se salta todo hasta la primera línea que parezca
// cuerpo (marcador de sección, ítem o línea larga), mirando como mucho las
// 14 primeras líneas.
function cuerpoSinCabecera(text) {
  const lines = text.split("\n");
  let start = 0;
  for (let i = 0; i < Math.min(lines.length, 14); i++) {
    const ln = lines[i].trim();
    if (!ln) continue;
    const esCuerpo = /^(OBJETIVOS?|RESPONSABLES?|PASOS A SEGUIR|POL[IÍ]TICAS)/i.test(ln)
      || NUM_ITEM_RE.test(ln) || BULLET_ITEM_RE.test(ln) || ln.length >= 60;
    const esCabecera = /SPI\s+SWISS|Manual de Procedimientos|P?[aá]gina\s*:|Fecha de elaboraci|Reemplaza a la norma|Elaborada por|Autorizado por|^Area|^[AÁ]rea|^Campo|^Norma/i.test(ln);
    if (esCuerpo && !esCabecera) { start = i; break; }
    start = i + 1;
  }
  return lines.slice(start).join("\n");
}

// Marcas de las páginas de la intranet que envuelven al manual (portada con
// el menú y página final con el pie del sitio): no llevan contenido y, a
// pedido del usuario, no se incluyen en el .docx.
const INTRANET_MARKERS = [
  /Intranet corporativa/i,
  /Cerrar sesi[oó]n/i,
  /Todos los derechos reservados/i,
  /Anexos telef[oó]nicos/i,
  /Procedimientos, eventos y comunicaciones/i,
];

// Página sin cabecera que además es solo "cáscara" de la intranet
// (menú/portada o pie del sitio), o que casi no tiene texto.
function esPaginaIntranet(rawText, textLimpio) {
  if (/SPI\s+SWISS/i.test(rawText)) return false;
  if (INTRANET_MARKERS.some((p) => p.test(rawText))) return true;
  return textLimpio.replace(/\s+/g, " ").trim().length < 120;
}

// Las páginas "ANEXO N" son formularios (tablas, líneas de puntos) que como
// texto quedan vacíos: se insertan como imagen bajo la cabecera.
function esPaginaAnexo(bodyText) {
  const primeras = bodyText.split("\n").map((ln) => ln.trim()).filter(Boolean).slice(0, 3);
  return primeras.some((ln) => /^ANEXO\s*\d*\s*$/i.test(ln));
}

// Inserta solo la zona del cuerpo de la página (bajo la cabecera).
async function buildBodyImage(canvas) {
  const [l, t, r, b] = BODY_CROP;
  const sx = Math.round(canvas.width * l);
  const sy = Math.round(canvas.height * t);
  const sw = Math.round(canvas.width * (r - l));
  const sh = Math.round(canvas.height * (b - t));
  const recorte = document.createElement("canvas");
  recorte.width = sw;
  recorte.height = sh;
  recorte.getContext("2d").drawImage(canvas, sx, sy, sw, sh, 0, 0, sw, sh);
  return buildFullPageImage(recorte);
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

// Renumera 'pagina' ("k/N") en las cabeceras reconocidas: k es el orden de
// la página entre las reconocidas (el OCR confunde "7/7" con "1/7" o se come
// la línea entera) y N el total que más se repite en lo leído, o el número
// de páginas reconocidas si no se leyó ninguno.
function completarNumerosDePagina(listaFields) {
  const reconocidas = listaFields.filter((f) => f !== null);
  const conteo = new Map();
  for (const f of reconocidas) {
    const m = /^\s*\d{1,2}\s*\/\s*(\d{1,2})\s*$/.exec(f.pagina || "");
    if (m) conteo.set(m[1], (conteo.get(m[1]) || 0) + 1);
  }
  let total = String(reconocidas.length);
  let mejor = 0;
  for (const [t, n] of conteo) {
    if (n > mejor) { mejor = n; total = t; }
  }
  reconocidas.forEach((f, i) => {
    f.pagina = `${i + 1}/${total}`;
  });
}

async function convertPdf(file, pageRangeSpec, onLog) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const numeroArchivo = numeroDesdeNombre(file.name);

  const range = parsePageRange(pageRangeSpec);
  const firstPage = range ? Math.max(1, range[0]) : 1;
  const lastPage = range ? Math.min(pdf.numPages, range[1]) : pdf.numPages;

  const worker = await getOcrWorker(onLog);

  const allChildren = [];
  const totalPages = lastPage - firstPage + 1;

  // Primero se hace el OCR de todas las páginas, para poder numerar ("k/N")
  // las que llegaron sin "Página : ..." legible.
  const paginas = [];
  for (let p = firstPage; p <= lastPage; p++) {
    const idx = p - firstPage;
    onLog(`OCR página ${idx + 1}/${totalPages} (página ${p} del PDF)...`);

    const canvas = await renderPageToCanvas(pdf, p);
    const { data: { text: rawText } } = await worker.recognize(canvas);
    const text = stripNoise(rawText);
    const [fields, bodyText] = parseHeader(text);
    if (fields !== null && numeroArchivo) fields.numero = numeroArchivo;
    if (fields === null && esPaginaIntranet(rawText, text)) {
      onLog(`  · Página ${p} omitida (portada/pie de la intranet, sin contenido).`);
      continue;
    }
    // Cabecera ilegible en una página de continuación: se reutiliza la de
    // la página anterior (misma norma) y el cuerpo se toma sin las líneas
    // del cuadro.
    const anterior = [...paginas].reverse().find((pg) => pg.fields);
    if (fields === null && anterior && text.trim().length >= 300) {
      onLog(`  · Página ${p}: cabecera ilegible; se reutiliza la de la página anterior.`);
      paginas.push({ p, canvas, text, fields: { ...anterior.fields, pagina: "" }, bodyText: cuerpoSinCabecera(text) });
      continue;
    }
    paginas.push({ p, canvas, text, fields, bodyText });
  }
  completarNumerosDePagina(paginas.map((pg) => pg.fields));

  let primera = true;
  for (const { p, canvas, text, fields, bodyText } of paginas) {
    if (!primera) {
      allChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }
    primera = false;

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
    if (esPaginaAnexo(bodyText)) {
      onLog(`  · Página de anexo (formulario): el cuerpo se insertó como imagen.`);
      allChildren.push(await buildBodyImage(canvas));
      continue;
    }
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
const zipRowEl = document.getElementById("zipRow");
const zipBtn = document.getElementById("zipBtn");
const zipInfoEl = document.getElementById("zipInfo");

// Documentos generados en la última conversión, para el ZIP.
let generados = [];

function actualizarZip() {
  zipRowEl.hidden = generados.length < 2;
  zipInfoEl.textContent = generados.length >= 2 ? `${generados.length} archivos` : "";
}

zipBtn.addEventListener("click", async () => {
  if (!generados.length) return;
  zipBtn.disabled = true;
  zipInfoEl.textContent = "Armando el ZIP...";
  try {
    const zip = new JSZip();
    for (const { name, blob } of generados) zip.file(name, blob);
    const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(zipBlob);
    a.download = `documentos_convertidos_${fechaDeHoy().replace(/\./g, "-")}.zip`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    zipInfoEl.textContent = `${generados.length} archivos`;
  } catch (err) {
    console.error(err);
    zipInfoEl.textContent = `No se pudo armar el ZIP: ${err.message || err}`;
  } finally {
    zipBtn.disabled = false;
  }
});

let selectedFiles = [];

function esWord(file) {
  return /\.docx$/i.test(file.name);
}

function esPdf(file) {
  return file.type === "application/pdf" || /\.pdf$/i.test(file.name);
}

function updateFileLabel() {
  if (selectedFiles.length === 0) {
    dropLabel.textContent = "Selecciona uno o varios PDF o Word (.docx), o arrástralos aquí";
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
  const files = Array.from(e.dataTransfer.files || []).filter((f) => esPdf(f) || esWord(f));
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
  generados = [];
  actualizarZip();

  const pageRangeSpec = pageRangeInput.value.trim();

  for (const file of selectedFiles) {
    const log = fileLogSection(file.name);
    try {
      log.line("Procesando...");
      let blob, outName;
      if (esWord(file)) {
        // Word -> PDF con el modelo de los manuales (ver word_a_pdf.js)
        blob = await convertDocxToPdf(file, (msg) => log.line(msg));
        outName = file.name.replace(/\.docx$/i, "") + ".pdf";
      } else {
        blob = await convertPdf(file, pageRangeSpec, (msg) => log.line(msg));
        outName = file.name.replace(/\.pdf$/i, "") + ".docx";
      }
      const url = URL.createObjectURL(blob);

      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = url;
      a.download = outName;
      a.textContent = `Descargar ${outName}`;
      li.appendChild(a);
      downloadListEl.appendChild(li);
      downloadsEl.hidden = false;
      generados.push({ name: outName, blob });
      actualizarZip();

      log.line(`Listo -> ${outName}`);
    } catch (err) {
      console.error(err);
      log.line(`Error: ${err.message || err}`, true);
    }
  }

  const finalLine = document.createElement("div");
  finalLine.className = "line";
  finalLine.textContent = "Listo. Revisa los documentos generados antes de usarlos (el OCR puede tener errores).";
  logEl.appendChild(finalLine);

  convertBtn.disabled = false;
});
