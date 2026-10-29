// ------------------------------------------------------------------ //
// Word (.docx) -> PDF con el modelo de los manuales "SPI Swissôtel Lima".
//
// Lee el .docx en el navegador (JSZip + DOMParser), reconoce el cuadro de
// cabecera (N°, Página, Área, Campo, Norma, fechas...) y el cuadro de
// contenido (OBJETIVO, RESPONSABLES, PASOS A SEGUIR, viñetas, numerados...),
// y los vuelve a dibujar en un PDF A4 con pdf-lib, siempre con la misma
// letra (Times, como el PDF original) aunque el Word venga con otra.
//
// Es la versión para navegador de word_a_pdf.py; ambos deben producir el
// mismo diseño.
// ------------------------------------------------------------------ //

// --- Diseño de página (puntos PostScript; 1 pt = 20 twips) ----------
const PDF_PAGE_W = 595.28;   // A4
const PDF_PAGE_H = 841.89;
const PDF_MARGIN = 35;        // 700 twips
const PDF_BOX_W = 495;        // 9900 twips: ancho de los cuadros
const PDF_HDR_LEFT_W = 340;   // 6800 twips
const PDF_BOX_X = (PDF_PAGE_W - PDF_BOX_W) / 2;
const PDF_PAD = 6;            // relleno interior de las celdas
const PDF_LINE_GAP = 1.25;    // interlineado (x tamaño de letra)
const PDF_PARA_GAP = 5;       // 100 twips de "after" en el Word
const PDF_INDENT_LIST = 18;   // 360 twips
const PDF_INDENT_LETTER = 25; // 500 twips

// Etiquetas del cuadro de cabecera tal como las escribe el .docx (o como
// las escribiría una persona a mano). Se comparan sin tildes ni mayúsculas.
const DOCX_HEADER_LABELS = [
  ["numero", /^n[°º*o]?\s*[:.]/i],
  ["pagina", /^p[aá]gina\s*:?/i],
  ["area", /^[aá]rea\s*:?/i],
  ["campo", /^campo\s*:?/i],
  ["norma_titulo", /^norma\s*:?/i],
  ["fecha_elaboracion", /^fecha de elaboraci[oó]n\s*:?/i],
  ["reemplaza", /^reemplaza a la norma del\s*:?/i],
  ["elaborada_por", /^elaborada? por\s*:?/i],
  ["revisado_por", /^revisado por\s*:?/i],
  ["referencia_corporativa", /^referencia corporativa(?: sl)?\s*:?/i],
  ["revisado_aprobado", /^revisado y aprobado por\s*:?/i],
  ["autorizado_por", /^autorizado por el gerente general\s*:?/i],
];

// --- 1. Lectura del .docx -------------------------------------------

// Devuelve una lista de "elementos" en el orden del documento:
//   { type: "table", rows: [[cellParagraphs...], ...] }
//   { type: "paragraph", runs: [{text, bold}], indent, text }
//   { type: "image", bytes, ext }
//   { type: "pagebreak" }
async function leerDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file("word/document.xml").async("string");
  const relsFile = zip.file("word/_rels/document.xml.rels");
  const relsXml = relsFile ? await relsFile.async("string") : "";
  const parser = new DOMParser();
  const doc = parser.parseFromString(docXml, "application/xml");
  const rels = parser.parseFromString(relsXml || "<Relationships/>", "application/xml");

  const relTargets = {};
  for (const r of rels.getElementsByTagName("Relationship")) {
    relTargets[r.getAttribute("Id")] = r.getAttribute("Target");
  }

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

  async function leerParrafo(p) {
    const runs = [];
    let indent = 0;
    let pageBreak = false;
    let image = null;
    const pPr = p.getElementsByTagNameNS(W, "pPr")[0];
    if (pPr) {
      const ind = pPr.getElementsByTagNameNS(W, "ind")[0];
      if (ind) indent = Number(ind.getAttributeNS(W, "left") || ind.getAttribute("w:left") || 0);
    }
    for (const r of p.getElementsByTagNameNS(W, "r")) {
      const rPr = r.getElementsByTagNameNS(W, "rPr")[0];
      let bold = false;
      if (rPr) {
        const b = rPr.getElementsByTagNameNS(W, "b")[0];
        if (b) {
          const val = b.getAttributeNS(W, "val") || b.getAttribute("w:val");
          bold = val === null || val === "" || val === "1" || val === "true" || val === "on";
        }
      }
      for (const child of r.childNodes) {
        if (child.namespaceURI !== W) continue;
        if (child.localName === "t") runs.push({ text: child.textContent, bold });
        else if (child.localName === "tab") runs.push({ text: "\t", bold });
        else if (child.localName === "br") {
          const type = child.getAttributeNS(W, "type") || child.getAttribute("w:type");
          if (type === "page") pageBreak = true;
          else runs.push({ text: "\n", bold });
        } else if (child.localName === "drawing") {
          const blips = child.getElementsByTagNameNS(A, "blip");
          if (blips.length) {
            const rid = blips[0].getAttributeNS(R, "embed") || blips[0].getAttribute("r:embed");
            const target = relTargets[rid];
            if (target) {
              const f = zip.file("word/" + target.replace(/^\/?word\//, ""));
              if (f) {
                const bytes = await f.async("uint8array");
                image = { bytes, ext: target.split(".").pop().toLowerCase() };
              }
            }
          }
        }
      }
    }
    const text = runs.map((x) => x.text).join("").replace(/\t/g, " ").trim();
    return { runs, indent, text, pageBreak, image };
  }

  const elementos = [];
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  for (const node of body.childNodes) {
    if (node.namespaceURI !== W) continue;
    if (node.localName === "p") {
      const par = await leerParrafo(node);
      if (par.image) elementos.push({ type: "image", ...par.image });
      if (par.text) elementos.push({ type: "paragraph", runs: par.runs, indent: par.indent, text: par.text });
      if (par.pageBreak) elementos.push({ type: "pagebreak" });
    } else if (node.localName === "tbl") {
      const rows = [];
      for (const tr of node.getElementsByTagNameNS(W, "tr")) {
        if (tr.parentNode !== node) continue;
        const cells = [];
        for (const tc of tr.getElementsByTagNameNS(W, "tc")) {
          if (tc.parentNode !== tr) continue;
          const pars = [];
          for (const p of tc.getElementsByTagNameNS(W, "p")) {
            const par = await leerParrafo(p);
            if (par.image) pars.push({ type: "image", ...par.image });
            if (par.text) pars.push({ type: "paragraph", runs: par.runs, indent: par.indent, text: par.text });
          }
          cells.push(pars);
        }
        rows.push(cells);
      }
      elementos.push({ type: "table", rows });
    }
  }
  return elementos;
}

// --- 2. Interpretación: cabecera y cuerpo -----------------------------

function limpiarValor(v) {
  return (v || "").replace(/^\s*[:.\-–—]+\s*/, "").replace(/\s+/g, " ").trim();
}

// Reconoce "Etiqueta : valor" en una lista de líneas. Devuelve null si no
// aparece N° ni Página (entonces no es un cuadro de cabecera).
function parseHeaderLineas(lineas) {
  const f = {};
  let encontrado = 0;
  for (const raw of lineas) {
    const ln = raw.trim();
    if (!ln) continue;
    if (/^SPI\s+SWISS/i.test(ln) || /^Manual de Procedimientos/i.test(ln)) { encontrado++; continue; }
    let matched = false;
    for (const [key, re] of DOCX_HEADER_LABELS) {
      const m = re.exec(ln);
      if (m) {
        f[key] = limpiarValor(ln.slice(m[0].length));
        encontrado++;
        matched = true;
        break;
      }
    }
    if (!matched && f.norma_titulo !== undefined && !f.fecha_elaboracion) {
      // continuación del título de la norma en otra línea
      f.norma_titulo = (f.norma_titulo + " " + ln).trim();
    }
  }
  if (f.numero === undefined && f.pagina === undefined) return null;
  return encontrado >= 3 ? f : null;
}

// Convierte los párrafos de un cuadro de contenido en bloques como los de
// parseBody(): objetivo / responsables / subheading / numbered / letter /
// bullet / paragraph.
function bloquesDesdeParrafos(pars) {
  const blocks = [];
  let i = 0;
  while (i < pars.length) {
    const p = pars[i];
    if (p.type === "image") { blocks.push({ type: "image", bytes: p.bytes, ext: p.ext }); i++; continue; }
    const text = p.text;
    const bold = p.runs.length > 0 && p.runs.every((r) => r.bold || !r.text.trim());
    let m;
    if (/^OBJETIVOS?\s*:?\s*$/i.test(text)) {
      const partes = [];
      i++;
      while (i < pars.length && pars[i].type === "paragraph" && !esTituloParrafo(pars[i])) { partes.push(pars[i].text); i++; }
      blocks.push({ type: "objetivo", text: partes.join(" ") });
      continue;
    }
    if (/^RESPONSABLES?\s*:?\s*$/i.test(text)) {
      const names = [];
      i++;
      while (i < pars.length && pars[i].type === "paragraph" && !esTituloParrafo(pars[i])) { names.push(pars[i].text); i++; }
      blocks.push({ type: "responsables", names });
      continue;
    }
    if ((m = /^(\d{1,2})[.)]\s+(.*)$/.exec(text)) && !bold) {
      blocks.push({ type: "numbered", n: m[1], text: m[2] });
    } else if ((m = /^([a-h])\)\s+(.*)$/.exec(text)) && !bold) {
      blocks.push({ type: "letter", letter: m[1], text: m[2] });
    } else if ((m = /^[•·\-*]\s+(.*)$/.exec(text))) {
      blocks.push({ type: "bullet", text: m[1] });
    } else if (bold || esTituloParrafo(p)) {
      blocks.push({ type: "subheading", text });
    } else {
      blocks.push({ type: "paragraph", text });
    }
    i++;
  }
  return blocks;
}

function esTituloParrafo(p) {
  if (p.type !== "paragraph") return false;
  const bold = p.runs.length > 0 && p.runs.every((r) => r.bold || !r.text.trim());
  const t = p.text;
  return bold || (t.length <= 45 && t.split(/\s+/).length <= 4 && /:$/.test(t) && /^[A-ZÁÉÍÓÚÑ]/.test(t))
    || (t.length <= 45 && t === t.toUpperCase() && /[A-ZÁÉÍÓÚÑ]/.test(t) && !/[.,;]$/.test(t));
}

// Agrupa los elementos del .docx en "páginas" del modelo:
//   { fields, blocks, images }  (cabecera + contenido)
//   { fullImage }               (página que venía como imagen completa)
function paginasDesdeDocx(elementos) {
  const paginas = [];
  let actual = null;
  const cerrar = () => { if (actual) { paginas.push(actual); actual = null; } };

  for (const el of elementos) {
    if (el.type === "pagebreak") { cerrar(); continue; }
    if (el.type === "table") {
      const celdas = el.rows.flat();
      const lineas = celdas.flat().filter((p) => p.type === "paragraph").map((p) => p.text);
      const fields = el.rows.length >= 2 ? parseHeaderLineas(lineas) : null;
      if (fields) {
        cerrar();
        actual = { fields, blocks: [] };
        continue;
      }
      // Cuadro de contenido (u otra tabla): sus párrafos van al cuerpo
      const pars = celdas.flat();
      if (!actual) actual = { fields: null, blocks: [] };
      actual.blocks.push(...bloquesDesdeParrafos(pars));
      continue;
    }
    if (el.type === "image") {
      if (actual && actual.fields && actual.blocks.length === 0) {
        actual.blocks.push({ type: "image", bytes: el.bytes, ext: el.ext });
      } else {
        cerrar();
        paginas.push({ fullImage: { bytes: el.bytes, ext: el.ext } });
      }
      continue;
    }
    if (el.type === "paragraph") {
      if (!actual) actual = { fields: null, blocks: [] };
      actual.blocks.push(...bloquesDesdeParrafos([el]));
    }
  }
  cerrar();

  // Word "plano" (sin tablas): intentar sacar la cabecera de las primeras
  // líneas de texto y el cuerpo del resto.
  for (const pg of paginas) {
    if (pg.fullImage || pg.fields) continue;
    const textos = pg.blocks.filter((b) => b.text).map((b) => b.text);
    const fields = parseHeaderLineas(textos.slice(0, 16));
    if (fields) {
      pg.fields = fields;
      let k = 0, vistos = 0;
      for (; k < pg.blocks.length && vistos < 16; k++) {
        const t = pg.blocks[k].text || "";
        if (/^SPI\s+SWISS/i.test(t) || /^Manual de Procedimientos/i.test(t) || DOCX_HEADER_LABELS.some(([, re]) => re.test(t))) vistos = k + 1;
      }
      pg.blocks = pg.blocks.slice(vistos);
    }
  }
  return paginas;
}

// --- 3. Dibujo del PDF (pdf-lib) ---------------------------------------

// Texto seguro para las fuentes estándar (WinAnsi): lo que no exista se
// sustituye para que pdf-lib no falle.
function textoWinAnsi(s) {
  return s
    .replace(/[‘’‚]/g, "'")
    .replace(/[“”„]/g, '"')
    .replace(/…/g, "...")
    .replace(/[–—]/g, "-")
    .replace(/ /g, " ")
    .replace(/[^\x09\x0A\x0D\x20-\x7E\xA0-\xFF€•]/g, "?");
}

function envolverTexto(texto, font, size, maxW) {
  const palabras = textoWinAnsi(texto).split(/\s+/).filter(Boolean);
  const lineas = [];
  let actual = "";
  for (const w of palabras) {
    const prueba = actual ? actual + " " + w : w;
    if (font.widthOfTextAtSize(prueba, size) <= maxW || !actual) {
      actual = prueba;
    } else {
      lineas.push(actual);
      actual = w;
    }
  }
  if (actual) lineas.push(actual);
  return lineas.length ? lineas : [""];
}

class PdfEscritor {
  constructor(pdfDoc, fonts) {
    this.pdf = pdfDoc;
    this.fonts = fonts; // { regular, bold }
    this.page = null;
    this.y = 0;
  }

  nuevaPagina() {
    this.page = this.pdf.addPage([PDF_PAGE_W, PDF_PAGE_H]);
    this.y = PDF_PAGE_H - PDF_MARGIN;
    return this.page;
  }

  // Dibuja el cuadro de cabecera en la parte superior de la página actual.
  dibujarCabecera(f) {
    const { regular, bold } = this.fonts;
    const x0 = PDF_BOX_X;
    const xMid = x0 + PDF_HDR_LEFT_W;
    const top = this.y;

    const izqFila1 = [["SPI SWISSÔTEL LIMA", bold, 11], ["Manual de Procedimientos", regular, 10]];
    const derFila1 = [[`N°  : ${f.numero || ""}`, bold, 10], [`Página: ${f.pagina || ""}`, regular, 10]];
    const izqFila2 = [
      [`Area  : ${f.area || ""}`, regular, 9],
      [`Campo : ${f.campo || ""}`, regular, 9],
      [`Norma : ${f.norma_titulo || ""}`, regular, 9],
    ];
    const derFila2 = [];
    for (const [key, label] of [
      ["fecha_elaboracion", "Fecha de elaboración"],
      ["reemplaza", "Reemplaza a la norma del"],
      ["elaborada_por", "Elaborada por"],
      ["revisado_por", "Revisado por"],
      ["referencia_corporativa", "Referencia Corporativa SL"],
      ["revisado_aprobado", "Revisado y aprobado por"],
    ]) {
      if (f[key]) derFila2.push([`${label} : ${f[key]}`, regular, 9]);
    }
    derFila2.push([`Autorizado por el Gerente General : ${f.autorizado_por || ""}`.trim(), regular, 9]);

    const altoCelda = (items, maxW) => {
      let h = PDF_PAD;
      for (const [t, font, size] of items) {
        h += envolverTexto(t, font, size, maxW).length * size * PDF_LINE_GAP + PDF_PARA_GAP;
      }
      return h + PDF_PAD - PDF_PARA_GAP;
    };
    const wIzq = PDF_HDR_LEFT_W - 2 * PDF_PAD;
    const wDer = PDF_BOX_W - PDF_HDR_LEFT_W - 2 * PDF_PAD;
    const h1 = Math.max(altoCelda(izqFila1, wIzq), altoCelda(derFila1, wDer));
    const h2 = Math.max(altoCelda(izqFila2, wIzq), altoCelda(derFila2, wDer));

    const dibujarCelda = (items, x, yTop, maxW) => {
      let y = yTop - PDF_PAD;
      for (const [t, font, size] of items) {
        for (const ln of envolverTexto(t, font, size, maxW)) {
          y -= size * PDF_LINE_GAP;
          this.page.drawText(ln, { x: x + PDF_PAD, y: y + size * 0.25, size, font });
        }
        y -= PDF_PARA_GAP;
      }
    };
    dibujarCelda(izqFila1, x0, top, wIzq);
    dibujarCelda(derFila1, xMid, top, wDer);
    dibujarCelda(izqFila2, x0, top - h1, wIzq);
    dibujarCelda(derFila2, xMid, top - h1, wDer);

    const grosor = 0.75;
    const rect = (x, y, w, h) => this.page.drawRectangle({ x, y, width: w, height: h, borderWidth: grosor, borderColor: PDFLib.rgb(0, 0, 0) });
    rect(x0, top - h1, PDF_HDR_LEFT_W, h1);
    rect(xMid, top - h1, PDF_BOX_W - PDF_HDR_LEFT_W, h1);
    rect(x0, top - h1 - h2, PDF_HDR_LEFT_W, h2);
    rect(xMid, top - h1 - h2, PDF_BOX_W - PDF_HDR_LEFT_W, h2);

    this.y = top - h1 - h2 - 12; // espaciador, como el párrafo vacío del Word
  }

  // Dibuja los bloques dentro de un cuadro con borde; si no caben, cierra
  // el cuadro, abre página nueva con la misma cabecera y continúa.
  async dibujarContenido(blocks, fields) {
    const { regular, bold } = this.fonts;
    const x0 = PDF_BOX_X;
    const maxW = PDF_BOX_W - 2 * PDF_PAD;
    let boxTop = this.y;
    let y = boxTop - PDF_PAD;
    const bottomLimit = PDF_MARGIN + PDF_PAD;

    // Un cuerpo que es solo imagen (anexo/formulario) ya trae su propio
    // marco en la imagen: no se le dibuja otro cuadro encima.
    const soloImagenes = blocks.length > 0 && blocks.every((b) => b.type === "image");
    const cerrarCuadro = () => {
      if (soloImagenes) return;
      const h = boxTop - (y - PDF_PAD);
      this.page.drawRectangle({ x: x0, y: y - PDF_PAD, width: PDF_BOX_W, height: h, borderWidth: 0.75, borderColor: PDFLib.rgb(0, 0, 0) });
    };
    const saltar = () => {
      cerrarCuadro();
      this.nuevaPagina();
      if (fields) this.dibujarCabecera(fields);
      boxTop = this.y;
      y = boxTop - PDF_PAD;
    };

    const lineasDe = [];
    for (const b of blocks) {
      if (b.type === "objetivo") {
        lineasDe.push({ font: bold, size: 10, text: "OBJETIVO:", indent: 0 });
        lineasDe.push({ font: regular, size: 10, text: b.text, indent: 0 });
      } else if (b.type === "responsables") {
        lineasDe.push({ font: bold, size: 10, text: "RESPONSABLES:", indent: 0 });
        for (const n of b.names) lineasDe.push({ font: regular, size: 10, text: n, indent: 0 });
      } else if (b.type === "subheading") {
        lineasDe.push({ font: bold, size: 10, text: b.text, indent: 0 });
      } else if (b.type === "numbered") {
        lineasDe.push({ font: regular, size: 10, text: `${b.n}.  ${b.text}`, indent: PDF_INDENT_LIST });
      } else if (b.type === "letter") {
        lineasDe.push({ font: regular, size: 10, text: `${b.letter})  ${b.text}`, indent: PDF_INDENT_LETTER });
      } else if (b.type === "bullet") {
        lineasDe.push({ font: regular, size: 10, text: `•  ${b.text}`, indent: PDF_INDENT_LIST });
      } else if (b.type === "image") {
        lineasDe.push({ image: b });
      } else {
        lineasDe.push({ font: regular, size: 10, text: b.text, indent: 0 });
      }
    }

    for (const item of lineasDe) {
      if (item.image) {
        const img = await this.incrustarImagen(item.image);
        if (!img) continue;
        const w = maxW;
        const h = (img.height / img.width) * w;
        if (y - h < bottomLimit) saltar();
        this.page.drawImage(img, { x: x0 + PDF_PAD, y: y - h, width: w, height: h });
        y -= h + PDF_PARA_GAP;
        continue;
      }
      const lineas = envolverTexto(item.text, item.font, item.size, maxW - item.indent);
      for (const ln of lineas) {
        if (y - item.size * PDF_LINE_GAP < bottomLimit) saltar();
        y -= item.size * PDF_LINE_GAP;
        this.page.drawText(ln, { x: x0 + PDF_PAD + item.indent, y: y + item.size * 0.25, size: item.size, font: item.font });
      }
      y -= PDF_PARA_GAP;
    }
    cerrarCuadro();
    this.y = y - PDF_PAD - 12;
  }

  async incrustarImagen(img) {
    try {
      if (img.ext === "png") return await this.pdf.embedPng(img.bytes);
      if (img.ext === "jpg" || img.ext === "jpeg") return await this.pdf.embedJpg(img.bytes);
      // Otros formatos: se pasan por un <canvas> para obtener PNG.
      const blob = new Blob([img.bytes]);
      const bmp = await createImageBitmap(blob);
      const c = document.createElement("canvas");
      c.width = bmp.width; c.height = bmp.height;
      c.getContext("2d").drawImage(bmp, 0, 0);
      const dataUrl = c.toDataURL("image/png");
      const bytes = new Uint8Array(await (await fetch(dataUrl)).arrayBuffer());
      return await this.pdf.embedPng(bytes);
    } catch (e) {
      console.warn("No se pudo incrustar una imagen del Word:", e);
      return null;
    }
  }

  async dibujarImagenCompleta(img) {
    const emb = await this.incrustarImagen(img);
    if (!emb) return;
    const w = PDF_PAGE_W - 2 * PDF_MARGIN;
    const h = Math.min((emb.height / emb.width) * w, PDF_PAGE_H - 2 * PDF_MARGIN);
    this.page.drawImage(emb, { x: PDF_MARGIN, y: this.y - h, width: w, height: h });
    this.y -= h;
  }
}

// --- 4. Orquestación ----------------------------------------------------

async function convertDocxToPdf(file, onLog) {
  const arrayBuffer = await file.arrayBuffer();
  onLog("Leyendo el Word...");
  const elementos = await leerDocx(arrayBuffer);
  const paginas = paginasDesdeDocx(elementos);
  if (!paginas.length) throw new Error("El Word no tiene contenido reconocible.");

  const pdfDoc = await PDFLib.PDFDocument.create();
  const fonts = {
    regular: await pdfDoc.embedFont(PDFLib.StandardFonts.TimesRoman),
    bold: await pdfDoc.embedFont(PDFLib.StandardFonts.TimesRomanBold),
  };
  const esc = new PdfEscritor(pdfDoc, fonts);

  let n = 0;
  for (const pg of paginas) {
    n++;
    esc.nuevaPagina();
    if (pg.fullImage) {
      onLog(`  · Página ${n}: imagen completa (se copia tal cual).`);
      await esc.dibujarImagenCompleta(pg.fullImage);
      continue;
    }
    if (pg.fields) {
      onLog(`  ✓ Página ${n}: cabecera (N° ${pg.fields.numero || "?"}, página ${pg.fields.pagina || "?"}).`);
      esc.dibujarCabecera(pg.fields);
    } else {
      onLog(`  ⚠ Página ${n}: sin cuadro de cabecera reconocible; se escribe el texto con el mismo formato.`);
    }
    if (pg.blocks.length) await esc.dibujarContenido(pg.blocks, pg.fields);
  }

  pdfDoc.setTitle(file.name.replace(/\.docx$/i, ""));
  const bytes = await pdfDoc.save();
  return new Blob([bytes], { type: "application/pdf" });
}
