# Conversión PDF → Word (manuales SPI Swissôtel Lima)

Convierte los manuales de procedimientos "SPI Swissôtel Lima" (PDF escaneado,
sin capa de texto) a un Word (.docx) editable que replica el diseño original.

Hay dos formas de usarlo:

## 1. App web (GitHub Pages)

👉 **[leidercostoshotel-code.github.io/conversionesleider](https://leidercostoshotel-code.github.io/conversionesleider/)**

Todo corre en el navegador (pdf.js + Tesseract.js + docx.js): no se sube
ningún archivo a un servidor. Solo abre el enlace, selecciona el/los PDF y
descarga el/los `.docx` generados.

Archivos: `index.html`, `app.js`, `style.css`.

## 2. Script de línea de comandos (`pdf_a_word.py`)

Para procesar muchos archivos por lotes o integrarlo en otro flujo, usa el
script de Python. Instalación y uso completos en la cabecera de
[`pdf_a_word.py`](pdf_a_word.py):

```bash
apt install tesseract-ocr tesseract-ocr-spa poppler-utils   # Linux
pip install python-docx pytesseract pdf2image pillow

python pdf_a_word.py archivo.pdf
python pdf_a_word.py carpeta_con_pdfs/
```

## Cómo funciona (y sus límites)

1. Cada página del PDF se rasteriza a imagen.
2. Se le aplica OCR en español.
3. Si el texto reconocido coincide con la cabecera fija de estos manuales
   (SPI SWISSÔTEL LIMA / N° / Área / Campo / Norma / fechas), la página se
   reconstruye como texto real y editable.
4. Si una página no tiene ese patrón (formularios, organigramas, portadas),
   se inserta tal cual como imagen, para no perder información aunque no
   quede editable.
5. El OCR nunca es 100% exacto. Revisa siempre el `.docx` resultante antes
   de usarlo.

## Activar GitHub Pages (si aún no está activo)

Si el enlace de arriba da 404, ve a **Settings → Pages** de este repositorio
y en *Build and deployment* elige *Source: Deploy from a branch*, rama
`main`, carpeta `/ (root)`. GitHub publica el sitio en un par de minutos.
