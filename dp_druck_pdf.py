"""
Streamlit App: OCR-gestützte PDF-Annotation nach Namenssuche
===========================================================
• Lädt eine mehrseitige PDF-Datei und eine Excel-Tabelle hoch
• Nutzer markiert per Maus (Drawable-Canvas) einen Suchbereich auf der ersten PDF-Seite
• OCR wird **nur** in diesem Bereich aller Seiten ausgeführt
• Wird ein Name gefunden, wird der zugehörige Wert aus der Excel-Tabelle auf der Fund-Seite platziert
• Annotierte PDF steht zum Download bereit

Benötigte Pakete
----------------
```
streamlit
streamlit-drawable-canvas
pandas
pdf2image
pillow
pytesseract
pymupdf  # fitz
openpyxl
```

> **Hinweis:** [Tesseract](https://github.com/tesseract-ocr) muss lokal installiert sein und erreichbar sein.
"""

from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Tuple

import streamlit as st
from streamlit_drawable_canvas import st_canvas
import pandas as pd
import pytesseract
import fitz  # PyMuPDF
from pdf2image import convert_from_bytes
from PIL import Image

# -----------------------------------------------------------------------------
# Seiteneinstellungen
# -----------------------------------------------------------------------------
st.set_page_config(page_title="PDF-Namenssuche & Annotation", layout="centered")

st.title("🔎 PDF-Namenssuche mit Excel-Referenz und Annotation")

with st.expander("Kurzanleitung", expanded=False):
    st.markdown(
        """
        1. **PDF hochladen** (mehrseitig, gescannt oder Bild-PDF).
        2. **Excel hochladen** mit Spalten *Name* und *Wert* (z.B. Abteilung).
        3. **Suchbereich festlegen**: Erste PDF-Seite wird angezeigt, Rah­men ziehen, wo sich der Name befindet.
        4. **Spalten & Textposition wählen**.
        5. **Start** – du erhältst eine annotierte PDF.
        """
    )

# -----------------------------------------------------------------------------
# Datei-Uploads
# -----------------------------------------------------------------------------
pdf_file = st.file_uploader("PDF hochladen", type=["pdf"], key="pdf")
excel_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"], key="excel")

if pdf_file and excel_file:
    # Excel einlesen
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Excel konnte nicht eingelesen werden: {e}")
        st.stop()

    if df.empty:
        st.warning("Die Excel-Datei enthält keine Daten.")
        st.stop()

    # Erster Seiten-Snapshot für Canvas
    pdf_bytes = pdf_file.read()
    try:
        first_page_img: Image.Image = convert_from_bytes(pdf_bytes, first_page=1, last_page=1)[0]
    except Exception as e:
        st.error(f"PDF-Konvertierung fehlgeschlagen: {e}")
        st.stop()

    st.subheader("1️⃣ Suchbereich markieren")
    CANVAS_WIDTH = 600  # angezeigte Pixelbreite (anpassbar)
    ratio = CANVAS_WIDTH / first_page_img.width
    canvas_result = st_canvas(
        fill_color="",  # keine Füllung, nur Kontur
        stroke_width=3,
        stroke_color="#FF0000",
        background_image=first_page_img.resize((CANVAS_WIDTH, int(first_page_img.height * ratio))),
        update_streamlit=True,
        height=int(first_page_img.height * ratio),
        width=CANVAS_WIDTH,
        drawing_mode="rect",
        key="canvas",
    )

    roi: Tuple[int, int, int, int] | None = None  # (left, top, right, bottom) in Original-Pixeln
    if canvas_result.json_data and canvas_result.json_data.get("objects"):
        # letzter gezeichneter Rahmen
        rect = canvas_result.json_data["objects"][-1]
        left_disp, top_disp = rect["left"], rect["top"]
        width_disp, height_disp = rect["width"], rect["height"]
        # Skalierung zurück auf Originalauflösung
        left, top = int(left_disp / ratio), int(top_disp / ratio)
        right = int((left_disp + width_disp) / ratio)
        bottom = int((top_disp + height_disp) / ratio)
        roi = (left, top, right, bottom)
        st.info(f"ROI gesetzt: x={left}:{right}, y={top}:{bottom} (Pixel im Original)")
    else:
        st.warning("Bitte Rechteck ziehen, um den Suchbereich festzulegen.")
        st.stop()

    # -------------------------------------------------------------------------
    # Spalten & Textoptionen
    # -------------------------------------------------------------------------
    st.subheader("2️⃣ Spalten & Textposition wählen")
    col1, col2 = st.columns(2)
    with col1:
        name_col = st.selectbox("Spalte mit Namen", df.columns)
    with col2:
        value_col = st.selectbox("Spalte mit einzutragender Information", df.columns, index=min(1, len(df.columns)-1))

    st.markdown("**Position des einzutragenden Textes (Pt, 1 Pt ≈ 1⁄72 Zoll):**")
    colx, coly, colf = st.columns(3)
    with colx:
        x_position = st.number_input("X-Offset", 0, 600, value=50)
    with coly:
        y_position = st.number_input("Y-Offset", 0, 800, value=50)
    with colf:
        font_size = st.number_input("Schriftgröße", 6, 48, value=12)

    case_sensitive = st.checkbox("Groß-/Kleinschreibung beachten", value=False)

    # Name-zu-Wert-Mapping
    name_map = {
        (str(r[name_col]) if case_sensitive else str(r[name_col]).lower()): r[value_col]
        for _, r in df.iterrows() if pd.notna(r[name_col])
    }

    if st.button("🚀 Starten", disabled=roi is None):
        with st.spinner("Verarbeite PDF… bitte warten"):
            # PDF erneut als Dokument laden
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")

            # OCR-Schleife
            for page_idx in range(len(doc)):
                page = doc[page_idx]

                # Seite als Bild (PIL) erzeugen
                pix = page.get_pixmap(dpi=300)
                page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # ROI zuschneiden
                crop_img = page_img.crop(roi)

                # OCR
                ocr_text = pytesseract.image_to_string(crop_img, lang="deu")
                search_space = ocr_text if case_sensitive else ocr_text.lower()

                # Suche nach Namen
                for search_name, value in name_map.items():
                    pattern = rf"\b{re.escape(search_name)}\b"
                    if re.search(pattern, search_space):
                        insertion_text = str(value)
                        page.insert_text(
                            (x_position, y_position),
                            insertion_text,
                            fontsize=font_size,
                            fontname="helv",
                            fill=(0, 0, 0),
                        )
                        break  # Optional: nur erster Treffer pro Seite

            # Ausgabe-PDF
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            doc.close()
            output_buffer.seek(0)

        st.success("Fertig! Die PDF ist annotiert.")
        st.download_button(
            label="📥 Annotierte PDF herunterladen",
            data=output_buffer,
            file_name="annotiert.pdf",
            mime="application/pdf",
        )

# -----------------------------------------------------------------------------
# Footer
# -----------------------------------------------------------------------------
with st.expander("ℹ️ Info & Troubleshooting"):
    st.markdown(
        """
        * Dieses Tool verwendet [PyMuPDF](https://pymupdf.readthedocs.io/) zum
          Schreiben in die PDF und [pytesseract](https://pypi.org/project/pytesseract/)
          für die Texterkennung im gewählten Suchbereich.
        * Bei falsch skalierten Koordinaten stelle sicher, dass dein Monitor-Zoom
          auf 100 % steht oder passe `CANVAS_WIDTH` an.
        * Für mehrsprachige Dokumente wähle das passende Tesseract-Language-Pack.
        """
    )
