"""
Streamlit App: OCR‑gestützte PDF‑Annotation nach Namenssuche
===========================================================
Deploy‑ready for **streamlit.io (Streamlit Cloud)**
-------------------------------------------------
• Lädt eine mehrseitige PDF‑Datei und eine Excel‑Tabelle hoch
• Nutzer markiert per Maus (Drawable‑Canvas) einen Suchbereich auf der ersten PDF‑Seite
• OCR wird **nur** in dieser Zone aller Seiten ausgeführt (Tesseract)
• Wird ein Name gefunden, wird der zugehörige Wert aus der Excel‑Tabelle auf der Fund‑Seite platziert
• Annotierte PDF steht zum Download bereit

### Python requirements (requirements.txt)
```
streamlit
streamlit-drawable-canvas
pymupdf  # fitz
pandas
pillow
pytesseract
openpyxl
```

### System requirements (packages.txt for Streamlit Cloud)
```
tesseract-ocr
# deutsches Sprachpaket (optional, sonst defaults zu eng)
tesseract-ocr-deu
```
> **Keine Abhängigkeit zu Poppler**: Wir rendern Seiten ausschließlich mit PyMuPDF und benötigen daher **kein** `pdf2image` / Poppler‑Utils.
"""

from __future__ import annotations

"""Hotfix für Streamlit ≥1.32
-----------------------------
`streamlit_drawable_canvas` verwendet intern `st_image.image_to_url`,
diese Funktion wurde in neueren Streamlit‑Versionen entfernt.  Das
nachfolgende Monkey‑Patch stellt sie wieder bereit, damit die Canvas‑
Komponente auf **Streamlit Cloud** funktioniert, ohne ältere Versionen
pinnen zu müssen.
"""

import base64
import io as _io
from PIL import Image as _PIL_Image
import streamlit.elements.image as _st_image_element

if not hasattr(_st_image_element, "image_to_url"):
    def _image_to_url(img, width=None, clamp=False, channels="RGB", output_format="auto"):
        """Ersatz für die entfernte Streamlit‑Funktion.  Wandelt ein PIL‑Bild
        oder eine NumPy‑Array in eine data‑URL um, ausreichend für
        `streamlit_drawable_canvas`.
        """
        if isinstance(img, _PIL_Image.Image):
            buf = _io.BytesIO()
            img.save(buf, format="PNG")
            data = buf.getvalue()
        else:
            # Fallback: Versuche ndarray → PIL
            try:
                import numpy as np
                if isinstance(img, np.ndarray):
                    pil_img = _PIL_Image.fromarray(img)
                    buf = _io.BytesIO()
                    pil_img.save(buf, format="PNG")
                    data = buf.getvalue()
                else:
                    raise TypeError("Unsupported image type for fallback image_to_url")
            except Exception as exc:
                raise TypeError("Unsupported image type for fallback image_to_url") from exc
        b64 = base64.b64encode(data).decode()
        return f"data:image/png;base64,{b64}"

    _st_image_element.image_to_url = _image_to_url

import io
import re
from typing import Tuple

import streamlit as st
from streamlit_drawable_canvas import st_canvas
import pandas as pd
import pytesseract
import fitz  # PyMuPDF
from PIL import Image

# -----------------------------------------------------------------------------
# Seiteneinstellungen
# -----------------------------------------------------------------------------
st.set_page_config(page_title="PDF‑Namenssuche & Annotation", layout="centered")

st.title("🔎 PDF‑Namenssuche mit Excel‑Referenz und Annotation")

with st.expander("Kurzanleitung", expanded=False):
    st.markdown(
        """
        1. **PDF hochladen** (mehrseitig, gescannt oder Bild‑PDF).
        2. **Excel hochladen** mit Spalten *Name* und *Wert* (z. B. Abteilung).
        3. **Suchbereich festlegen**: Erste PDF‑Seite wird angezeigt, Rahmen ziehen, wo sich der Name befindet.
        4. **Spalten & Textposition wählen**.
        5. **Start** – du erhältst eine annotierte PDF.
        """
    )

# -----------------------------------------------------------------------------
# Datei‑Uploads
# -----------------------------------------------------------------------------
pdf_file = st.file_uploader("PDF hochladen", type=["pdf"], key="pdf")
excel_file = st.file_uploader("Excel‑Datei hochladen", type=["xlsx", "xls"], key="excel")

if pdf_file and excel_file:
    # Excel einlesen
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Excel konnte nicht eingelesen werden: {e}")
        st.stop()

    if df.empty:
        st.warning("Die Excel‑Datei enthält keine Daten.")
        st.stop()

    pdf_bytes = pdf_file.read()

    # PyMuPDF‑Dokument einmalig öffnen (auch für erste Vorschau)
    try:
        doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        st.error(f"PDF konnte nicht gelesen werden: {e}")
        st.stop()

    # Erste Seite als Bild für Canvas (150 dpi reicht fürs Zeichnen)
    first_pix = doc_preview[0].get_pixmap(dpi=150)
    first_page_img = Image.frombytes("RGB", [first_pix.width, first_pix.height], first_pix.samples)

    st.subheader("1️⃣ Suchbereich markieren")
    CANVAS_WIDTH = 600  # angezeigte Pixelbreite
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

    roi: Tuple[int, int, int, int] | None = None  # (left, top, right, bottom) in Original‑Pixeln
    if canvas_result.json_data and canvas_result.json_data.get("objects"):
        rect = canvas_result.json_data["objects"][-1]
        left_disp, top_disp = rect["left"], rect["top"]
        width_disp, height_disp = rect["width"], rect["height"]
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
        value_col = st.selectbox(
            "Spalte mit einzutragender Information",
            df.columns,
            index=min(1, len(df.columns) - 1),
        )

    st.markdown("**Position des einzutragenden Textes (Pt, 1 Pt ≈ 1⁄72 Zoll):**")
    colx, coly, colf = st.columns(3)
    with colx:
        x_position = st.number_input("X‑Offset", 0, 600, value=50)
    with coly:
        y_position = st.number_input("Y‑Offset", 0, 800, value=50)
    with colf:
        font_size = st.number_input("Schriftgröße", 6, 48, value=12)

    case_sensitive = st.checkbox("Groß‑/Kleinschreibung beachten", value=False)

    # Name‑zu‑Wert‑Mapping
    name_map = {
        (str(r[name_col]) if case_sensitive else str(r[name_col]).lower()): r[value_col]
        for _, r in df.iterrows()
        if pd.notna(r[name_col])
    }

    if st.button("🚀 Starten", disabled=roi is None):
        with st.spinner("Verarbeite PDF… bitte warten"):
            # PDF als bearbeitbares Dokument erneut öffnen
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")

            for page_idx in range(len(doc)):
                page = doc[page_idx]

                # Seite als Bild (300 dpi) für OCR
                pix = page.get_pixmap(dpi=300)
                page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # ROI zuschneiden
                crop_img = page_img.crop(roi)

                ocr_text = pytesseract.image_to_string(crop_img, lang="deu")
                search_space = ocr_text if case_sensitive else ocr_text.lower()

                # Suche nach Namen
                for search_name, value in name_map.items():
                    pattern = rf"\b{re.escape(search_name)}\b"
                    if re.search(pattern, search_space):
                        page.insert_text(
                            (x_position, y_position),
                            str(value),
                            fontsize=font_size,
                            fontname="helv",
                            fill=(0, 0, 0),
                        )
                        break  # optional: nur erster Treffer pro Seite

            # Annotierte PDF speichern
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
        * Keine Poppler‑Abhängigkeit: Das Rendering übernimmt **PyMuPDF**.
        * Lege eine `packages.txt` mit `tesseract-ocr` (und optional `tesseract-ocr-deu`) an, damit OCR auf Streamlit Cloud funktioniert.
        * Passe `CANVAS_WIDTH` an, falls dein PDF extrem breit ist oder du eine hochauflösende Vorschau brauchst.
        """
    )
