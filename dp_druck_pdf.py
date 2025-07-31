from __future__ import annotations

"""
Streamlit Utility – Interaktiver PDF-ROI-Finder & OCR
====================================================
**Ziel**
--------
1. **ROI visual bestimmen**: Nach dem PDF-Upload wird die **erste Seite** mit 300 DPI gerendert und angezeigt. Du wählst per Zahl-Inputs oder Schieberegler (`x1, y1, x2, y2`) den Bereich, in dem sich die Namen befinden.
2. Vorschau des **ausgeschnittenen Bereichs** & Sofort-OCR auf dieser ersten Seite, damit man sieht, ob Text erkannt wird.
3. **Seite-1-Overlay**: Ein zweites Bild zeigt die komplette Seite mit *rotem Rechteck*, sodass du sofort siehst, ob der Ausschnitt richtig liegt.
4. Wenn das Ergebnis passt → Button **„OCR auf alle Seiten“**: derselbe ROI wird auf _alle_ Seiten angewendet; Tabelle & Namen-Liste werden ausgegeben.

*(Excel-Abgleich & Beschriftung kommen in einem späteren Schritt.)*

### requirements.txt
```
streamlit
pymupdf  # fitz
pytesseract
pandas
pillow
```

### packages.txt (nur für Streamlit Cloud)
```
poppler-utils
tesseract-ocr
tesseract-ocr-deu
```
"""

import io
import re
import shutil
from functools import lru_cache
from typing import List, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image, ImageDraw  # ImageDraw neu ➜ Rechteck-Overlay

# ──────────────────────────────────────────────────────────────────────────────
# Tesseract-Pfad (wichtig für Streamlit Cloud)
# ──────────────────────────────────────────────────────────────────────────────
TESSERACT_CMD = shutil.which("tesseract")
if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
else:
    st.error(
        "Tesseract-Executable nicht gefunden. Bitte in **packages.txt** `tesseract-ocr` "
        "und optional `tesseract-ocr-deu` eintragen und App neu starten."
    )
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Page config & Title
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Interaktiver PDF-ROI-Finder", layout="wide")
st.title("📄 PDF-ROI interaktiv bestimmen & OCR")

with st.expander("Kurzanleitung", expanded=False):
    st.markdown(
        """
        1. **PDF hochladen**
        2. Erste Seite wird dargestellt ➜ wähle mit den Reglern links / oben / rechts / unten dein ROI aus.
        3. Vorschau-Bild & Sofort-OCR helfen dir, die Koordinaten anzupassen.
        4. Wenn alles stimmt → *OCR auf alle Seiten*.
        """
    )

# ──────────────────────────────────────────────────────────────────────────────
# Helper – Cache gerenderte Seite, damit Slider-Änderungen schnell bleiben
# ──────────────────────────────────────────────────────────────────────────────
@lru_cache(maxsize=4)
def render_first_page(pdf_bytes: bytes, dpi: int = 300) -> Tuple[Image.Image, int, int]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img, pix.width, pix.height

# ──────────────────────────────────────────────────────────────────────────────
# Upload
# ──────────────────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("PDF hochladen", type=["pdf"], key="pdf")

if pdf_file:
    pdf_bytes = pdf_file.read()

    # Render first page once (cached)
    img, width, height = render_first_page(pdf_bytes)

    st.subheader("1️⃣ ROI wählen (Koordinaten in Pixel – 300 DPI)")
    col1, col2 = st.columns([1, 2])
    with col1:
        st.write("**Bildgröße**:", f"{width} × {height} px")
        x1 = st.number_input("x1 (links)", 0, width - 1, value=st.session_state.get("x1", 200))
        y1 = st.number_input("y1 (oben)", 0, height - 1, value=st.session_state.get("y1", 890))
        x2 = st.number_input("x2 (rechts)", x1 + 1, width, value=st.session_state.get("x2", 560))
        y2 = st.number_input("y2 (unten)", y1 + 1, height, value=st.session_state.get("y2", 980))

        # Store in session_state for convenience
        st.session_state.update({"x1": x1, "y1": y1, "x2": x2, "y2": y2})

    with col2:
        roi = (x1, y1, x2, y2)

        # 🔲 Overlay auf ganzer Seite
        overlay = img.copy()
        draw = ImageDraw.Draw(overlay)
        draw.rectangle(roi, outline="red", width=6)
        st.image(overlay, caption="Seite 1 mit markiertem ROI", use_column_width=True)

        # Preview crop
        crop = img.crop(roi)
        st.image(crop, caption="ROI-Vorschau", use_column_width=True)

        # Sofort-OCR
        ocr_text = pytesseract.image_to_string(crop, lang="deu").strip()
        st.text_area("OCR-Ergebnis (Seite 1, ROI)", ocr_text, height=120)

    # Button to process all pages
    if st.button("🚀 OCR auf alle Seiten", type="primary"):
        with st.spinner("Starte OCR …"):
            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            except Exception as e:
                st.error(f"PDF konnte nicht geöffnet werden: {e}")
                st.stop()

            data: List[Tuple[int, str]] = []
            name_candidates: set[str] = set()
            roi_tuple = (x1, y1, x2, y2)

            for page_index, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=300)
                page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                crop_img = page_img.crop(roi_tuple)
                txt = pytesseract.image_to_string(crop_img, lang="deu").strip()
                data.append((page_index, txt))

                # Regex: Großbuchstaben am Anfang, mind. 2 Zeichen, ggf. Bindestrich
                candidates = re.findall(r"\b[ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]{1,}\b", txt)
                name_candidates.update(candidates)

            df = pd.DataFrame(data, columns=["Seite", "Text (ROI)"])

        st.success("OCR abgeschlossen ✔️")
        st.dataframe(df, use_container_width=True)

        # Download CSV
        buf = io.StringIO()
        df.to_csv(buf, index=False)
        st.download_button(
            "📥 Tabelle als CSV",
            buf.getvalue(),
            file_name="roi_ocr.csv",
            mime="text/csv",
        )

        st.subheader("Potentielle Namen")
        if name_candidates:
            st.write(", ".join(sorted(name_candidates)))
        else:
            st.info("Keine großgeschriebenen Wörter gefunden.")
