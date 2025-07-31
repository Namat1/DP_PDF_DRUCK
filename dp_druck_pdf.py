"""
Streamlit App: PDF‑Namenssuche & Annotation (Fixe ROI)
=====================================================
Dieses Skript ist **deploy‑fertig für streamlit.io**.  Es nutzt einen **fest
verdrahteten OCR‑Ausschnitt** und überspringt jegliche Canvas‑Interaktion.

### Ablauf
1. **PDF** hochladen (mehrseitig, gescannt oder Bild‑PDF).  
2. **Excel** hochladen mit Spalten `Name` und Wert‑Spalte (z. B. `Abteilung`).  
3. Das Skript liest jede Seite im **fixen ROI** `(left=99, top=426, right=280, bottom=488)`
   per Tesseract‑OCR aus, sucht nach Namen und druckt den zugehörigen Wert auf
   dieselbe Seite.  
4. Annotierte PDF steht zum Download bereit.

### Fester OCR‑Bereich
```python
ROI = (LEFT, TOP, RIGHT, BOTTOM) = (99, 426, 280, 488)  # Pixel bei 300 DPI
```

### requirements.txt
```
streamlit
pymupdf  # fitz
pandas
pytesseract
pillow
openpyxl
```

### packages.txt (Streamlit Cloud)
```
tesseract-ocr
# deutsches Sprachpaket (empfohlen):
tesseract-ocr-deu
```

> **Kein Poppler nötig** – Rendering und Manipulation übernimmt **PyMuPDF**.
"""

from __future__ import annotations

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
import io
import re
from typing import Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image

# -----------------------------------------------------------------------------
# Feste ROI‑Koordinaten (Pixel bezogen auf 300 DPI‑Renderebene)
# -----------------------------------------------------------------------------
ROI: Tuple[int, int, int, int] = (99, 426, 280, 488)  # (left, top, right, bottom)

# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------
st.set_page_config(page_title="PDF‑Namenssuche (Fixe ROI)", layout="centered")

st.title("🔍 PDF‑Namenssuche & Annotation – Fixe ROI")

with st.expander("Anleitung", expanded=False):
    st.markdown(
        f"""
        1. **PDF hochladen** – mehrseitig, gescannt oder als Bild‑PDF.
        2. **Excel hochladen** mit Spalten *Name* und Wert (z. B. *Abteilung*).
        3. Das Script durchsucht **nur** den Bereich
           `x = {ROI[0]}:{ROI[2]}`, `y = {ROI[1]}:{ROI[3]}` (Pixel in 300 DPI)
           jeder Seite nach den Namen.
        4. Wird ein Name gefunden, wird der Wert auf die Seite geschrieben und
           du erhältst eine annotierte PDF zum Download.
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

    st.subheader("Parameter wählen")
    col1, col2 = st.columns(2)
    with col1:
        name_col = st.selectbox("Spalte mit Namen", df.columns)
    with col2:
        value_col = st.selectbox("Spalte mit einzutragendem Wert", df.columns, index=min(1, len(df.columns) - 1))

    x_position = st.number_input("X‑Position (Pt)", 0, 600, value=50)
    y_position = st.number_input("Y‑Position (Pt)", 0, 800, value=50)
    font_size = st.number_input("Schriftgröße (Pt)", 6, 48, value=12)

    case_sensitive = st.checkbox("Groß-/Kleinschreibung beachten", value=False)

    # Mapping Name → Wert
    name_map = {
        (str(r[name_col]) if case_sensitive else str(r[name_col]).lower()): r[value_col]
        for _, r in df.iterrows() if pd.notna(r[name_col])
    }

    if st.button("🚀 Starten"):
        with st.spinner("Verarbeite PDF… bitte warten"):
            pdf_bytes = pdf_file.read()
            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            except Exception as e:
                st.error(f"PDF konnte nicht geöffnet werden: {e}")
                st.stop()

            # Durch jede Seite iterieren
            for page_idx in range(len(doc)):
                page = doc[page_idx]

                # Seite als 300 DPI‑Bild
                pix = page.get_pixmap(dpi=300)
                page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # ROI ausscheiden
                crop_img = page_img.crop(ROI)

                # OCR im ROI
                ocr_text = pytesseract.image_to_string(crop_img, lang="deu")
                search_space = ocr_text if case_sensitive else ocr_text.lower()

                # Namen suchen
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
                        break  # nur erster Treffer pro Seite

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

else:
    st.info("Bitte PDF und Excel hochladen, um zu starten.")
