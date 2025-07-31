"""
Streamlit App: PDF-Namenssuche & Annotation (Fixe ROI) + Vorschau der Namen
============================================================================
Die App zeigt **sofort nach dem Hochladen der Excel-Datei** die Liste aller
Namen an, die später im PDF gesucht werden. Erst danach kann der Nutzer auf
**Starten** klicken, um die PDF zu verarbeiten.

### Workflow
1. PDF + Excel hochladen  
2. **Liste der Namen aus Excel wird angezeigt**  
3. Parameter anpassen (Spalten, Textposition, Groß/Klein)  
4. **Starten** → OCR in fixem ROI `(99, 426, 280, 488)`  
5. Gefundene Namen & annotierte PDF herunterladen

### Fester OCR-Bereich
```python
ROI = (99, 426, 280, 488)  # Pixel bei 300 DPI
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

### packages.txt (Streamlit Cloud)
```
tesseract-ocr
tesseract-ocr-deu  # optional: deutsches Sprachpaket
```
"""

from __future__ import annotations

import io
import re
from typing import Dict, Set, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image

# -----------------------------------------------------------------------------
# Fester ROI (Pixel @300 DPI)
# -----------------------------------------------------------------------------
ROI: Tuple[int, int, int, int] = (99, 426, 280, 488)

# -----------------------------------------------------------------------------
# UI-Setup
# -----------------------------------------------------------------------------
st.set_page_config(page_title="PDF-Namenssuche (Fixe ROI)", layout="centered")

st.title("🔍 PDF-Namenssuche & Annotation – Fixe ROI")

with st.expander("Anleitung", expanded=False):
    st.markdown(
        f"""
        1. **PDF hochladen** – mehrseitig, gescannt oder als Bild-PDF.
        2. **Excel hochladen** mit Spalten *Name* und Wert (z. B. *Abteilung*).
        3. Die App zeigt sofort **alle Namen** aus der gewählten Spalte an.
        4. Klick **Starten**, um nur den Bereich `x={ROI[0]}:{ROI[2]}, y={ROI[1]}:{ROI[3]}`
           (Pixel in 300 DPI) jeder Seite per OCR zu durchsuchen.
        5. Treffer werden in die PDF geschrieben, die annotierte Datei steht
           anschließend zum Download bereit.
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

    st.subheader("1️⃣ Spalten & Parameter wählen")
    col1, col2 = st.columns(2)
    with col1:
        name_col = st.selectbox("Spalte mit Namen", df.columns)
    with col2:
        value_col = st.selectbox(
            "Spalte mit einzutragendem Wert",
            df.columns,
            index=min(1, len(df.columns) - 1),
        )

    # ------------------------ Vorschau der Namen -----------------------------
    names_preview = df[name_col].dropna().unique()
    st.subheader("2️⃣ Namen in der Excel-Datei")
    if names_preview.size:
        st.write(f"Es werden **{len(names_preview)}** Namen gesucht:")
        st.table(pd.DataFrame(sorted(names_preview), columns=["Name"]))
    else:
        st.error("In der gewählten Spalte wurden keine Namen gefunden.")
        st.stop()

    # ---------------------- weitere Parameter -------------------------------
    x_position = st.number_input("X-Position (Pt)", 0, 600, value=50)
    y_position = st.number_input("Y-Position (Pt)", 0, 800, value=50)
    font_size = st.number_input("Schriftgröße (Pt)", 6, 48, value=12)
    case_sensitive = st.checkbox("Groß-/Kleinschreibung beachten", value=False)

    # Mapping Suche → Originalname + Wert
    search_to_original: Dict[str, str] = {}
    name_map: Dict[str, str] = {}
    for _, r in df.iterrows():
        if pd.notna(r[name_col]):
            original = str(r[name_col])
            key = original if case_sensitive else original.lower()
            search_to_original[key] = original
            name_map[key] = r[value_col]

    if st.button("🚀 Starten"):
        with st.spinner("Verarbeite PDF… bitte warten"):
            pdf_bytes = pdf_file.read()
            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            except Exception as e:
                st.error(f"PDF konnte nicht geöffnet werden: {e}")
                st.stop()

            found_names: Set[str] = set()

            for page in doc:
                # 300 DPI-Bild
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                crop = img.crop(ROI)

                ocr_text = pytesseract.image_to_string(crop, lang="deu")
                search_space = ocr_text if case_sensitive else ocr_text.lower()

                for key, value in name_map.items():
                    if re.search(rf"\b{re.escape(key)}\b", search_space):
                        page.insert_text(
                            (x_position, y_position),
                            str(value),
                            fontsize=font_size,
                            fontname="helv",
                            fill=(0, 0, 0),
                        )
                        found_names.add(search_to_original[key])
                        break

            buf = io.BytesIO()
            doc.save(buf)
            doc.close()
            buf.seek(0)

        st.success("Fertig! Die PDF ist annotiert.")
        st.download_button("📥 Annotierte PDF herunterladen", buf, file_name="annotiert.pdf", mime="application/pdf")

        st.subheader("Gefundene Namen im PDF")
        if found_names:
            st.table(pd.DataFrame(sorted(found_names), columns=["Name"]))
        else:
            st.info("Es wurden keine Namen im PDF gefunden.")
else:
    st.info("Bitte PDF **und** Excel hochladen, um zu starten.")
