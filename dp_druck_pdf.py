"""
Streamlit Utility – PDF-OCR im festen ROI (Namens-Ermittlung)
===========================================================
**Ziel**:  Nur den Text im vordefinierten Ausschnitt aller PDF-Seiten per OCR
extrahieren und anzeigen – *ohne* Excel, *ohne* Annotation. Damit kannst du
prüfen, ob die ROI-Koordinaten korrekt sind und welche Namen überhaupt im
Dokument stehen.  Weitere Schritte (Excel-Abgleich, Beschriftung) bauen wir
später darauf auf.

### Workflow
1. **PDF hochladen**
2. App rendert jede Seite bei 300 DPI, schneidet den festen ROI
   `(99, 426, 280, 488)` aus und führt OCR (Tesseract DE) durch.
3. Ergebnis:
   * Tabelle: Seitenzahl | erkannter Text (ROI)
   * Zusätzliche Liste aller **Großgeschriebenen Wörter** (potentielle Namen) –
     gefiltert auf A–Z-Anfang.
4. CSV-Download der Tabelle möglich.

### ROI
```python
ROI = (left=99, top=426, right=280, bottom=488)  # Pixel @300 DPI
```

### requirements.txt
```
streamlit
pymupdf  # fitz
pytesseract
pandas
pillow
```

### packages.txt (Streamlit Cloud)
```
poppler-utils
tesseract-ocr
tesseract-ocr-deu  # OCR-Sprache DE
```
"""

from __future__ import annotations

import csv
import io
import re
from typing import List, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image

import shutil  # NEW

# -----------------------------------------------------------------------------
# Tesseract-Pfad ermitteln (Streamlit Cloud)
# -----------------------------------------------------------------------------
TESSERACT_CMD = shutil.which("tesseract")
if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
else:
    st.error(
        "Tesseract-Executable nicht gefunden. "
        "Bitte `packages.txt` mit `tesseract-ocr` (und optional `tesseract-ocr-deu`) "
        "anlegen und die App neu deployen."
    )
    st.stop()

# -----------------------------------------------------------------------------
# Fester ROI (Pixel @300 DPI)
# -----------------------------------------------------------------------------
ROI: Tuple[int, int, int, int] = (99, 426, 280, 488)

st.set_page_config(page_title="PDF-OCR ROI", layout="centered")

st.title("📄 PDF-OCR im festen ROI – Namens-Ermittlung")

with st.expander("Anleitung", expanded=False):
    st.markdown(
        f"""
        * **PDF hochladen** – die App liest jede Seite in 300 DPI, schneidet den
          Bereich `x={ROI[0]}:{ROI[2]}`, `y={ROI[1]}:{ROI[3]}` aus und führt
          OCR durch.
        * Das Ergebnis wird als Tabelle angezeigt und kann als CSV
          heruntergeladen werden.
        * Zusätzlich wird eine Liste aller Wörter gezeigt, die mit Großbuchstaben
          beginnen – praktisch als **Potentielle Namen**.
        """
    )

pdf_file = st.file_uploader("PDF hochladen", type=["pdf"], key="pdf")

if pdf_file:
    if st.button("🚀 OCR starten"):
        with st.spinner("Lese PDF …"):
            pdf_bytes = pdf_file.read()
            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            except Exception as e:
                st.error(f"PDF konnte nicht geöffnet werden: {e}")
                st.stop()

            data: List[Tuple[int, str]] = []
            names_candidates: set[str] = set()

            for idx, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                crop = img.crop(ROI)

                text = pytesseract.image_to_string(crop, lang="deu")
                data.append((idx, text.strip()))

                # Namen-Kandidaten: Wörter ab 2 Buchstaben, Großschreibung am Anfang
                # Python-`re` unterstützt keine Unicode-Property-Escapes (\p{{Lu}}),
                # daher explizite Zeichenklasse für DE + A–Z.
                candidates = re.findall(r"\b[ÄÖÜA-Z][ÄÖÜA-Za-zäöüß]{1,}\b", text)
                names_candidates.update(candidates)

            df = pd.DataFrame(data, columns=["Seite", "Text (ROI)"])

        st.success("OCR abgeschlossen")
        st.dataframe(df, use_container_width=True)

        # CSV-Download
        csv_buf = io.StringIO()
        df.to_csv(csv_buf, index=False)
        st.download_button(
            "📥 Tabelle als CSV herunterladen",
            csv_buf.getvalue(),
            file_name="roi_ocr.csv",
            mime="text/csv",
        )

        # Namen-Kandidaten
        st.subheader("Potentielle Namen (Großschreibung)")
        if names_candidates:
            st.write(", ".join(sorted(names_candidates)))
        else:
            st.info("Keine großgeschriebenen Wörter gefunden.")
else:
    st.info("Bitte zunächst ein PDF hochladen.")
