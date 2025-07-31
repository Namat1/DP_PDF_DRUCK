from __future__ import annotations

"""
Streamlit Utility – PDF-Dienstplan Matcher
========================================
End-to-End-Workflow:
-------------------
1. **Dateien hochladen** – PDF (gescannter Dienstplan) **und** die zugehörige Excel-Tabelle.
2. **ROI festlegen** – Rechteck, in dem auf jeder PDF-Seite der Fahrername steht.
3. **OCR** – Namen pro Seite auslesen & zwischenspeichern.
4. **Excel parsen** – Fahrer + Datum + Tour-Nr. extrahieren.
5. **Verteilungsdatum wählen**.
6. **Match & Annotate** – Namen ↔︎ Excel-Zeilen verbinden, Tour-Nr. unten rechts auf jede PDF-Seite schreiben.
7. **Download** der beschrifteten PDF.

### Python-Pakete (requirements.txt)
```
streamlit
pymupdf      # fitz
pytesseract
pandas
pillow
openpyxl
```

### System-Pakete (packages.txt – Streamlit Cloud)
```
poppler-utils
pytesseract-ocr
pytesseract-ocr-deu
```
"""

import io
import re
import shutil
from datetime import datetime, timedelta, date
from functools import lru_cache
from typing import List, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image, ImageDraw

# ──────────────────────────────────────────────────────────────────────────────
# Tesseract – Pfad setzen (wichtig für Streamlit Cloud)
# ──────────────────────────────────────────────────────────────────────────────
TESS_CMD = shutil.which("tesseract")
if TESS_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESS_CMD
else:
    st.error(
        "Tesseract-Executable nicht gefunden. Bitte in **packages.txt** `tesseract-ocr` "
        "und optional `tesseract-ocr-deu` eintragen und App neu starten."
    )
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Streamlit Basics
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")
st.title("📄 Dienstpläne beschriften & verteilen")

with st.expander("Kurze Anleitung", expanded=False):
    st.markdown(
        """
        **Workflow**
        1. PDF & Excel hochladen.
        2. ROI auf Seite 1 definieren → Vorschau prüfen.
        3. Verteilungs-Datum auswählen.
        4. *OCR & Annotate* starten → fertige PDF herunterladen.
        """
    )

# ──────────────────────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────────────────────────────────────
WEEKDAYS_DE = {
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag",
}

def kw_year_sunday(d: datetime) -> Tuple[int, int]:
    """KW-Berechnung mit **Sonntag** als Wochen-Start (ISO +1 Tag)."""
    s = d + timedelta(days=1)
    return int(s.strftime("%V")), int(s.strftime("%G"))

def extract_entries(row: pd.Series) -> List[dict]:
    """Liest bis zu **2 Fahrer** aus einer Excel-Zeile (Spalten hart codiert)."""
    entries: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
    if pd.isna(datum):
        return entries

    kw, year = kw_year_sunday(datum)
    datum_fmt = datum.strftime("%d.%m.%Y")
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())
    datum_lang = f"{weekday}, {datum_fmt}"

    tour = row[15] if len(row) > 15 else ""
    uhrzeit = row[16] if len(row) > 16 else ""
    lkw = row[11] if len(row) > 11 else ""

    # Fahrer 1 (D,E)
    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        entries.append(
            {
                "KW": kw,
                "Jahr": year,
                "Datum": datum_lang,
                "Datum_raw": datum,
                "Name": name,
                "Tour": tour,
                "Uhrzeit": uhrzeit,
                "LKW": lkw,
            }
        )
    # Fahrer 2 (G,H)
    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        entries.append(
            {
                "KW": kw,
                "Jahr": year,
                "Datum": datum_lang,
                "Datum_raw": datum,
                "Name": name,
                "Tour": tour,
                "Uhrzeit": uhrzeit,
                "LKW": lkw,
            }
        )
    return entries

# OCR-Regex – zwei **aufeinander­folgende** Großbuchstaben-Wörter → Vor- & Nachname
NAME_PATTERN = re.compile(r"([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)\s+([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)")

# ──────────────────────────────────────────────────────────────────────────────
# Datei-Uploads
# ──────────────────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("📑 PDF hochladen", type=["pdf"], key="pdf")
excel_file = st.file_uploader("📊 Excel hochladen", type=["xlsx", "xlsm"], key="excel")

if not pdf_file:
    st.info("👉 Bitte zuerst ein PDF hochladen.")
    st.stop()

pdf_bytes = pdf_file.read()

# ──────────────────────────────────────────────────────────────────────────────
# Seite 1 rendern & ROI auswählen
# ──────────────────────────────────────────────────────────────────────────────
@lru_cache(maxsize=2)
def render_page1(pdf: bytes, dpi: int = 300):
    d = fitz.open(stream=pdf, filetype="pdf")
    p = d.load_page(0)
    pix = p.get_pixmap(dpi=dpi)
    pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pil, pix.width, pix.height

page1_img, W, H = render_page1(pdf_bytes)

st.subheader("1️⃣ ROI definieren")
colA, colB = st.columns([1, 2])

with colA:
    st.write("**Seite 1 Größe:**", f"{W} × {H} px")
    x1 = st.number_input("x1 (links)", 0, W - 1, value=st.session_state.get("x1", 200))
    y1 = st.number_input("y1 (oben)", 0, H - 1, value=st.session_state.get("y1", 890))
    x2 = st.number_input("x2 (rechts)", x1 + 1, W, value=st.session_state.get("x2", 560))
    y2 = st.number_input("y2 (unten)", y1 + 1, H, value=st.session_state.get("y2", 980))
    st.session_state.update({"x1": x1, "y1": y1, "x2": x2, "y2": y2})

with colB:
    roi_box = (x1, y1, x2, y2)
    overlay = page1_img.copy()
    ImageDraw.Draw(overlay).rectangle(roi_box, outline="red", width=5)
    st.image(overlay, caption="Seite 1 mit ROI", use_column_width=True)
    st.image(page1_img.crop(roi_box), caption="ROI-Vorschau", use_column_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# Verteilungs-Datum (vom Nutzer bestimmen lassen)
# ──────────────────────────────────────────────────────────────────────────────
verteil_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today())

# ──────────────────────────────────────────────────────────────────────────────
# Haupt-Button – OCR, Excel, Match & Annotate
# ──────────────────────────────────────────────────────────────────────────────
if st.button("🚀 OCR & PDF beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel-Datei hoch
