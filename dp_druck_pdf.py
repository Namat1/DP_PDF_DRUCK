from __future__ import annotations

"""
Clean Streamlit Utility – PDF‑Dienstplan Matcher
================================================
Minimal‑UI Work‑flow:
--------------------
1. PDF & Excel hochladen
2. ROI definieren
3. Verteilungs­datum wählen
4. OCR → Tour‑Nr. unten rechts annotieren
5. Fertige PDF herunterladen
"""

import io
import re
import shutil
import warnings
from datetime import date, datetime, timedelta
from functools import lru_cache
from typing import List, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image, ImageDraw

# ────────────────────────────────────────────────────────────────
# Suppress noisy library warnings
# ────────────────────────────────────────────────────────────────
warnings.filterwarnings(
    "ignore", category=UserWarning, module="openpyxl.worksheet.header_footer"")

# ────────────────────────────────────────────────────────────────
# Tesseract path (needed on Streamlit Cloud)
# ────────────────────────────────────────────────────────────────
TESS_CMD = shutil.which("tesseract")
if TESS_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESS_CMD
else:
    st.error("Tesseract‑Executable nicht gefunden. Bitte installieren.")
    st.stop()

# ────────────────────────────────────────────────────────────────
# Streamlit basic layout (minimal – no verbose markdown)
# ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")

# ────────────────────────────────────────────────────────────────
# Helper dictionaries & functions
# ────────────────────────────────────────────────────────────────
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
    """KW‑Berechnung mit Sonntag als Wochen­start (ISO + 1 Tag)."""
    s = d + timedelta(days=1)
    return int(s.strftime("%V")), int(s.strftime("%G"))

NAME_PATTERN = re.compile(r"([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)\s+([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)")


def extract_entries(row: pd.Series) -> List[dict]:
    """Extrahiert bis zu zwei Fahrer + Tour etc. aus einer Excel‑Zeile."""
    out: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")
    if pd.isna(datum):
        return out

    kw, year = kw_year_sunday(datum)
    datum_fmt = datum.strftime("%d.%m.%Y")
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())
    datum_lang = f"{weekday}, {datum_fmt}"

    tour = row[15] if len(row) > 15 else ""
    uhrzeit = row[16] if len(row) > 16 else ""
    lkw = row[11] if len(row) > 11 else ""

    def add(name):
        if name:
            out.append(
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

    if pd.notna(row[3]) and pd.notna(row[4]):
        add(f"{str(row[3]).strip()} {str(row[4]).strip()}")
    if pd.notna(row[6]) and pd.notna(row[7]):
        add(f"{str(row[6]).strip()} {str(row[7]).strip()}")

    return out

# ────────────────────────────────────────────────────────────────
# File uploads
# ────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("📑 PDF", type=["pdf"], key="pdf")
excel_file = st.file_uploader("📊 Excel", type=["xlsx", "xlsm"], key="excel")

if not pdf_file:
    st.stop()

pdf_bytes = pdf_file.read()

# ────────────────────────────────────────────────────────────────
# Render first page & ROI selector
# ────────────────────────────────────────────────────────────────
@lru_cache(maxsize=2)
def render_page1(pdf: bytes, dpi: int = 300):
    d = fitz.open(stream=pdf, filetype="pdf")
    p = d.load_page(0)
    pix = p.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img, pix.width, pix.height

page1, W, H = render_page1(pdf_bytes)

st.subheader("ROI wählen")

colA, colB = st.columns([1, 2])

with colA:
    x1 = st.number_input("x1", 0, W - 1, value=200)
    y1 = st.number_input("y1", 0, H - 1, value=890)
    x2 = st.number_input("x2", x1 + 1, W, value=560)
    y2 = st.number_input("y2", y1 + 1, H, value=980)
    roi = (x1, y1, x2, y2)

with colB:
    overlay = page1.copy()
    ImageDraw.Draw(overlay).rectangle(roi, outline="red", width=4)
    st.image(overlay, use_container_width=True)
    st.image(page1.crop(roi), use_container_width=True)

# ────────────────────────────────────────────────────────────────
# Distribution date
# ────────────────────────────────────────────────────────────────
verteil_date: date = st.date_input("Verteilungs­datum", value=date.today())

# ────────────────────────────────────────────────────────────────
# Start processing
# ────────────────────────────────────────────────────────────────
if st.button("Start", type="primary"):
    if not excel_file:
        st.error("Excel fehlt")
        st.stop()

    with st.spinner("Excel lesen …"):
        try:
            xl_df = pd.read_excel(excel_file, engine="openpyxl", header=None)
        except Exception as exc:
            st.error(f"Excel‑Fehler: {exc}")
            st.stop()

    excel_entries: List[dict] = []
    for _, r in xl_df.iterrows():
        excel_entries.extend(extract_entries(r))

    if not excel_entries:
        st.warning("Keine relevanten Daten in der Excel gefunden.")
        st.stop()

    # ── OCR & annotate PDF ───────────────────────────────────────
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    matches = 0

    for i in range(len(pdf_doc)):
        page = pdf_doc.load_page(i)
        try:
            pix = page.get_pixmap(clip=fitz.Rect(*roi))
        except ValueError:
            continue  # ROI außerhalb der Seite → überspringen

        if pix.width == 0 or pix.height == 0:
            continue  # leere ROI

        crop = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        try:
            text = pytesseract.image_to_string(crop, lang="deu")
        except Exception:
            continue  # OCR‑Fehler

        m = NAME_PATTERN.search(text)
        if not m:
            continue

        name = f"{m.group(1)} {m.group(2)}".strip()

        match = next(
            (
                e
                for e in excel_entries
                if e["Name"].lower() == name.lower()
                and e["Datum_raw"].date() == verteil_date
            ),
            None,
        )
        if not match:
            continue

        tour = str(match["Tour"]).strip()
        if not tour:
            continue

        # annotate bottom‑right
        bbox = page.bound()
        dest = fitz.Point(bbox.x1 - 50, bbox.y1 - 20)
        page.insert_text(dest, tour, fontsize=9, fontname="helv", fill=(0, 0, 0))
        matches += 1

    if matches == 0:
        st.warning("Keine Namen‑Tour‑Treffer gefunden.")
        st.stop()

    out = io.BytesIO()
    pdf_doc.save(out)
    st.download_button(
        "PDF herunterladen", data=out.getvalue(), file_name="dienstplaene.pdf", mime="application/pdf"
    )
