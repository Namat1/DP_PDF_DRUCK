from __future__ import annotations

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

# ────────────────────────────────────────────────────────────────
# Tesseract Pfad
# ────────────────────────────────────────────────────────────────
TESS_CMD = shutil.which("tesseract")
if TESS_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESS_CMD
else:
    st.error("Tesseract nicht gefunden – bitte installieren.")
    st.stop()

# ────────────────────────────────────────────────────────────────
# Streamlit‑Grundlayout (ohne Zusatz‑Markdown)
# ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Dienstplan", layout="wide")

# ────────────────────────────────────────────────────────────────
# Dateiupload
# ────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("PDF", type=["pdf"])
excel_file = st.file_uploader("Excel", type=["xlsx", "xlsm"])
if not pdf_file:
    st.stop()

pdf_bytes = pdf_file.read()

# ────────────────────────────────────────────────────────────────
# Helper
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
    s = d + timedelta(days=1)
    return int(s.strftime("%V")), int(s.strftime("%G"))

NAME_PATTERN = re.compile(r"([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)\s+([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)")


def extract_entries(row: pd.Series) -> List[dict]:
    entries: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
    if pd.isna(datum):
        return entries

    kw, year = kw_year_sunday(datum)
    datum_fmt = datum.strftime("%d.%m.%Y")
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())
    datum_lang = f"{weekday}, {datum_fmt}"

    tour = row[15] if len(row) > 15 else ""
    uhrzeit = row[16] if len(row) > 16 else ""
    lkw = row[11] if len(row) > 11 else ""

    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        entries.append({
            "KW": kw,
            "Jahr": year,
            "Datum": datum_lang,
            "Datum_raw": datum,
            "Name": name,
            "Tour": tour,
            "Uhrzeit": uhrzeit,
            "LKW": lkw,
        })
    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        entries.append({
            "KW": kw,
            "Jahr": year,
            "Datum": datum_lang,
            "Datum_raw": datum,
            "Name": name,
            "Tour": tour,
            "Uhrzeit": uhrzeit,
            "LKW": lkw,
        })
    return entries

# ────────────────────────────────────────────────────────────────
# PDF Render & ROI
# ────────────────────────────────────────────────────────────────
@lru_cache(maxsize=2)
def render_page1(pdf: bytes, dpi: int = 300):
    doc = fitz.open(stream=pdf, filetype="pdf")
    p = doc.load_page(0)
    pix = p.get_pixmap(dpi=dpi)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples), pix.width, pix.height

img, W, H = render_page1(pdf_bytes)

x1 = st.number_input("x1", 0, W - 1, 200)
y1 = st.number_input("y1", 0, H - 1, 890)
x2 = st.number_input("x2", x1 + 1, W, 560)
y2 = st.number_input("y2", y1 + 1, H, 980)
roi = (x1, y1, x2, y2)

ov = img.copy()
ImageDraw.Draw(ov).rectangle(roi, outline="red", width=3)
st.image(ov, use_column_width=True)

verteil_date = st.date_input("Verteilungsdatum", value=date.today())

# ────────────────────────────────────────────────────────────────
# Hauptablauf
# ────────────────────────────────────────────────────────────────
if st.button("Start"):
    if not excel_file:
        st.error("Excel fehlt")
        st.stop()

    # Excel verarbeiten
    xl_df = pd.read_excel(excel_file, engine="openpyxl", header=None)
    excel_entries: List[dict] = []
    for _, r in xl_df.iterrows():
        excel_entries.extend(extract_entries(r))

    if not excel_entries:
        st.error("Keine Daten in der Excel gefunden")
        st.stop()

    # PDF OCR + Annotation
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    matches = 0

    for page_index in range(len(pdf_doc)):
        page = pdf_doc.load_page(page_index)
        pix = page.get_pixmap(clip=fitz.Rect(*roi))
        crop_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = pytesseract.image_to_string(crop_img, lang="deu")
        m = NAME_PATTERN.search(text)
        if not m:
            continue
        name_ocr = f"{m.group(1)} {m.group(2)}".strip()

        # passenden Eintrag suchen (gleicher Name & Datum)
        page_date = verteil_date
        match = next((e for e in excel_entries if e["Name"].lower() == name_ocr.lower() and e["Datum_raw"].date() == page_date), None)
        if not match:
            continue

        tour_text = str(match["Tour"]).strip()
        if not tour_text:
            continue

        # Annotation unten rechts
        bbox = page.bound()
        text_point = fitz.Point(bbox.x1 - 50, bbox.y1 - 20)
        page.insert_text(text_point, tour_text, fontname="helv", fontsize=10, fontfile=None, fill=(0, 0, 0), render_mode=3)
        matches += 1

    if matches == 0:
        st.warning("Keine Übereinstimmungen gefunden")
        st.stop()

    # PDF speichern
    output = io.BytesIO()
    pdf_doc.save(output)
    st.download_button("PDF herunterladen", data=output.getvalue(), file_name="dienstplaene.pdf", mime="application/pdf")
