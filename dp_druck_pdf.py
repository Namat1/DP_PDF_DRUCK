from __future__ import annotations

import io
import re
import shutil
from datetime import date, datetime, timedelta, time
from functools import lru_cache
from typing import List, Tuple, Dict, Optional

import difflib
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
    st.error("Tesseract‑Executable nicht gefunden. Bitte in **packages.txt** `tesseract-ocr` und optional `tesseract-ocr-deu` eintragen und App neu starten.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Streamlit‑Basics
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")
st.title("📄 Dienstpläne beschriften & verteilen")

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
    s = d + timedelta(days=1)
    return int(s.strftime("%V")), int(s.strftime("%G"))

def format_time(value) -> str:
    if pd.isna(value): return ""
    if isinstance(value, time): return value.strftime("%H:%M")
    if isinstance(value, (datetime, pd.Timestamp)): return value.strftime("%H:%M")
    if isinstance(value, (int, float)):
        total_minutes = round((value % 1) * 1440)
        return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"
    if isinstance(value, str):
        try: return pd.to_datetime(value).strftime("%H:%M")
        except: return value
    return str(value)

def extract_entries(row: pd.Series) -> List[dict]:
    entries = []
    datum = pd.to_datetime(row[14], errors="coerce")
    if pd.isna(datum): return entries

    kw, year = kw_year_sunday(datum)
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())
    datum_lang = f"{weekday}, {datum.strftime('%d.%m.%Y')}"

    base_entry = {
        "KW": kw,
        "Jahr": year,
        "Datum": datum_lang,
        "Datum_raw": datum,
        "Wochentag": weekday,
        "Tour": row[15] if len(row) > 15 else "",
        "Uhrzeit": format_time(row[8]) if len(row) > 8 else "",
        "LKW": row[11] if len(row) > 11 else "",
    }
    if pd.notna(row[3]) and pd.notna(row[4]):
        entries.append({**base_entry, "Name": f"{str(row[3]).strip()} {str(row[4]).strip()}"})
    if pd.notna(row[6]) and pd.notna(row[7]):
        entries.append({**base_entry, "Name": f"{str(row[6]).strip()} {str(row[7]).strip()}"})
    return entries

def normalize_name(name: str) -> str:
    return " ".join(sorted(name.upper().split()))

def fuzzy_match_name(ocr_name: str, excel_names: List[str]) -> str:
    if not ocr_name.strip(): return ""
    normalized_ocr = normalize_name(ocr_name)
    normalized_excel = {name: normalize_name(name) for name in excel_names}
    best = difflib.get_close_matches(normalized_ocr, normalized_excel.values(), n=1, cutoff=0.6)
    if not best: return ""
    return next((orig for orig, norm in normalized_excel.items() if norm == best[0]), "")

def extract_names_from_pdf_by_excel_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = []
    for i, page in enumerate(doc):
        text = page.get_text()
        found = ""
        for name in excel_names:
            if name in text or normalize_name(name) in normalize_name(text):
                found = name
                break
        st.markdown(f"**Seite {i+1} – Gefundener Name:** `{found}`")
        results.append(found)
    doc.close()
    return results

def parse_excel_data(excel_file) -> pd.DataFrame:
    df = pd.read_excel(excel_file, header=None)
    all_entries = []
    for _, row in df.iterrows():
        all_entries.extend(extract_entries(row))
    return pd.DataFrame(all_entries)

def annotate_pdf_with_tours(pdf_bytes: bytes, annotations: List[Optional[Dict[str, str]]]) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page_num, annotation in enumerate(annotations):
        if page_num < len(doc) and annotation:
            page = doc.load_page(page_num)
            text = " - ".join(
                filter(None, [annotation.get("tour", ""), annotation.get("weekday", ""), f"{annotation.get('time', '')} Uhr"])
            )
            rect = page.rect
            box = fitz.Rect(rect.width - 650, rect.height - 60, rect.width - 20, rect.height - 15)
            page.insert_textbox(box, text, fontsize=12, fontname="hebo", color=(1, 0, 0), align=fitz.TEXT_ALIGN_RIGHT)
    output = io.BytesIO()
    doc.save(output)
    doc.close()
    output.seek(0)
    return output.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# Datei‑Uploads
# ──────────────────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("📁 PDF hochladen", type=["pdf"])
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xlsm"])

if not pdf_file:
    st.info("👉 Bitte zuerst ein PDF hochladen.")
    st.stop()

pdf_bytes = pdf_file.read()

verteil_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 OCR & PDF beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel‑Datei hochladen!")
        st.stop()

    with st.spinner("🔍 OCR läuft und Excel wird verarbeitet..."):
        excel_data = parse_excel_data(excel_file)
        filtered_data = excel_data[excel_data['Datum_raw'].dt.date == verteil_date]
        excel_names = filtered_data['Name'].unique().tolist()
        ocr_names = extract_names_from_pdf_by_excel_match(pdf_bytes, excel_names)

    if filtered_data.empty:
        st.warning(f"⚠️ Keine Einträge für {verteil_date.strftime('%d.%m.%Y')} in der Excel-Datei gefunden!")
    else:
        page_annotations = []
        for ocr_name in ocr_names:
            matched = fuzzy_match_name(ocr_name, excel_names)
            if matched:
                row = filtered_data[filtered_data['Name'] == matched].iloc[0]
                page_annotations.append({
                    "matched_name": matched,
                    "tour": str(row['Tour']),
                    "weekday": str(row['Wochentag']),
                    "time": str(row['Uhrzeit'])
                })
            else:
                page_annotations.append(None)

        st.markdown("---")
        st.subheader("🔍 Ergebnis der Zuordnung")
        st.dataframe(pd.DataFrame([{
            "PDF Seite": i + 1,
            "Gefundener Name (OCR)": ocr or "N/A",
            "Zugeordnet (Excel)": a.get("matched_name", "❌ Nein") if a else "❌ Nein",
            "Tour": a.get("tour", "") if a else "",
            "Wochentag": a.get("weekday", "") if a else "",
            "Uhrzeit": a.get("time", "") if a else ""
        } for i, (ocr, a) in enumerate(zip(ocr_names, page_annotations))]), use_container_width=True)

        matched = sum(1 for a in page_annotations if a)
        if matched:
            with st.spinner("📝 PDF wird beschriftet..."):
                result = annotate_pdf_with_tours(pdf_bytes, page_annotations)
            st.download_button("📥 Beschriftete PDF herunterladen", data=result,
                file_name=f"dienstplan_annotiert_{verteil_date.strftime('%Y%m%d')}.pdf", mime="application/pdf")
        else:
            st.error("❌ Keine Zuordnungen gefunden – PDF bleibt unbearbeitet.")

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.6 – Jetzt mit intelligentem Abgleich gegen Excel-Inhalte*")
