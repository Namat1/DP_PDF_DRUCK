from __future__ import annotations

import io
import re
import shutil
from datetime import date, datetime, timedelta, time
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
    st.error("Tesseract‑Executable nicht gefunden. Bitte in packages.txt tesseract-ocr eintragen.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")
st.title("📄 Dienstpläne beschriften & verteilen")

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
    if pd.isna(value):
        return ""
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%H:%M")
    if isinstance(value, (int, float)):
        total_minutes = round((value % 1) * 1440)
        return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"
    if isinstance(value, str):
        try:
            return pd.to_datetime(value).strftime("%H:%M")
        except Exception:
            return value
    return str(value)

def extract_entries(row: pd.Series) -> List[dict]:
    entries = []
    datum = pd.to_datetime(row[14], errors="coerce")
    if pd.isna(datum):
        return entries

    kw, year = kw_year_sunday(datum)
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())
    datum_lang = f"{weekday}, {datum.strftime('%d.%m.%Y')}"

    tour = row[15] if len(row) > 15 else ""
    uhrzeit = format_time(row[8]) if len(row) > 8 else ""
    lkw = row[11] if len(row) > 11 else ""

    base_entry = {
        "KW": kw,
        "Jahr": year,
        "Datum": datum_lang,
        "Datum_raw": datum,
        "Wochentag": weekday,
        "Tour": tour,
        "Uhrzeit": uhrzeit,
        "LKW": lkw,
    }

    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        entry1 = base_entry.copy()
        entry1["Name"] = name
        entries.append(entry1)

    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        entry2 = base_entry.copy()
        entry2["Name"] = name
        entries.append(entry2)

    return entries

def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.upper().strip())

def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = []
    normalized_excel_names = [normalize_name(name) for name in excel_names]

    for i, page in enumerate(doc):
        text = page.get_text()
        found_name = ""
        for word in text.split():
            for orig_name, norm_excel in zip(excel_names, normalized_excel_names):
                if normalize_name(word) in norm_excel:
                    found_name = orig_name
                    break
            if found_name:
                break
        st.markdown(f"**Seite {i+1} – Gefundener Name:** `{found_name}`")
        results.append(found_name)

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
                filter(None, [annotation.get("tour"), annotation.get("weekday"), annotation.get("time") + " Uhr"])
            )
            rect = page.rect
            text_rect = fitz.Rect(rect.width - 650, rect.height - 60, rect.width - 20, rect.height - 15)
            page.insert_textbox(text_rect, text, fontsize=12, fontname="hebo", color=(1, 0, 0), align=fitz.TEXT_ALIGN_RIGHT)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("📑 PDF hochladen", type=["pdf"])
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xlsm"])
if not pdf_file:
    st.info("👉 Bitte zuerst ein PDF hochladen.")
    st.stop()
pdf_bytes = pdf_file.read()

verteil_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 PDF analysieren & beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel‑Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel-Daten laden & Namen extrahieren..."):
        excel_data = parse_excel_data(excel_file)
        kw, jahr = kw_year_sunday(verteil_date)
        filtered_data = excel_data[(excel_data['KW'] == kw) & (excel_data['Jahr'] == jahr)]

    if filtered_data.empty:
        st.warning(f"⚠️ Keine Einträge für KW {kw} ({verteil_date.strftime('%d.%m.%Y')}) in der Excel-Datei gefunden!")
    else:
        excel_names = filtered_data['Name'].unique().tolist()
        ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names)

        page_annotations = []
        for ocr_name in ocr_names:
            matched_row = filtered_data[filtered_data['Name'] == ocr_name]
            if not matched_row.empty:
                entry = matched_row.iloc[0]
                page_annotations.append({
                    "matched_name": ocr_name,
                    "tour": str(entry['Tour']),
                    "weekday": str(entry['Wochentag']),
                    "time": str(entry['Uhrzeit'])
                })
            else:
                page_annotations.append(None)

        display_data = []
        for i, (ocr_name, annotation) in enumerate(zip(ocr_names, page_annotations)):
            display_data.append({
                "Seite": i + 1,
                "Gefundener Name": ocr_name or "❌ Nicht erkannt",
                "Zugeordnet": annotation["matched_name"] if annotation else "❌ Nein",
                "Tour": annotation["tour"] if annotation else "",
                "Wochentag": annotation["weekday"] if annotation else "",
                "Uhrzeit": annotation["time"] if annotation else ""
            })

        st.dataframe(pd.DataFrame(display_data), use_container_width=True)

        if any(page_annotations):
            st.success("✅ Übereinstimmungen gefunden. PDF wird beschriftet...")
            annotated_pdf = annotate_pdf_with_tours(pdf_bytes, page_annotations)
            st.download_button("📥 Beschriftete PDF herunterladen", data=annotated_pdf, file_name="dienstplan_annotiert.pdf", mime="application/pdf")
        else:
            st.error("Keine passenden Namen im PDF gefunden.")

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.7 – Volltextsuche nach Namen aktiviert*")
