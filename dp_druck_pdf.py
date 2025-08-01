from __future__ import annotations



import io
import re
import shutil
from datetime import date, datetime, timedelta, time
from typing import List, Tuple, Dict, Optional

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# Tesseract – Pfad setzen (wichtig für Streamlit Cloud)
# ──────────────────────────────────────────────────────────────────────────────
TESS_CMD = shutil.which("tesseract")
if TESS_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESS_CMD
else:
    st.error("Tesseract‑Executable nicht gefunden. Bitte in **packages.txt** `tesseract-ocr` eintragen.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")
st.title("📄 Dienstpläne beschriften & verteilen (Multi‑PDF)")

# ──────────────────────────────────────────────────────────────────────────────
WEEKDAYS_DE: Dict[str, str] = {
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag",
}


def kw_year_sunday(d: datetime) -> Tuple[int, int]:
    """Ermittelt Kalenderwoche & Jahr basierend auf einer Woche, die *Sonntag* beginnt."""
    s = d + timedelta(days=1)  # Montag‑ISO‑Offset → Sonntag‑System
    return int(s.strftime("%V")), int(s.strftime("%G"))


def format_time(value) -> str:
    """Zahl, Excel‑Serial, Timestamp oder Time → `HH:MM` String."""
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
    """Zieht aus *einer Zeile* alle (ggf. zwei) Fahrer‑Einträge heraus."""
    entries: List[dict] = []

    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O (15) laut User‑Layout
    if pd.isna(datum):
        return entries

    kw, year = kw_year_sunday(datum)
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())

    base_entry = {
        "KW": kw,
        "Jahr": year,
        "Datum": f"{weekday}, {datum.strftime('%d.%m.%Y')}",
        "Datum_raw": datum,
        "Wochentag": weekday,
        "Tour": row[15] if len(row) > 15 else "",  # Spalte P (16)
        "Uhrzeit": format_time(row[8]) if len(row) > 8 else "",  # Spalte I (9)
        "LKW": row[11] if len(row) > 11 else "",  # Spalte L (12)
    }

    # 1. Fahrer (Spalten D+E)
    if pd.notna(row[3]) and pd.notna(row[4]):
        entries.append({**base_entry, "Name": f"{str(row[3]).strip()} {str(row[4]).strip()}"})

    # 2. Fahrer (Spalten G+H)
    if pd.notna(row[6]) and pd.notna(row[7]):
        entries.append({**base_entry, "Name": f"{str(row[6]).strip()} {str(row[7]).strip()}"})

    return entries


def normalize_name(name: str) -> str:
    """Groß‑/Kleinschreibung & Whitespaces egalisieren."""
    return re.sub(r"\s+", " ", name.upper().strip())


def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """Versucht, für **jede Seite** den Namen (ein Wort genügt) zu finden."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []

    normalized_excel_names = [normalize_name(n) for n in excel_names]

    for page_idx, page in enumerate(doc, start=1):
        text = page.get_text()
        found_name = ""
        for word in text.split():
            for orig_name, norm_excel in zip(excel_names, normalized_excel_names):
                if normalize_name(word) in norm_excel:
                    found_name = orig_name
                    break
            if found_name:
                break
        st.markdown(f"**Seite {page_idx} – Gefundener Name:** `{found_name or '❌ N/A'}`")
        results.append(found_name)

    doc.close()
    return results


def parse_excel_data(excel_file) -> pd.DataFrame:
    """Liest Excel *ohne* Header (User‑Layout) → DataFrame normalisiert."""
    df = pd.read_excel(excel_file, header=None)
    all_entries: List[dict] = []
    for _, row in df.iterrows():
        all_entries.extend(extract_entries(row))
    return pd.DataFrame(all_entries)


def annotate_pdf_with_tours(pdf_bytes: bytes, annotations: List[Optional[Dict[str, str]]]) -> bytes:
    """Beschriftet jede Seite mit Tour, Wochentag & Zeit (unten rechts)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page_num, annotation in enumerate(annotations):
        if page_num >= len(doc):
            break
        if not annotation:
            continue

        page = doc.load_page(page_num)
        text_parts = [annotation.get("tour"), annotation.get("weekday"), annotation.get("time")]
        text = " - ".join(filter(None, text_parts))
        if not text:
            continue

        rect = page.rect
        box = fitz.Rect(rect.width - 650, rect.height - 60, rect.width - 20, rect.height - 15)
        page.insert_textbox(box, text, fontsize=12, fontname="helv", color=(1, 0, 0), align=fitz.TEXT_ALIGN_RIGHT)

    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def merge_annotated_pdfs(buffers: List[bytes]) -> bytes:
    """Alle annotierten PDFs nacheinander in **eine** PDF zusammenführen."""
    if not buffers:
        return b""

    base_doc = fitz.open(stream=buffers[0], filetype="pdf")
    for extra in buffers[1:]:
        tmp = fitz.open(stream=extra, filetype="pdf")
        base_doc.insert_pdf(tmp)
        tmp.close()

    merged = io.BytesIO()
    base_doc.save(merged)
    base_doc.close()
    return merged.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# 🔽 UI – Uploads & Eingaben
# ──────────────────────────────────────────────────────────────────────────────
\pdf_files = st.file_uploader("📑 PDFs hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan‑Excel hochladen", type=["xlsx", "xls", "xlsm"])

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF‑Dateien hochladen.")
    st.stop()

verteil_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY")

run = st.button("🚀 PDFs analysieren, beschriften & zusammenführen", type="primary")

if run:
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel‑Datei hochladen!")
        st.stop()

    # Excel‑Daten einlesen & auf KW/Jahr filtern
    with st.spinner("🔍 Excel‑Daten laden & Namen extrahieren …"):
        excel_df = parse_excel_data(excel_file)
        kw, yr = kw_year_sunday(verteil_date)
        filtered = excel_df[(excel_df["KW"] == kw) & (excel_df["Jahr"] == yr)]

    if filtered.empty:
        st.warning(f"⚠️ Keine Einträge für KW {kw} ({verteil_date:%d.%m.%Y}) in der Excel‑Datei gefunden!")
        st.stop()

    excel_names = filtered["Name"].unique().tolist()

    display_rows: List[Dict[str, str]] = []
    annotated_buffers: List[bytes] = []

    # ──────────────────── jede PDF einzeln verarbeiten ────────────────────
    for pdf_idx, pdf_file in enumerate(pdf_files, start=1):
        pdf_bytes = pdf_file.read()
        ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names)

        page_annots: List[Optional[Dict[str, str]]] = []
        for ocr_name in ocr_names:
            match = filtered[filtered["Name"] == ocr_name]
            if not match.empty:
                row = match.iloc[0]
                page_annots.append({
                    "matched_name": ocr_name,
                    "tour": str(row["Tour"]),
                    "weekday": str(row["Wochentag"]),
                    "time": str(row["Uhrzeit"]),
                })
            else:
                page_annots.append(None)

        # Tabelle für UI
        for page_no, (ocr_name, annot) in enumerate(zip(ocr_names, page_annots), start=1):
            display_rows.append({
                "PDF": pdf_file.name,
                "Seite": page_no,
                "Gefundener Name": ocr_name or "❌ Nicht erkannt",
                "Zugeordnet": annot["matched_name"] if annot else "❌ Nein",
                "Tour": annot["tour"] if annot else "",
                "Wochentag": annot["weekday"] if annot else "",
                "Uhrzeit": annot["time"] if annot else "",
            })

        # Annotierte PDF‑Bytes sammeln
        annotated_buffers.append(annotate_pdf_with_tours(pdf_bytes, page_annots))

    # ───────────────────────────────────────────────────────────────────────
    st.dataframe(pd.DataFrame(display_rows), use_container_width=True)

    if annotated_buffers:
        st.success("✅ Alle PDFs beschriftet. Dateien werden zusammengeführt …")
        final_pdf = merge_annotated_pdfs(annotated_buffers)
        st.download_button(
            "📥 Gesamte beschriftete PDF herunterladen",
            data=final_pdf,
            file_name="dienstplaene_annotiert_gesamt.pdf",
            mime="application/pdf",
        )
    else:
        st.error("🚫 Keine Übereinstimmungen – keine beschriftete Datei erzeugt.")

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.8 – Mehrfach‑PDF‑Support*")
