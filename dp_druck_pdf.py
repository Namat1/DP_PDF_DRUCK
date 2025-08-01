from __future__ import annotations

"""
PDF Dienstplan Matcher – v1.8 (Multi‑PDF‑Support)
=================================================
• Lädt **beliebig viele PDF‑Dienstpläne** gleichzeitig.
• Vergleicht OCR‑er­kannte Namen pro Seite mit einem hochgeladenen Tourplan‑Excel.
• Beschriftet jede Seite mit Tour‑Nr., Wochentag und Uhrzeit.
• Fügt alle beschrifteten PDFs zu **einer einzigen Datei** zusammen, die direkt heruntergeladen werden kann.

Eingabedaten (laut User‑Layout):
────────────────────────────────
Excel‑Spalten (0‑basiert):
  3 = Nachname 1   |  4 = Vorname 1
  6 = Nachname 2   |  7 = Vorname 2
  8 = Uhrzeit      | 11 = LKW      | 14 = Datum   | 15 = Tour
PDF: ein oder mehrere Dateien, jeweils **eine Seite pro Fahrer**.
Die Kalenderwoche zählt **Sonntag‑bis‑Samstag** (FUHRPARK‑System).
"""

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


# ──────────────────────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────────────────────────────────────

def kw_year_sunday(d: datetime) -> Tuple[int, int]:
    """Kalenderwoche & Jahr berechnen – Woche startet Sonntag."""
    s = d + timedelta(days=1)  # ISO -> Sonntag‑Offset
    return int(s.strftime("%V")), int(s.strftime("%G"))


def format_time(value) -> str:
    """Zahl, Excel‑Serial, Timestamp oder Time → `HH:MM` String."""
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
    """Extrahiert 0‑2 Fahrer‑Einträge aus einer Excel‑Zeile."""
    entries: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
    if pd.isna(datum):
        return entries

    kw, year = kw_year_sunday(datum)
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())

    base = {
        "KW": kw,
        "Jahr": year,
        "Datum": f"{weekday}, {datum.strftime('%d.%m.%Y')}",
        "Datum_raw": datum,
        "Wochentag": weekday,
        "Tour": row[15] if len(row) > 15 else "",
        "Uhrzeit": format_time(row[8]) if len(row) > 8 else "",
        "LKW": row[11] if len(row) > 11 else "",
    }

    # Fahrer 1
    if pd.notna(row[3]) and pd.notna(row[4]):
        entries.append({**base, "Name": f"{str(row[3]).strip()} {str(row[4]).strip()}"})
    # Fahrer 2
    if pd.notna(row[6]) and pd.notna(row[7]):
        entries.append({**base, "Name": f"{str(row[6]).strip()} {str(row[7]).strip()}"})

    return entries


def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.upper().strip())


def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """Liefert für jede PDF‑Seite den *erkannten* Namen (falls Treffer)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []
    normalized_excel = [normalize_name(n) for n in excel_names]

    for page_idx, page in enumerate(doc, start=1):
        text = page.get_text()
        found = ""
        for word in text.split():
            for orig, norm in zip(excel_names, normalized_excel):
                if normalize_name(word) in norm:
                    found = orig
                    break
            if found:
                break
        st.markdown(f"**Seite {page_idx} – Gefundener Name:** `{found or '❌ nicht erkannt'}`")
        results.append(found)
    doc.close()
    return results


def parse_excel_data(excel_file) -> pd.DataFrame:
    df = pd.read_excel(excel_file, header=None)
    entries: List[dict] = []
    for _, row in df.iterrows():
        entries.extend(extract_entries(row))
    return pd.DataFrame(entries)


def annotate_pdf_with_tours(pdf_bytes: bytes, ann: List[Optional[Dict[str, str]]]) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for pno, a in enumerate(ann):
        if pno >= len(doc) or not a:
            continue
        page = doc.load_page(pno)
        txt = " - ".join(filter(None, [a.get("tour"), a.get("weekday"), a.get("time")]))
        if not txt:
            continue
        rect = page.rect
        box = fitz.Rect(rect.width - 650, rect.height - 60, rect.width - 20, rect.height - 15)
        page.insert_textbox(box, txt, fontsize=12, fontname="helv", color=(1, 0, 0), align=fitz.TEXT_ALIGN_RIGHT)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def merge_annotated_pdfs(buffers: List[bytes]) -> bytes:
    if not buffers:
        return b""
    base = fitz.open(stream=buffers[0], filetype="pdf")
    for extra in buffers[1:]:
        tmp = fitz.open(stream=extra, filetype="pdf")
        base.insert_pdf(tmp)
        tmp.close()
    out = io.BytesIO()
    base.save(out)
    base.close()
    return out.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# 🔽 UI
# ──────────────────────────────────────────────────────────────────────────────

pdf_files = st.file_uploader("📑 PDFs hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan‑Excel hochladen", type=["xlsx", "xls", "xlsm"])

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF‑Dateien hochladen.")
    st.stop()

merged_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 PDFs analysieren & beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel‑Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel‑Daten einlesen …"):
        df_excel = parse_excel_data(excel_file)
        kw, jahr = kw_year_sunday(merged_date)
        filtered = df_excel[(df_excel["KW"] == kw) & (df_excel["Jahr"] == jahr)]

    if filtered.empty:
        st.warning(f"Keine Einträge für KW {kw} ({merged_date.strftime('%d.%m.%Y')}) im Excel gefunden!")
        st.stop()

    excel_names = filtered["Name"].unique().tolist()

    annotated_buffers: List[bytes] = []
    display_rows: List[dict] = []

    for pdf_file in pdf_files:
        st.subheader(f"📄 **{pdf_file.name}**")
        pdf_bytes = pdf_file.read()
        ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names)

        page_ann: List[Optional[dict]] = []
        for ocr in ocr_names:
            match_row = filtered[filtered["Name"] == ocr]
            if not match_row.empty:
                e = match_row.iloc[0]
                page_ann.append({
                    "matched_name": ocr,
                    "tour": str(e["Tour"]),
                    "weekday": str(e["Wochentag"]),
                    "time": str(e["Uhrzeit"]),
                })
            else:
                page_ann.append(None)

        # Tabelle Vorbereitung
        for i, (ocr, a) in enumerate(zip(ocr_names, page_ann), start=1):
            display_rows.append({
                "PDF": pdf_file.name,
                "Seite": i,
                "Gefundener Name": ocr or "❌",
                "Zugeordnet": a["matched_name"] if a else "❌ Nein",
                "Tour": a["tour"] if a else "",
                "Wochentag": a["weekday"] if a else "",
                "Uhrzeit": a["time"] if a else "",
            })

        annotated_buffers.append(annotate_pdf_with_tours(pdf_bytes, page_ann))

    st.dataframe(pd.DataFrame(display_rows), use_container_width=True)

    if any(annotated_buffers):
        st.success("✅ Alle PDFs beschriftet. Finale Datei wird erzeugt …")
        merged_pdf = merge_annotated_pdfs(annotated_buffers)
        st.download_button("📥 Zusammengeführte beschriftete PDF herunterladen", data=merged_pdf, file_name="dienstplaene_annotiert.pdf", mime="application/pdf")
    else:
        st.error("❌ Es konnten keine passenden Namen in den PDFs erkannt werden.")

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.8 – Mehrfach‑PDF‑Beschriftung*")
