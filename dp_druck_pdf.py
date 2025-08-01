from __future__ import annotations

import io
import re
import shutil
from datetime import date, datetime, timedelta, time
from typing import List, Tuple, Dict, Optional

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st

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
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")
st.title("📄 Dienstpläne beschriften & verteilen – Mehrfach‑PDF‑Modus")

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
    """Kalenderwoche nach *Sonntag–Samstag‑Logik* ermitteln."""
    s = d + timedelta(days=1)  # einen Tag vorziehen – Sonntag‑Start
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
    """Eine Zeile aus dem Excel‑Plan in 1–2 Fahrereinträge aufbrechen."""
    entries: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
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

    # Fahrer 1 (Spalten D & E)
    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        entry1 = base_entry.copy()
        entry1["Name"] = name
        entries.append(entry1)

    # Fahrer 2 (Spalten G & H)
    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        entry2 = base_entry.copy()
        entry2["Name"] = name
        entries.append(entry2)

    return entries


def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.upper().strip())


def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """Einfacher Wortabtast‑Match – schnell & ausreichend robust."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []
    normalized_excel = [normalize_name(n) for n in excel_names]

    for i, page in enumerate(doc):
        text = page.get_text()
        found = ""
        for word in text.split():
            for orig, norm in zip(excel_names, normalized_excel):
                if normalize_name(word) in norm:
                    found = orig
                    break
            if found:
                break
        st.markdown(f"**Seite {i+1} – Gefundener Name:** `{found}`")
        results.append(found)

    doc.close()
    return results


def parse_excel_data(excel_file) -> pd.DataFrame:
    df = pd.read_excel(excel_file, header=None)
    all_entries: List[dict] = []
    for _, row in df.iterrows():
        all_entries.extend(extract_entries(row))
    return pd.DataFrame(all_entries)


def annotate_pdf_with_tours(pdf_bytes: bytes, annotations: List[Optional[Dict[str, str]]]) -> bytes:
    """PDF‑Seiten mit Tour‑Infos beschriften."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page_num, annotation in enumerate(annotations):
        if page_num < len(doc) and annotation:
            page = doc.load_page(page_num)
            text = " - ".join(
                filter(None, [annotation.get("tour"), annotation.get("weekday"), annotation.get("time") + " Uhr"])
            )
            rect = page.rect
            text_rect = fitz.Rect(rect.width - 650, rect.height - 60, rect.width - 20, rect.height - 15)
            page.insert_textbox(
                text_rect,
                text,
                fontsize=12,
                fontname="helv",
                color=(1, 0, 0),
                align=fitz.TEXT_ALIGN_RIGHT,
            )
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# 📑 Datei‑Uploads
# ──────────────────────────────────────────────────────────────────────────────

pdf_files = st.file_uploader("📑 PDF‑Dienstpläne hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan‑Excel hochladen", type=["xlsx", "xlsm"])

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF‑Dateien hochladen.")
    st.stop()

verteil_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY")

# ──────────────────────────────────────────────────────────────────────────────
# 🚀 Hauptlogik nach Button‑Klick
# ──────────────────────────────────────────────────────────────────────────────

if st.button("🚀 PDFs analysieren, beschriften & zusammenführen", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel‑Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel‑Daten laden & Namen extrahieren..."):
        excel_data = parse_excel_data(excel_file)
        kw, jahr = kw_year_sunday(verteil_date)
        filtered_data = excel_data[(excel_data["KW"] == kw) & (excel_data["Jahr"] == jahr)]

    if filtered_data.empty:
        st.warning(f"⚠️ Keine Einträge für KW {kw} ({verteil_date.strftime('%d.%m.%Y')}) in der Excel‑Datei gefunden!")
        st.stop()

    excel_names = filtered_data["Name"].unique().tolist()

    display_rows: List[Dict[str, str]] = []
    final_doc = fitz.open()  # Leere PDF für Zusammenführung

    for pdf_idx, pdf_file in enumerate(pdf_files, start=1):
        st.markdown(f"### 📄 PDF {pdf_idx}: {pdf_file.name}")
        pdf_bytes = pdf_file.read()

        # 1️⃣ OCR / Wortvergleich
        ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names)

        # 2️⃣ Excel‑Zuordnung & Annotationen vorbereiten
        page_annotations: List[Optional[Dict[str, str]]] = []
        for ocr_name in ocr_names:
            matched = filtered_data[filtered_data["Name"] == ocr_name]
            if not matched.empty:
                entry = matched.iloc[0]
                page_annotations.append(
                    {
                        "matched_name": ocr_name,
                        "tour": str(entry["Tour"]),
                        "weekday": str(entry["Wochentag"]),
                        "time": str(entry["Uhrzeit"]),
                    }
                )
            else:
                page_annotations.append(None)

        # 3️⃣ Anzeige zur Kontrolle vorbereiten
        for i, (ocr_name, ann) in enumerate(zip(ocr_names, page_annotations), start=1):
            display_rows.append(
                {
                    "PDF": pdf_file.name,
                    "Seite": i,
                    "Gefundener Name": ocr_name or "❌ Nicht erkannt",
                    "Zugeordnet": ann["matched_name"] if ann else "❌ Nein",
                    "Tour": ann["tour"] if ann else "",
                    "Wochentag": ann["weekday"] if ann else "",
                    "Uhrzeit": ann["time"] if ann else "",
                }
            )

        # 4️⃣ PDF beschriften
        annotated_bytes = annotate_pdf_with_tours(pdf_bytes, page_annotations)

        # 5️⃣ Beschriftete Seiten an Gesamtdokument anhängen
        with fitz.open(stream=annotated_bytes, filetype="pdf") as annotated_doc:
            final_doc.insert_pdf(annotated_doc)

    # 6️⃣ Ergebnis anzeigen
    st.dataframe(pd.DataFrame(display_rows), use_container_width=True)

    # 7️⃣ Gesamte PDF speichern & Download anbieten
    out_buf = io.BytesIO()
    final_doc.save(out_buf)
    final_doc.close()

    st.success("✅ Alle PDFs wurden beschriftet und zusammengeführt.")
    st.download_button(
        "📥 Gesamte beschriftete PDF herunterladen",
        data=out_buf.getvalue(),
        file_name="dienstplaene_annotiert_gesamt.pdf",
        mime="application/pdf",
    )

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.8 – Mehrfach‑PDF‑Support*  © 2025")
