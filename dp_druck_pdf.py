from __future__ import annotations

import io
import re
import shutil
from datetime import date, datetime, timedelta, time
from functools import lru_cache
from typing import List, Tuple, Dict, Optional

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
        "Tesseract‑Executable nicht gefunden. Bitte in **packages.txt** `tesseract-ocr` "
        "und optional `tesseract-ocr-deu` eintragen und App neu starten."
    )
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
    """KW‑Berechnung mit **Sonntag** als Wochen‑Start (ISO +1 Tag)."""
    s = d + timedelta(days=1)
    return int(s.strftime("%V")), int(s.strftime("%G"))

def format_time(value) -> str:
    """
    Konvertiert einen Excel-Zeitwert in einen 'HH:MM'-String.
    Behandelt datetime.time, datetime.datetime, float und String-Eingaben.
    """
    if pd.isna(value):
        return ""
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%H:%M")
    if isinstance(value, (int, float)):
        fractional_part = value % 1
        total_minutes = round(fractional_part * 1440)
        hours = total_minutes // 60
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"
    if isinstance(value, str):
        try:
            return pd.to_datetime(value).strftime("%H:%M")
        except (ValueError, TypeError):
            return value
    return str(value)

def extract_entries(row: pd.Series) -> List[dict]:
    """Liest bis zu **2 Fahrer** aus einer Excel‑Zeile (Spalten hart codiert)."""
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

    # Fahrer 1 (D,E)
    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        entry1 = base_entry.copy()
        entry1["Name"] = name
        entries.append(entry1)

    # Fahrer 2 (G,H)
    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        entry2 = base_entry.copy()
        entry2["Name"] = name
        entries.append(entry2)

    return entries

# OCR‑Regex – zwei **aufeinander­folgende** Großbuchstaben‑Wörter → Vor‑ & Nachname
NAME_PATTERN = re.compile(r"([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)\s+([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)")

def ocr_names_from_roi(pdf_bytes: bytes, roi: Tuple[int, int, int, int], dpi: int = 300) -> List[str]:
    """
    OCR für alle PDF‑Seiten im definierten ROI‑Bereich.
    Erwartet roi im Format (x, y, w, h) – also linke obere Ecke + Breite & Höhe.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    names = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=dpi)
        pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # ROI umrechnen: (x, y, w, h) → (left, upper, right, lower)
        x, y, w, h = roi
        left = x
        upper = y
        right = x + w
        lower = y + h

        # Sicherheitsprüfung
        W, H = pil_img.size
        if not (0 <= left < right <= W and 0 <= upper < lower <= H):
            raise ValueError(f"ROI {roi} ({left}, {upper}, {right}, {lower}) liegt außerhalb des Bildes ({W}×{H})")

        # Bildausschnitt extrahieren
        roi_img = pil_img.crop((left, upper, right, lower))

        # OCR ausführen
        try:
            text = pytesseract.image_to_string(roi_img, lang="deu+eng")
        except Exception:
            text = pytesseract.image_to_string(roi_img)

        # Namen per Regex extrahieren
        matches = NAME_PATTERN.findall(text)
        if matches:
            name = f"{matches[0][0]} {matches[0][1]}"
            names.append(name)
        else:
            names.append("")

    doc.close()
    return names


def parse_excel_data(excel_file) -> pd.DataFrame:
    """Excel-Datei parsen und Fahrer-Einträge extrahieren."""
    df = pd.read_excel(excel_file, header=None)
    all_entries = []
    for _, row in df.iterrows():
        entries = extract_entries(row)
        all_entries.extend(entries)
    return pd.DataFrame(all_entries)

def fuzzy_match_name(ocr_name: str, excel_names: List[str]) -> str:
    """Einfaches Fuzzy Matching für Namen."""
    if not ocr_name.strip():
        return ""
    ocr_words = set(ocr_name.upper().split())
    best_match = ""
    best_score = 0
    for excel_name in excel_names:
        excel_words = set(excel_name.upper().split())
        overlap = len(ocr_words & excel_words)
        if overlap > best_score:
            best_score = overlap
            best_match = excel_name
    return best_match if best_score > 0 else ""

def annotate_pdf_with_tours(pdf_bytes: bytes, annotations: List[Optional[Dict[str, str]]]) -> bytes:
    """PDF mit Tour-Informationen (Wochentag, Tour, Uhrzeit) annotieren."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for page_num, annotation in enumerate(annotations):
        if page_num < len(doc) and annotation:
            page = doc.load_page(page_num)
            
            tour = annotation.get("tour", "")
            weekday = annotation.get("weekday", "")
            uhrzeit = annotation.get("time", "")
            
            # Text für die Beschriftung im neuen Format zusammenbauen
            parts = []
            if tour:
                parts.append(tour)
            if weekday:
                parts.append(weekday)
            if uhrzeit:
                parts.append(f"{uhrzeit} Uhr")
            
            text_to_insert = " - ".join(parts)
            
            # Tour-Nr. unten rechts einfügen
            rect = page.rect
            text_rect = fitz.Rect(rect.width - 650, rect.height - 60, rect.width - 20, rect.height - 15)
            
            page.insert_textbox(
                text_rect,
                text_to_insert,
                fontsize=12,
                fontname="hebo",
                color=(1, 0, 0),  # Rot
                align=fitz.TEXT_ALIGN_RIGHT
            )
    
    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    doc.close()
    output_buffer.seek(0)
    return output_buffer.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# Datei‑Uploads
# ──────────────────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("📑 PDF hochladen", type=["pdf"], key="pdf")
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xlsm"], key="excel")

if not pdf_file:
    st.info("👉 Bitte zuerst ein PDF hochladen.")
    st.stop()

pdf_bytes = pdf_file.read()

# ──────────────────────────────────────────────────────────────────────────────
# ROI-Koordinaten festlegen
# ──────────────────────────────────────────────────────────────────────────────
roi_box = (59, 264, 137, 53)



    


# ──────────────────────────────────────────────────────────────────────────────
# Verteilungs‑Datum (vom Nutzer bestimmen lassen)
# ──────────────────────────────────────────────────────────────────────────────
verteil_date: date = st.date_input(
    "📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY"
)

# ──────────────────────────────────────────────────────────────────────────────
# Haupt‑Button – OCR, Excel, Match & Annotate
# ──────────────────────────────────────────────────────────────────────────────
if st.button("🚀 OCR & PDF beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel‑Datei hochladen!")
        st.stop()
    
    with st.spinner("🔍 OCR läuft und Excel wird verarbeitet..."):
        ocr_names = ocr_names_from_roi(pdf_bytes, roi_box)
        excel_data = parse_excel_data(excel_file)
        filtered_data = excel_data[excel_data['Datum_raw'].dt.date == verteil_date]
    
    if filtered_data.empty:
        st.warning(f"⚠️ Keine Einträge für {verteil_date.strftime('%d.%m.%Y')} in der Excel-Datei gefunden!")
    else:
        excel_names = filtered_data['Name'].unique().tolist()
        page_annotations = []
        
        for ocr_name in ocr_names:
            matched_name = fuzzy_match_name(ocr_name, excel_names)
            if matched_name:
                match_entry = filtered_data[filtered_data['Name'] == matched_name].iloc[0]
                annotation_info = {
                    "matched_name": matched_name,
                    "tour": str(match_entry['Tour']),
                    "weekday": str(match_entry['Wochentag']),
                    "time": str(match_entry['Uhrzeit'])
                }
                page_annotations.append(annotation_info)
            else:
                page_annotations.append(None)
        
        # --- NEU: Diagnose-Tabelle erstellen und anzeigen ---
        st.markdown("---")
        st.subheader("🔍 Ergebnis der Zuordnung")
        
        display_data = []
        for i, (ocr_name, annotation) in enumerate(zip(ocr_names, page_annotations)):
            row_data = {
                "PDF Seite": i + 1,
                "Gefundener Name (OCR)": ocr_name or "N/A",
                "Zugeordnet (Excel)": annotation.get("matched_name", "❌ Nein") if annotation else "❌ Nein",
                "Tour": annotation.get("tour", "") if annotation else "",
                "Wochentag": annotation.get("weekday", "") if annotation else "",
                "Uhrzeit": annotation.get("time", "") if annotation else ""
            }
            display_data.append(row_data)
        
        st.dataframe(pd.DataFrame(display_data), use_container_width=True)
        st.info("Bitte überprüfen Sie die Zuordnung. Nur für Seiten mit einem zugeordneten Namen wird eine Beschriftung erzeugt.")
        st.markdown("---")
        
        matched_count = sum(1 for anno in page_annotations if anno)
        
        if matched_count > 0:
            with st.spinner("📝 PDF wird beschriftet..."):
                annotated_pdf = annotate_pdf_with_tours(pdf_bytes, page_annotations)
            
            st.download_button(
                label="📥 Beschriftete PDF herunterladen",
                data=annotated_pdf,
                file_name=f"dienstplan_annotiert_{verteil_date.strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                type="primary"
            )
        else:
            st.error("Es konnten keine Übereinstimmungen zwischen PDF und Excel-Liste gefunden werden.")


# ──────────────────────────────────────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.6 – Format der Beschriftung angepasst*")
