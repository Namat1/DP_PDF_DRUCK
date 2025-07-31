from __future__ import annotations



import io
import re
import shutil
from datetime import date, datetime, timedelta
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

def extract_entries(row: pd.Series) -> List[dict]:
    """Liest bis zu **2 Fahrer** aus einer Excel‑Zeile (Spalten hart codiert)."""
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

# OCR‑Regex – zwei **aufeinander­folgende** Großbuchstaben‑Wörter → Vor‑ & Nachname
NAME_PATTERN = re.compile(r"([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)\s+([ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]+)")

def ocr_names_from_roi(pdf_bytes: bytes, roi: Tuple[int, int, int, int], dpi: int = 300) -> List[str]:
    """OCR für alle PDF‑Seiten im definierten ROI‑Bereich."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    names = []
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=dpi)
        pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # ROI ausschneiden
        roi_img = pil_img.crop(roi)
        
        # OCR mit deutschem Sprachmodell
        try:
            text = pytesseract.image_to_string(roi_img, lang="deu+eng")
        except:
            # Fallback ohne deutsche Sprache
            text = pytesseract.image_to_string(roi_img)
        
        # Namen extrahieren
        matches = NAME_PATTERN.findall(text)
        if matches:
            # Ersten gefundenen Namen nehmen
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
        # Anzahl übereinstimmender Wörter
        overlap = len(ocr_words & excel_words)
        if overlap > best_score:
            best_score = overlap
            best_match = excel_name
    
    return best_match if best_score > 0 else ""

def annotate_pdf_with_tours(pdf_bytes: bytes, names: List[str], tours: List[str]) -> bytes:
    """PDF mit Tour-Nummern annotieren."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        if page_num < len(tours) and tours[page_num]:
            # Tour-Nr. unten rechts einfügen
            rect = page.rect
            text_rect = fitz.Rect(rect.width - 450, rect.height - 41, rect.width - 50, rect.height - 1)
            
            page.insert_textbox(
                text_rect,
                f"Tour: {tours[page_num]}",
                fontsize=12,
                color=(1, 0, 0),  # Rot
                align=fitz.TEXT_ALIGN_RIGHT
            )
    
    # PDF in Bytes umwandeln
    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    doc.close()
    output_buffer.seek(0)
    return output_buffer.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# Datei‑Uploads
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
    doc = fitz.open(stream=pdf, filetype="pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=dpi)
    pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pil_img, pix.width, pix.height

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
    st.image(page1_img.crop(roi_box), caption="ROI‑Vorschau", use_column_width=True)

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
    
    with st.spinner("🔍 OCR läuft..."):
        # OCR für alle Seiten
        ocr_names = ocr_names_from_roi(pdf_bytes, roi_box)
        st.success(f"✅ OCR abgeschlossen: {len(ocr_names)} Seiten verarbeitet")
    
    with st.spinner("📊 Excel wird geparst..."):
        # Excel-Daten laden
        excel_data = parse_excel_data(excel_file)
        st.success(f"✅ Excel geparst: {len(excel_data)} Einträge gefunden")
    
    # Verteilungsdatum zu datetime konvertieren
    verteil_datetime = datetime.combine(verteil_date, datetime.min.time())
    
    # Nur Einträge für das Verteilungsdatum filtern
    filtered_data = excel_data[excel_data['Datum_raw'].dt.date == verteil_date]
    
    if filtered_data.empty:
        st.warning(f"⚠️ Keine Einträge für {verteil_date.strftime('%d.%m.%Y')} gefunden!")
        st.subheader("Verfügbare Daten in Excel:")
        if not excel_data.empty:
            available_dates = excel_data['Datum_raw'].dt.date.unique()
            st.write("Gefundene Datumsangaben:", sorted(available_dates))
    else:
        st.subheader("2️⃣ Matching-Ergebnisse")
        
        # Namen matching
        excel_names = filtered_data['Name'].unique().tolist()
        tours = []
        
        for i, ocr_name in enumerate(ocr_names):
            matched_name = fuzzy_match_name(ocr_name, excel_names)
            
            if matched_name:
                # Tour für gematchten Namen finden
                match_entry = filtered_data[filtered_data['Name'] == matched_name].iloc[0]
                tour = match_entry['Tour']
                tours.append(str(tour))
            else:
                tours.append("")
        
        # Ergebnisse anzeigen
        results_df = pd.DataFrame({
            'Seite': range(1, len(ocr_names) + 1),
            'OCR Name': ocr_names,
            'Matched Name': [fuzzy_match_name(name, excel_names) for name in ocr_names],
            'Tour': tours
        })
        
        st.dataframe(results_df, use_container_width=True)
        
        # Statistiken
        matched_count = sum(1 for tour in tours if tour)
        st.metric("Erfolgreich gematcht", f"{matched_count}/{len(tours)}")
        
        if matched_count > 0:
            with st.spinner("📝 PDF wird annotiert..."):
                # PDF annotieren
                annotated_pdf = annotate_pdf_with_tours(pdf_bytes, ocr_names, tours)
                
                # Download-Button
                st.download_button(
                    label="📥 Annotierte PDF herunterladen",
                    data=annotated_pdf,
                    file_name=f"dienstplan_annotiert_{verteil_date.strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    type="primary"
                )
                
                st.success("✅ PDF erfolgreich annotiert!")
        
        # Debug-Informationen
        with st.expander("🔍 Debug-Informationen"):
            st.write("**Gefilterte Excel-Daten:**")
            st.dataframe(filtered_data)
            
            st.write("**OCR-Namen (roh):**")
            for i, name in enumerate(ocr_names):
                st.write(f"Seite {i+1}: '{name}'")

# ──────────────────────────────────────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v1.0 – Automatische Tour-Zuordnung*")
