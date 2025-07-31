from __future__ import annotations

"""
Streamlit Utility – Interaktiver PDF‑ROI‑Finder & Excel‑Extraktor
===============================================================
### Funktionen
1. **ROI visual bestimmen**: PDF‑Upload → Seite 1 wird mit 300 DPI gerendert. Über vier Zahl‑Inputs legst du die Koordinaten fest.
2. **Live‑OCR & Vorschau**: Ausgeschnittener Bereich + sofortiges OCR‑Ergebnis.
3. **OCR auf alle Seiten**: Wenn das Rechteck passt, wird dasselbe ROI für jede Seite verwendet; Texte & erkannte (groß geschriebene) Wörter werden geloggt.
4. **Excel‑Einträge auslesen**: Gleichzeitig kannst du eine Excel‑Datei hochladen. Die Funktion `extract_entries_both_sides` liest pro Zeile bis zu zwei Fahrernamen samt Datum/KW/Tour/LKW aus und erzeugt eine tabellarische Übersicht.

*Hinweis*: Für einen ersten Test reicht die mitgelieferte Beispiel‑Excel.

### requirements.txt
```
streamlit
pymupdf  # fitz
pytesseract
pandas
pillow
openpyxl
```

### packages.txt (nur für Streamlit Cloud)
```
poppler-utils
tesseract-ocr
tesseract-ocr-deu
```
"""

import io
import locale
import re
import shutil
from datetime import datetime, timedelta
from functools import lru_cache
from typing import List, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image, ImageDraw

# ──────────────────────────────────────────────────────────────────────────────
# Tesseract – Pfad fixieren (wichtig für Streamlit Cloud)
# ──────────────────────────────────────────────────────────────────────────────
TESSERACT_CMD = shutil.which("tesseract")
if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
else:
    st.error(
        "Tesseract‑Executable nicht gefunden. Bitte in **packages.txt** `tesseract-ocr` "
        "und optional `tesseract-ocr-deu` eintragen und App neu starten."
    )
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Seiteneinstellungen & UI‑Titel
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF‑ROI & Excel‑Extraktor", layout="wide")
st.title("📄 PDF‑ROI bestimmen & Excel‑Einträge extrahieren")

with st.expander("Kurzanleitung", expanded=False):
    st.markdown(
        """
        1. **PDF hochladen** und ROI setzen.
        2. **Excel hochladen** (optional) – wird automatisch eingelesen.
        3. Prüfe Vorschau‑OCR und Excel‑Tabelle.
        4. Wenn alles passt ➜ *OCR auf alle Seiten*.
        """
    )

# ──────────────────────────────────────────────────────────────────────────────
# Excel‑Hilfsfunktionen
# ──────────────────────────────────────────────────────────────────────────────

# Deutschsprachige Wochentags‑Mapping
wochentage_deutsch = {
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag",
}

# "KW"‑Berechnung: Kalenderwoche mit **Sonntag** als erstem Tag

def get_kw_and_year_sunday_start(datum: datetime) -> Tuple[int, int]:
    # Python ISO KW (Montag‑Start) → wir verschieben um einen Tag
    sonntag_basiert = datum + timedelta(days=1)
    kw = int(sonntag_basiert.strftime("%V"))  # ISO KW aus Datum +1 Tag
    jahr = int(sonntag_basiert.strftime("%G"))
    return kw, jahr


def extract_entries_both_sides(row: pd.Series) -> List[dict]:
    """Fasst pro Excel‑Zeile bis zu zwei Fahrer‑Einträge zusammen.

    Erwartete Spalten (0‑Index):
    - O (14): Datum
    - D/E (3,4): Fahrer 1 Vor‑ & Nachname
    - G/H (6,7): Fahrer 2 Vor‑ & Nachname
    - L (11): LKW
    - P (15): Tour
    - Q (16): Uhrzeit (optional)
    """
    eintraege: List[dict] = []

    # Datum parsen/validieren
    datum = pd.to_datetime(row[14], errors="coerce")
    if pd.isna(datum):
        return eintraege

    kw, jahr = get_kw_and_year_sunday_start(datum)
    wochentag_en = datum.day_name()
    wochentag_de = wochentage_deutsch.get(wochentag_en, wochentag_en)
    datum_formatiert = datum.strftime("%d.%m.%Y")
    datum_komplett = f"{wochentag_de}, {datum_formatiert}"

    tour = row[15] if len(row) > 15 else ""
    uhrzeit = row[16] if len(row) > 16 else ""
    lkw = row[11] if len(row) > 11 else ""

    # Fahrer 1
    if pd.notna(row[3]) and pd.notna(row[4]):
        name = f"{str(row[3]).strip()} {str(row[4]).strip()}"
        eintraege.append(
            {
                "KW": kw,
                "Jahr": jahr,
                "Datum": datum_komplett,
                "Datum_sortierbar": datum,
                "Name": name,
                "Tour": tour,
                "Uhrzeit": uhrzeit,
                "LKW": lkw,
            }
        )

    # Fahrer 2
    if pd.notna(row[6]) and pd.notna(row[7]):
        name = f"{str(row[6]).strip()} {str(row[7]).strip()}"
        eintraege.append(
            {
                "KW": kw,
                "Jahr": jahr,
                "Datum": datum_komplett,
                "Datum_sortierbar": datum,
                "Name": name,
                "Tour": tour,
                "Uhrzeit": uhrzeit,
                "LKW": lkw,
            }
        )

    return eintraege

# ──────────────────────────────────────────────────────────────────────────────
# PDF‑Upload + ROI
# ──────────────────────────────────────────────────────────────────────────────
pdf_file = st.file_uploader("📑 PDF hochladen", type=["pdf"], key="pdf")
excel_file = st.file_uploader("📊 Excel‑Datei hochladen", type=["xlsx", "xlsm"], key="excel")

if pdf_file:
    pdf_bytes = pdf_file.read()

    # Render erste Seite → Cache, damit Koordinatenänderung flott ist
    @lru_cache(maxsize=2)
    def render_first_page(pdf: bytes, dpi: int = 300):
        d = fitz.open(stream=pdf, filetype="pdf")
        p = d.load_page(0)
        pix = p.get_pixmap(dpi=dpi)
        pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return pil, pix.width, pix.height

    img, width, height = render_first_page(pdf_bytes)

    st.subheader("1️⃣ ROI wählen (Koordinaten in Pixel)")
    col_left, col_right = st.columns([1, 2])

    with col_left:
        st.write("**Bildgröße**:", f"{width} × {height} px")
        x1 = st.number_input("x1 (links)", 0, width - 1, value=st.session_state.get("x1", 200))
        y1 = st.number_input("y1 (oben)", 0, height - 1, value=st.session_state.get("y1", 890))
        x2 = st.number_input("x2 (rechts)", x1 + 1, width, value=st.session_state.get("x2", 560))
        y2 = st.number_input("y2 (unten)", y1 + 1, height, value=st.session_state.get("y2", 980))
        st.session_state.update({"x1": x1, "y1": y1, "x2": x2, "y2": y2})

    with col_right:
        roi = (x1, y1, x2, y2)
        overlay = img.copy()
        ImageDraw.Draw(overlay).rectangle(roi, outline="red", width=6)
        st.image(overlay, caption="Seite 1 mit markiertem ROI", use_column_width=True)
        crop = img.crop(roi)
        st.image(crop, caption="ROI‑Vorschau", use_column_width=True)
        ocr_preview = pytesseract.image_to_string(crop, lang="deu").strip()
        st.text_area("OCR‑Ergebnis (Seite 1)", ocr_preview, height=120)

    # ──────────────────────────────────────────────────────────────────────────
    # Excel‑Einlesen (falls vorhanden)
    # ──────────────────────────────────────────────────────────────────────────
    if excel_file:
        try:
            df_xl = pd.read_excel(excel_file, engine="openpyxl", header=None)  # ohne Header
        except Exception as exc:
            st.error(f"Excel konnte nicht gelesen werden: {exc}")
            df_xl = pd.DataFrame()

        if not df_xl.empty:
            st.subheader("2️⃣ Excel‑Vorschau (erste 15 Zeilen)")
            st.dataframe(df_xl.head(15), use_container_width=True)

            # Einträge extrahieren
            eintraege: list[dict] = []
            for _, r in df_xl.iterrows():
                eintraege.extend(extract_entries_both_sides(r))

            if eintraege:
                df_entries = pd.DataFrame(eintraege).sort_values("Datum_sortierbar")
                st.subheader("3️⃣ Extrahierte Einträge")
                st.dataframe(df_entries.drop(columns=["Datum_sortierbar"]), use_container_width=True)

                # CSV‑Download
                csv_buf = io.StringIO()
                df_entries.to_csv(csv_buf, index=False)
                st.download_button("📥 Einträge als CSV", csv_buf.getvalue(), "excel_eintraege.csv", "text/csv")
            else:
                st.info("Keine gültigen Fahrer‑Einträge gefunden.")

    # ──────────────────────────────────────────────────────────────────────────
    # Button: OCR auf alle Seiten
    # ──────────────────────────────────────────────────────────────────────────
    if st.button("🚀 OCR auf *alle* PDF‑Seiten", type="primary"):
        with st.spinner("Starte OCR …"):
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            roi_tuple = (x1, y1, x2, y2)
            data: list[tuple[int, str]] = []
            name_candidates: set[str] = set()

            for page_idx, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=300)
                page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                txt = pytesseract.image_to_string(page_img.crop(roi_tuple), lang="deu").strip()
                data.append((page_idx, txt))
                name_candidates.update(re.findall(r"\b[ÄÖÜA-Z][ÄÖÜA-Za-zäöüß-]{1,}\b", txt))

            df_pdf = pd.DataFrame(data, columns=["Seite", "Text (ROI)"])

        st.success("OCR abgeschlossen ✔️")
        st.dataframe(df_pdf, use_container_width=True)
        csv = df_pdf.to_csv(index=False)
        st.download_button("📥 PDF‑OCR‑Tabelle als CSV", csv, "pdf_roi_ocr.csv", "text/csv")

        st.subheader("Potentielle Namen aus dem PDF")
        if name_candidates:
            st.write("; ".join(sorted(name_candidates)))
        else:
            st.info("Keine großgeschriebenen Wörter gefunden.")

else:
    st.info("👉 Bitte zuerst ein PDF hochladen (und optional Excel).")
