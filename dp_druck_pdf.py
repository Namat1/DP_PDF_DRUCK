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
    st.error("Tesseract-Executable nicht gefunden. Bitte in **packages.txt** tesseract-ocr eintragen.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PDF Dienstplan Matcher", layout="wide")
st.title("📄 Dienstpläne beschriften & verteilen (Multi-PDF)")

# ──────────────────────────────────────────────────────────────────────────────
WEEKDAYS_DE: Dict[str, str] = {
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag",
}

# ──────────────────────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────────────────────────────────────

def format_time(value) -> str:
    """Zahl, Excel-Serial, Timestamp oder Time → HH:MM String."""
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
    """Extrahiert 0-2 Fahrer-Einträge aus einer Excel-Zeile (Datum = Spalte O / Index 14)."""
    entries: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
    if pd.isna(datum):
        return entries

    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())

    base = {
        "Datum": f"{weekday}, {datum.strftime('%d.%m.%Y')}",
        "Datum_raw": datum,  # pd.Timestamp
        "Wochentag": weekday,
        "Tour": row[15] if len(row) > 15 else "",
        "Uhrzeit": format_time(row[8]) if len(row) > 8 else "",
        "LKW": row[11] if len(row) > 11 else "",
    }

    # Fahrer 1 (D + E)
    if pd.notna(row[3]) and pd.notna(row[4]):
        entries.append({**base, "Name": f"{str(row[3]).strip()} {str(row[4]).strip()}"})
    # Fahrer 2 (G + H)
    if pd.notna(row[6]) and pd.notna(row[7]):
        entries.append({**base, "Name": f"{str(row[6]).strip()} {str(row[7]).strip()}"})

    return entries

def normalize_name(name: str) -> str:
    """Normalisiert Namen für besseren Vergleich."""
    return re.sub(r"\s+", " ", name.upper().strip())

def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """
    Ermittelt je Seite einen Namen durch EXAKTES Wort-Matching (Vor- und Nachname müssen getrennt auftauchen).
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []

    # Excel-Namen in Vor- und Nachnamen aufteilen
    excel_name_parts = []
    for name in excel_names:
        parts = name.strip().split()
        if len(parts) >= 2:
            nachname = normalize_name(parts[0])
            vorname = normalize_name(parts[1])
            excel_name_parts.append({'original': name, 'vorname': vorname, 'nachname': nachname})

    for page_idx, page in enumerate(doc, start=1):
        text = page.get_text()
        text_words = [normalize_name(word) for word in text.split()]
        found_name = ""

        for name_info in excel_name_parts:
            vorname_found = any(word == name_info['vorname'] for word in text_words)
            nachname_found = any(word == name_info['nachname'] for word in text_words)
            if vorname_found and nachname_found:
                found_name = name_info['original']
                st.markdown(f"**Seite {page_idx} – Gefundener Name:** ✅ {found_name}")
                break

        if not found_name:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ❌ nicht erkannt")

        results.append(found_name)

    doc.close()
    return results

def extract_names_from_pdf_fuzzy_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """
    Fuzzy-Matching für robusteren Namensabgleich (benötigt fuzzywuzzy + python-levenshtein).
    """
    try:
        from fuzzywuzzy import fuzz
    except ImportError:
        st.warning("FuzzyWuzzy nicht installiert. Verwende Standard-Matching.")
        return extract_names_from_pdf_by_word_match(pdf_bytes, excel_names)

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []

    excel_name_parts = []
    for name in excel_names:
        parts = name.strip().split()
        if len(parts) >= 2:
            nachname = normalize_name(parts[0])
            vorname = normalize_name(parts[1])
            excel_name_parts.append({'original': name, 'vorname': vorname, 'nachname': nachname})

    for page_idx, page in enumerate(doc, start=1):
        text = page.get_text()
        text_words = [normalize_name(word) for word in text.split()]
        found_name = ""
        best_score = 0

        for name_info in excel_name_parts:
            vs, ns = 0, 0
            for w in text_words:
                vs = max(vs, fuzz.ratio(name_info['vorname'], w))
                ns = max(ns, fuzz.ratio(name_info['nachname'], w))
            if vs >= 90 and ns >= 90:
                score = (vs + ns) / 2
                if score > best_score:
                    best_score = score
                    found_name = name_info['original']

        if found_name:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ✅ {found_name} (≈{best_score:.0f}%)")
        else:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ❌ nicht erkannt")

        results.append(found_name)

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
        txt = " - ".join(filter(None, [a.get("tour"), a.get("weekday"), a.get("time")]))
        if not txt:
            continue
        rect = page.rect
        box = fitz.Rect(rect.width - 650, rect.height - 80, rect.width - 20, rect.height - 25)
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

# ──────────────────────────────────────────────────────────────────────────────
# 🔽 UI
# ──────────────────────────────────────────────────────────────────────────────

pdf_files = st.file_uploader("📑 PDFs hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xls", "xlsm"])

# Matching-Methode
matching_method = st.selectbox(
    "🔍 Matching-Methode wählen:",
    ["Standard (Exakter Match)", "Fuzzy-Matching (90% Ähnlichkeit)"],
    help="Standard: Nur bei exakter Übereinstimmung von Vor- und Nachname. Fuzzy: erkennt kleine Abweichungen (90%)."
)

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF-Dateien hochladen.")
    st.stop()

# *** WICHTIG: Genau dieses Datum wird gesucht (nicht KW-weit) ***
search_date: date = st.date_input("📅 Gesuchtes Datum (aus Excel Spalte O)", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 PDFs analysieren & beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel-Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel-Daten einlesen …"):
        df_excel = parse_excel_data(excel_file)

    # Nur dieses Datum verwenden (kein KW-Fallback!)
    day_df = df_excel[df_excel["Datum_raw"].dt.date == search_date].copy()

    if day_df.empty:
        st.warning(f"Keine Einträge für das Datum {search_date.strftime('%d.%m.%Y')} in der Excel gefunden!")
        st.stop()

    excel_names = day_df["Name"].unique().tolist()
    st.info(f"📋 Gefundene Namen für {search_date.strftime('%d.%m.%Y')}: {', '.join(excel_names)}")

    annotated_buffers: List[bytes] = []
    display_rows: List[dict] = []

    for pdf_file in pdf_files:
        st.subheader(f"📄 **{pdf_file.name}**")
        pdf_bytes = pdf_file.read()

        # OCR-Namen je Seite (gewählte Methode)
        if matching_method == "Fuzzy-Matching (90% Ähnlichkeit)":
            ocr_names = extract_names_from_pdf_fuzzy_match(pdf_bytes, excel_names)
        else:
            ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names)

        # Für jede Seite strikt: Name + exakt dieses Datum
        page_ann: List[Optional[dict]] = []
        for page_idx, ocr in enumerate(ocr_names, start=1):
            if not ocr:
                page_ann.append(None)
                continue

            match_row = day_df[day_df["Name"] == ocr]
            if not match_row.empty:
                e = match_row.iloc[0]
                page_ann.append({
                    "matched_name": ocr,
                    "tour": str(e["Tour"]),
                    "weekday": str(e["Wochentag"]),
                    "time": str(e["Uhrzeit"]),
                })
            else:
                # Kein Eintrag für diesen Namen **am gewählten Datum** → unzugeordnet lassen
                page_ann.append(None)

        # Übersichtstabelle
        for i, (ocr, a) in enumerate(zip(ocr_names, page_ann), start=1):
            display_rows.append({
                "PDF": pdf_file.name,
                "Seite": i,
                "Datum (fix)": search_date.strftime("%d.%m.%Y"),
                "Gefundener Name": ocr or "❌",
                "Zugeordnet": a["matched_name"] if a else "❌ Nein",
                "Tour": a["tour"] if a else "",
                "Wochentag": a["weekday"] if a else "",
                "Uhrzeit": a["time"] if a else "",
            })

        annotated_buffers.append(annotate_pdf_with_tours(pdf_bytes, page_ann))

    st.subheader("📊 Übersicht aller Zuordnungen")
    st.dataframe(pd.DataFrame(display_rows), use_container_width=True)

    if any(annotated_buffers):
        st.success("✅ Alle PDFs beschriftet. Finale Datei wird erzeugt …")
        merged_pdf = merge_annotated_pdfs(annotated_buffers)
        st.download_button(
            "📥 Zusammengeführte beschriftete PDF herunterladen",
            data=merged_pdf,
            file_name=f"dienstplaene_annotiert_{search_date.strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf"
        )
    else:
        st.error("❌ Es konnten keine passenden Namen am gewählten Datum erkannt werden.")

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher – Striktes Datumsmatching (Spalte O)*")
