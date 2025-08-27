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

def kw_year_sunday(d: date | datetime) -> Tuple[int, int]:
    """Kalenderwoche & Jahr berechnen – Woche startet Sonntag."""
    if isinstance(d, date) and not isinstance(d, datetime):
        d = datetime.combine(d, time.min)
    s = d + timedelta(days=1)  # ISO -> Sonntag-Offset
    return int(s.strftime("%V")), int(s.strftime("%G"))

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
    """Extrahiert 0-2 Fahrer-Einträge aus einer Excel-Zeile."""
    entries: List[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
    if pd.isna(datum):
        return entries

    kw, year = kw_year_sunday(datum)
    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())

    base = {
        "KW": kw,
        "Jahr": year,
        "Datum": f"{weekday}, {datum.strftime('%d.%m.%Y')}",
        "Datum_raw": datum,  # pd.Timestamp
        "Wochentag": weekday,
        "Tour": row[15] if len(row) > 15 else "",
        "Uhrzeit": format_time(row[8]) if len(row) > 8 else "",
        "LKW": row[11] if len(row) > 11 else "",
    }

    # Fahrer 1
    if pd.notna(row[3]) and pd.notna(row[4]):
        entries.append({**base, "Name": f"{str(row[3]).strip()} {str(row[4]).strip()}"})
    # Fahrer 2
    if pd.notna(row[6]) and pd.notna(row[7]):
        entries.append({**base, "Name": f"{str(row[6]).strip()} {str(row[7]).strip()}"})

    return entries

def normalize_name(name: str) -> str:
    """Normalisiert Namen für besseren Vergleich."""
    return re.sub(r"\s+", " ", name.upper().strip())

def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """
    Liefert für jede PDF-Seite den erkannten Namen (falls Treffer).
    Vergleicht Vor- UND Nachnamen separat – nur bei exaktem Match beider Namen.
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
                st.markdown(f"**Seite {page_idx} – Gefundener Name:** ✅ {found_name} (exakt)")
                break

        if not found_name:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ❌ nicht erkannt")

        results.append(found_name)

    doc.close()
    return results

def extract_names_from_pdf_fuzzy_match(pdf_bytes: bytes, excel_names: List[str]) -> List[str]:
    """
    Erweiterte Version mit Fuzzy-String-Matching für robusteren Namensabgleich.
    Benötigt: pip install fuzzywuzzy python-levenshtein
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

def pick_closest_row_by_date(rows: pd.DataFrame, target_date: date) -> Optional[pd.Series]:
    """
    Wählt aus rows (gleicher Name, gleiche KW) die Zeile, deren Datum_raw dem target_date am nächsten ist.
    Verhindert den Bias, immer den Montag (erste Zeile) zu nehmen.
    """
    if rows.empty:
        return None
    tmp = rows.copy()
    # Differenz als absolute Tage
    tmp["_diff"] = (tmp["Datum_raw"].dt.normalize() - pd.Timestamp(target_date)).abs()
    tmp = tmp.sort_values(["_diff", "Datum_raw"])
    return tmp.iloc[0]

# ──────────────────────────────────────────────────────────────────────────────
# 🔽 UI
# ──────────────────────────────────────────────────────────────────────────────

pdf_files = st.file_uploader("📑 PDFs hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xls", "xlsm"])

# Option für Matching-Methode
matching_method = st.selectbox(
    "🔍 Matching-Methode wählen:",
    ["Standard (Exakter Match)", "Fuzzy-Matching (90% Ähnlichkeit)"],
    help="Standard: Nur bei exakter Übereinstimmung von Vor- und Nachname. Fuzzy: Erkennt auch kleine Abweichungen (90% Ähnlichkeit erforderlich)"
)

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF-Dateien hochladen.")
    st.stop()

merged_date: date = st.date_input("📅 Dienstpläne verteilen am:", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 PDFs analysieren & beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel-Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel-Daten einlesen …"):
        df_excel = parse_excel_data(excel_file)
        kw, jahr = kw_year_sunday(merged_date)
        filtered = df_excel[(df_excel["KW"] == kw) & (df_excel["Jahr"] == jahr)].copy()

    if filtered.empty:
        st.warning(f"Keine Einträge für KW {kw} ({merged_date.strftime('%d.%m.%Y')}) im Excel gefunden!")
        st.stop()

    # Die tatsächlich vorhandenen Tage der KW aus Excel (sortiert) → robuste Seiten→Datum-Map
    dates_in_week = sorted({ts.date() for ts in filtered["Datum_raw"]})
    if not dates_in_week:
        st.error("In der Excel sind keine Datumswerte für diese KW vorhanden.")
        st.stop()

    st.info("🗓️ Tage in Excel/KW: " + ", ".join(d.strftime("%a %d.%m.%Y") for d in dates_in_week))

    excel_names = filtered["Name"].unique().tolist()
    st.info(f"📋 Gefundene Namen in Excel für KW {kw}: {', '.join(excel_names)}")

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

        # Seiten→Datum: verwende die in Excel vorhandenen Tage der KW (Mo-Sa, So-Sa, etc.)
        # Bei mehr Seiten als Tage werden die letzten Tage wiederverwendet (selten)
        page_dates: List[date] = [
            dates_in_week[i] if i < len(dates_in_week) else dates_in_week[-1]
            for i in range(len(ocr_names))
        ]

        page_ann: List[Optional[dict]] = []
        for page_idx, (ocr, page_date) in enumerate(zip(ocr_names, page_dates), start=1):
            if not ocr:
                page_ann.append(None)
                continue

            # 1) Exakt: Name + Datum (Seitendatum)
            exact = filtered[(filtered["Name"] == ocr) & (filtered["Datum_raw"].dt.date == page_date)]

            if not exact.empty:
                e = exact.iloc[0]
            else:
                # 2) Nächstliegendes Datum für diesen Namen in der KW
                rows_same_name = filtered[filtered["Name"] == ocr]
                chosen = pick_closest_row_by_date(rows_same_name, page_date)
                if chosen is None:
                    page_ann.append(None)
                    continue
                e = chosen

            page_ann.append({
                "matched_name": ocr,
                "tour": str(e["Tour"]),
                "weekday": str(e["Wochentag"]),
                "time": str(e["Uhrzeit"]),
            })

        # Übersichtstabelle
        for i, (ocr, a, pdate) in enumerate(zip(ocr_names, page_ann, page_dates), start=1):
            display_rows.append({
                "PDF": pdf_file.name,
                "Seite": i,
                "Datum (Seite)": pdate.strftime("%d.%m.%Y"),
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
            file_name=f"dienstplaene_annotiert_KW{kw}_{jahr}.pdf", 
            mime="application/pdf"
        )
    else:
        st.error("❌ Es konnten keine passenden Namen in den PDFs erkannt werden.")

st.markdown("---")
st.markdown("*PDF Dienstplan Matcher v2.3 – robuste Seiten→Datum-Zuordnung & nächstliegendes Datum als Fallback*")
