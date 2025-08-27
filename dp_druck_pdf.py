from __future__ import annotations

import io
import re
import shutil
import unicodedata
from datetime import date, datetime, time
from typing import List, Dict, Optional

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
# Excel + Normalisierung
# ──────────────────────────────────────────────────────────────────────────────

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

def extract_entries(row: pd.Series) -> list[dict]:
    entries: list[dict] = []
    datum = pd.to_datetime(row[14], errors="coerce")  # Spalte O
    if pd.isna(datum):
        return entries

    weekday = WEEKDAYS_DE.get(datum.day_name(), datum.day_name())

    base = {
        "Datum": f"{weekday}, {datum.strftime('%d.%m.%Y')}",
        "Datum_raw": datum,
        "Wochentag": weekday,
        "Tour": row[15] if len(row) > 15 else "",
        "Uhrzeit": format_time(row[8]) if len(row) > 8 else "",
        "LKW": row[11] if len(row) > 11 else "",
    }

    if pd.notna(row[3]) and pd.notna(row[4]):
        entries.append({**base, "Name": f"{str(row[3]).strip()} {str(row[4]).strip()}"})
    if pd.notna(row[6]) and pd.notna(row[7]):
        entries.append({**base, "Name": f"{str(row[6]).strip()} {str(row[7]).strip()}"})
    return entries

def parse_excel_data(excel_file) -> pd.DataFrame:
    df = pd.read_excel(excel_file, header=None)
    entries: list[dict] = []
    for _, row in df.iterrows():
        entries.extend(extract_entries(row))
    return pd.DataFrame(entries)

def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.upper().strip())

def de_ascii_normalize(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    s = (s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
           .replace("Ä", "AE").replace("Ö", "OE").replace("Ü", "UE")
           .replace("ß", "ss").replace("ẞ", "SS"))
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Za-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.upper()

def prepend_filename_to_text(page_text: str, pdf_name: str) -> str:
    return f"__FILENAME__: {pdf_name}\n{page_text}"

# ──────────────────────────────────────────────────────────────────────────────
# Dateiname → Priorisierung + Adler-Regel
# ──────────────────────────────────────────────────────────────────────────────

def filename_tokens(pdf_name: str) -> list[str]:
    base = re.sub(r"\.[Pp][Dd][Ff]$", "", pdf_name)
    base = base.replace("ä","ae").replace("ö","oe").replace("ü","ue") \
               .replace("Ä","AE").replace("Ö","OE").replace("Ü","UE") \
               .replace("ß","ss").replace("ẞ","SS")
    base = re.sub(r"[^A-Za-z0-9]+", " ", base)
    return [t for t in base.upper().split() if t]

def choose_best_candidate(candidates: list[str], pdf_name: str) -> Optional[str]:
    """
    - Bevorzugt Nachname im Dateinamen.
    - Sonderregel: 'Adler' wird ignoriert, außer wenn 'ADLER' im Dateinamen vorkommt.
    """
    if not candidates:
        return None
    toks = set(filename_tokens(pdf_name))

    filtered = []
    for c in candidates:
        parts = [p for p in c.strip().split() if p]
        if parts and parts[0].upper() == "ADLER" and "ADLER" not in toks:
            continue
        filtered.append(c)

    if not filtered:
        return None

    for c in filtered:
        parts = [p for p in c.strip().split() if p]
        if len(parts) >= 2 and parts[0].upper() in toks:
            return c
    return filtered[0]

# ──────────────────────────────────────────────────────────────────────────────
# Matching-Methoden (Standard, Fuzzy, Robust)
# ──────────────────────────────────────────────────────────────────────────────

def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: list[str], pdf_name: str) -> list[str]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: list[str] = []

    excel_name_parts = []
    for name in excel_names:
        parts = name.strip().split()
        if len(parts) >= 2:
            excel_name_parts.append({'original': name, 'vorname': normalize_name(parts[1]), 'nachname': normalize_name(parts[0])})

    for page_idx, page in enumerate(doc, start=1):
        text = prepend_filename_to_text(page.get_text("text"), pdf_name)
        text_words = [normalize_name(w) for w in text.split()]

        candidates = []
        for info in excel_name_parts:
            if info['vorname'] in text_words and info['nachname'] in text_words:
                candidates.append(info['original'])

        chosen = choose_best_candidate(candidates, pdf_name) if candidates else ""
        results.append(chosen)
        st.markdown(f"**Seite {page_idx} – Gefundener Name:** {'✅ ' + chosen if chosen else '❌ nicht erkannt'}")

    doc.close()
    return results

def extract_names_from_pdf_fuzzy_match(pdf_bytes: bytes, excel_names: list[str], pdf_name: str) -> list[str]:
    try:
        from fuzzywuzzy import fuzz
    except ImportError:
        st.warning("FuzzyWuzzy fehlt → Standard-Matching.")
        return extract_names_from_pdf_by_word_match(pdf_bytes, excel_names, pdf_name)

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: list[str] = []

    excel_name_parts = []
    for name in excel_names:
        parts = name.strip().split()
        if len(parts) >= 2:
            excel_name_parts.append({'original': name, 'vorname': normalize_name(parts[1]), 'nachname': normalize_name(parts[0])})

    for page_idx, page in enumerate(doc, start=1):
        text = prepend_filename_to_text(page.get_text("text"), pdf_name)
        text_words = [normalize_name(w) for w in text.split()]

        candidates = []
        for info in excel_name_parts:
            vs = max((fuzz.ratio(info['vorname'], w) for w in text_words), default=0)
            ns = max((fuzz.ratio(info['nachname'], w) for w in text_words), default=0)
            if vs >= 90 and ns >= 90:
                candidates.append(info['original'])

        chosen = choose_best_candidate(candidates, pdf_name) if candidates else ""
        results.append(chosen)
        st.markdown(f"**Seite {page_idx} – Gefundener Name:** {'✅ ' + chosen if chosen else '❌ nicht erkannt'}")

    doc.close()
    return results

def extract_names_from_pdf_robust_text(pdf_bytes: bytes, excel_names: list[str], pdf_name: str) -> list[str]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: list[str] = []

    variants = []
    for full in excel_names:
        parts = full.strip().split()
        if len(parts) >= 2:
            nachname, vorname = parts[0], parts[1]
            v1 = de_ascii_normalize(f"{nachname} {vorname}")
            v2 = de_ascii_normalize(f"{vorname} {nachname}")
            variants.append((full, v1, v2, v1.replace(" ",""), v2.replace(" ","")))

    for page_idx, page in enumerate(doc, start=1):
        text_with_filename = prepend_filename_to_text(page.get_text("text"), pdf_name)
        norm = de_ascii_normalize(text_with_filename)
        norm_nosp = norm.replace(" ", "")

        candidates = []
        for original, v1, v2, v1n, v2n in variants:
            if v1 in norm or v2 in norm or v1n in norm_nosp or v2n in norm_nosp:
                candidates.append(original)

        chosen = choose_best_candidate(candidates, pdf_name) if candidates else ""
        results.append(chosen)
        st.markdown(f"**Seite {page_idx} – Gefundener Name:** {'✅ ' + chosen if chosen else '❌ nicht erkannt'}")

    doc.close()
    return results

# ──────────────────────────────────────────────────────────────────────────────
# PDF annotieren / mergen
# ──────────────────────────────────────────────────────────────────────────────

def annotate_pdf_with_tours(pdf_bytes: bytes, ann: list[Optional[Dict[str, str]]]) -> bytes:
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
        page.insert_textbox(box, txt, fontsize=12, fontname="helv", color=(1,0,0), align=fitz.TEXT_ALIGN_RIGHT)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()

def merge_annotated_pdfs(buffers: list[bytes]) -> bytes:
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
# UI
# ──────────────────────────────────────────────────────────────────────────────

pdf_files = st.file_uploader("📑 PDFs hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xls", "xlsm"])

matching_method = st.selectbox(
    "🔍 Matching-Methode wählen:",
    ["Standard (Exakter Match)", "Fuzzy-Matching (90% Ähnlichkeit)", "Robust (Text-Normalisierung)"]
)

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF-Dateien hochladen.")
    st.stop()

search_date: date = st.date_input("📅 Gesuchtes Datum (Spalte O)", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 PDFs analysieren & beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel-Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel-Daten einlesen …"):
        df_excel = parse_excel_data(excel_file)

    day_df = df_excel[df_excel["Datum_raw"].dt.date == search_date].copy()
    if day_df.empty:
        st.warning(f"Keine Einträge für {search_date.strftime('%d.%m.%Y')} gefunden!")
        st.stop()

    excel_names = day_df["Name"].unique().tolist()
    st.info(f"📋 Namen in Excel für {search_date.strftime('%d.%m.%Y')}: {', '.join(excel_names)}")

    annotated_buffers: list[bytes] = []
    display_rows: list[dict] = []

    for pdf_file in pdf_files:
        st.subheader(f"📄 **{pdf_file.name}**")
        pdf_bytes = pdf_file.read()

        if matching_method == "Fuzzy-Matching (90% Ähnlichkeit)":
            ocr_names = extract_names_from_pdf_fuzzy_match(pdf_bytes, excel_names, pdf_file.name)
        elif matching_method == "Robust (Text-Normalisierung)":
            ocr_names = extract_names_from_pdf_robust_text(pdf_bytes, excel_names, pdf_file.name)
        else:
            ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names, pdf_file.name)

        page_ann: list[Optional[dict]] = []
        for ocr in ocr_names:
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
                page_ann.append(None)

        for i, (ocr, a) in enumerate(zip(ocr_names, page_ann), start=1):
            display_rows.append({
                "PDF": pdf_file.name,
                "Seite": i,
                "Gefundener Name": ocr or "❌",
                "Zugeordnet": a["matched_name"] if a else "❌ Nein",
                "Tour": a["tour"] if a else "",
                "Wochentag": a["weekday"] if a else "",
                "Uhrzeit": a["time"] if a else "",
            })

        annotated_buffers.append(annotate_pdf_with_tours(pdf_bytes, page_ann))

    st.dataframe(pd.DataFrame(display_rows), use_container_width=True)

    if any(annotated_buffers):
        merged_pdf = merge_annotated_pdfs(annotated_buffers)
        st.download_button(
            "📥 Beschriftete PDF herunterladen",
            data=merged_pdf,
            file_name=f"dienstplaene_annotiert_{search_date.strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf"
        )
    else:
        st.error("❌ Keine passenden Namen erkannt.")
