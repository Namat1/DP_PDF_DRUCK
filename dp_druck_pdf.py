from __future__ import annotations

import io
import re
import shutil
import unicodedata
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
# Hilfsfunktionen – Excel & Normalisierung
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

def parse_excel_data(excel_file) -> pd.DataFrame:
    """Excel einlesen und alle Fahrer-Einträge extrahieren."""
    df = pd.read_excel(excel_file, header=None)
    entries: List[dict] = []
    for _, row in df.iterrows():
        entries.extend(extract_entries(row))
    return pd.DataFrame(entries)

def normalize_name(name: str) -> str:
    """Einfache Normalisierung für Wort-genaues Matching."""
    return re.sub(r"\s+", " ", name.upper().strip())

def de_ascii_normalize(s: str) -> str:
    """
    Robuste DE-Normalisierung:
    - ä/ö/ü → ae/oe/ue; ß → ss
    - Unicode NFKD (zerlegt Akzente/Ligaturen)
    - entfernt diakritische Zeichen
    - Nicht-Buchstaben/Ziffern → Leerzeichen
    - Mehrfach-Leerzeichen zu eins
    - Uppercase
    """
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
    """Fügt den originalen Dateinamen als erste Textzeile ein (im Rohtext sichtbar)."""
    return f"__FILENAME__: {pdf_name}\n{page_text}"

# ──────────────────────────────────────────────────────────────────────────────
# Hilfsfunktionen – Dateiname → Priorisierung
# ──────────────────────────────────────────────────────────────────────────────

def filename_tokens(pdf_name: str) -> List[str]:
    """Zerlegt den Dateinamen in Tokens (ähnlich normalisiert wie der Textvergleich)."""
    base = re.sub(r"\.[Pp][Dd][Ff]$", "", pdf_name)
    base = base.replace("ä","ae").replace("ö","oe").replace("ü","ue") \
               .replace("Ä","AE").replace("Ö","OE").replace("Ü","UE") \
               .replace("ß","ss").replace("ẞ","SS")
    base = re.sub(r"[^A-Za-z0-9]+", " ", base)
    return [t for t in base.upper().split() if t]

def choose_best_candidate(candidates: List[str], pdf_name: str) -> Optional[str]:
    """
    Bevorzugt den Kandidaten, dessen NACHNAME im Dateinamen vorkommt.
    (Für 'Nachname Vorname' wird das erste Wort als Nachname angenommen.)
    """
    if not candidates:
        return None
    toks = set(filename_tokens(pdf_name))
    for c in candidates:
        parts = [p for p in c.strip().split() if p]
        if len(parts) >= 2 and parts[0].upper() in toks:
            return c
    return candidates[0]  # deterministischer Fallback

# ──────────────────────────────────────────────────────────────────────────────
# Matching-Methoden
# ──────────────────────────────────────────────────────────────────────────────

def extract_names_from_pdf_by_word_match(pdf_bytes: bytes, excel_names: List[str], pdf_name: str) -> List[str]:
    """
    EXAKTES Wort-Matching:
    - Vor- und Nachname müssen als einzelne Wörter im Text stehen.
    - Mehrtreffer: sammelt alle, priorisiert via Dateiname.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []

    # Vorbereiten
    excel_name_parts = []
    for name in excel_names:
        parts = name.strip().split()
        if len(parts) >= 2:
            nachname = normalize_name(parts[0])
            vorname = normalize_name(parts[1])
            excel_name_parts.append({'original': name, 'vorname': vorname, 'nachname': nachname})

    for page_idx, page in enumerate(doc, start=1):
        raw = page.get_text("text")
        text = prepend_filename_to_text(raw, pdf_name)
        text_words = [normalize_name(word) for word in text.split()]

        candidates: List[str] = []
        for name_info in excel_name_parts:
            vorname_found = any(w == name_info['vorname'] for w in text_words)
            nachname_found = any(w == name_info['nachname'] for w in text_words)
            if vorname_found and nachname_found:
                candidates.append(name_info['original'])

        if candidates:
            chosen = choose_best_candidate(candidates, pdf_name)
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ✅ {chosen} (exakt, {len(candidates)} Treffer)")
            results.append(chosen or "")
        else:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ❌ nicht erkannt")
            results.append("")

    doc.close()
    return results

def extract_names_from_pdf_fuzzy_match(pdf_bytes: bytes, excel_names: List[str], pdf_name: str) -> List[str]:
    """
    FUZZY Matching (benötigt fuzzywuzzy + python-levenshtein):
    - Toleriert kleinere OCR-/Textextraktionsfehler.
    - Mehrtreffer: sammelt alle, priorisiert via Dateiname.
    """
    try:
        from fuzzywuzzy import fuzz
    except ImportError:
        st.warning("FuzzyWuzzy nicht installiert. Verwende Standard-Matching.")
        return extract_names_from_pdf_by_word_match(pdf_bytes, excel_names, pdf_name)

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
        raw = page.get_text("text")
        text = prepend_filename_to_text(raw, pdf_name)
        text_words = [normalize_name(word) for word in text.split()]

        candidates: List[str] = []
        best_scores: Dict[str, float] = {}

        for name_info in excel_name_parts:
            vs, ns = 0, 0
            for w in text_words:
                vs = max(vs, fuzz.ratio(name_info['vorname'], w))
                ns = max(ns, fuzz.ratio(name_info['nachname'], w))
            if vs >= 90 and ns >= 90:
                candidates.append(name_info['original'])
                best_scores[name_info['original']] = (vs + ns) / 2

        if candidates:
            chosen = choose_best_candidate(candidates, pdf_name)
            score_info = f" (≈{best_scores.get(chosen, 0):.0f}%)" if chosen in best_scores else ""
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ✅ {chosen}{score_info} (fuzzy, {len(candidates)} Treffer)")
            results.append(chosen or "")
        else:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ❌ nicht erkannt")
            results.append("")

    doc.close()
    return results

def extract_names_from_pdf_robust_text(pdf_bytes: bytes, excel_names: List[str], pdf_name: str) -> List[str]:
    """
    ROBUST (nur Text):
    - Ganze Seite als Fließtext + Dateiname als 1. Zeile
    - Starke DE-Normalisierung
    - Substring-Suche für 'NACHNAME VORNAME' & 'VORNAME NACHNAME'
    - akzeptiert Zeilenumbrüche, Ligaturen, Umlaute/ß, Sonderzeichen, zusammengeklebte Namen
    - Mehrtreffer → via Dateiname priorisiert
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: List[str] = []

    variants = []  # (original, v1_norm, v2_norm, v1_nosp, v2_nosp)
    for full in excel_names:
        parts = [p for p in str(full).strip().split() if p]
        if len(parts) >= 2:
            nachname, vorname = parts[0], parts[1]
            v1 = f"{nachname} {vorname}"
            v2 = f"{vorname} {nachname}"
            v1n = de_ascii_normalize(v1)
            v2n = de_ascii_normalize(v2)
            variants.append((full, v1n, v2n, v1n.replace(" ",""), v2n.replace(" ","")))

    for page_idx, page in enumerate(doc, start=1):
        raw = page.get_text("text")
        text_with_filename = prepend_filename_to_text(raw, pdf_name)
        norm = de_ascii_normalize(text_with_filename)
        norm_nosp = norm.replace(" ", "")

        found_candidates: List[str] = []
        for original, v1, v2, v1_nosp, v2_nosp in variants:
            if v1 in norm or v2 in norm or v1_nosp in norm_nosp or v2_nosp in norm_nosp:
                found_candidates.append(original)

        if found_candidates:
            chosen = choose_best_candidate(found_candidates, pdf_name)
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ✅ {chosen} (robust, {len(found_candidates)} Treffer, Dateiname priorisiert)")
            results.append(chosen or "")
        else:
            st.markdown(f"**Seite {page_idx} – Gefundener Name:** ❌ nicht erkannt")
            results.append("")

    doc.close()
    return results

# ──────────────────────────────────────────────────────────────────────────────
# PDF Beschriften / Mergen
# ──────────────────────────────────────────────────────────────────────────────

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

matching_method = st.selectbox(
    "🔍 Matching-Methode wählen:",
    ["Standard (Exakter Match)", "Fuzzy-Matching (90% Ähnlichkeit)", "Robust (Text-Normalisierung)"],
    help=(
        "Standard: Vor- und Nachname müssen als getrennte Wörter im PDF stehen. "
        "Fuzzy: toleriert kleinere OCR-/Textextraktionsfehler. "
        "Robust: sucht im normalisierten Fließtext (empfohlen bei Zeilenumbrüchen/Ligaturen/Umlauten/ß)."
    )
)

if not pdf_files:
    st.info("👉 Bitte zuerst eine oder mehrere PDF-Dateien hochladen.")
    st.stop()

# *** WICHTIG: Striktes Datumsmatching – nur dieses Datum wird verwendet ***
search_date: date = st.date_input("📅 Gesuchtes Datum (aus Excel Spalte O)", value=date.today(), format="DD.MM.YYYY")

if st.button("🚀 PDFs analysieren & beschriften", type="primary"):
    if not excel_file:
        st.error("⚠️ Bitte auch die Excel-Datei hochladen!")
        st.stop()

    with st.spinner("🔍 Excel-Daten einlesen …"):
        try:
            df_excel = parse_excel_data(excel_file)
        except Exception as e:
            st.exception(e)
            st.stop()

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
            ocr_names = extract_names_from_pdf_fuzzy_match(pdf_bytes, excel_names, pdf_file.name)
        elif matching_method == "Robust (Text-Normalisierung)":
            ocr_names = extract_names_from_pdf_robust_text(pdf_bytes, excel_names, pdf_file.name)
        else:
            ocr_names = extract_names_from_pdf_by_word_match(pdf_bytes, excel_names, pdf_file.name)

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
                "Dateiname (Text)": f"__FILENAME__: {pdf_file.name}",
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
st.markdown("*PDF Dienstplan Matcher – Striktes Datumsmatching (Spalte O) · Robust/ Fuzzy/ Exakt · Dateiname im Text & als Priorisierung*")
