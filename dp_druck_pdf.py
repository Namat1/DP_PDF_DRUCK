from __future__ import annotations

import io
import re
import shutil
import unicodedata
from datetime import date, datetime, time
from typing import List, Dict, Optional
from PIL import Image

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
# Excel + Normalisierung (bestehende Funktionen bleiben gleich)
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
# Dateiname → Priorisierung + Adler-Regel (bestehende Funktionen bleiben gleich)
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
# VERBESSERTE NAMENERKENNUNG mit OCR + Mehreren Strategien
# ──────────────────────────────────────────────────────────────────────────────

def extract_text_with_ocr(page, use_ocr: bool = True) -> str:
    """Extrahiert Text aus PDF-Seite mit optionalem OCR für Bilder"""
    # Zuerst normalen Text versuchen
    text = page.get_text("text")
    
    if not use_ocr or len(text.strip()) > 50:  # Wenn genug Text vorhanden, kein OCR nötig
        return text
    
    try:
        # OCR für die gesamte Seite als Fallback
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x Vergrößerung für bessere OCR
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # OCR mit deutschen Optionen
        ocr_text = pytesseract.image_to_string(
            img, 
            lang='deu+eng',
            config='--psm 6 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzäöüßÄÖÜ0123456789 .-'
        )
        
        # Kombiniere beide Texte
        combined_text = f"{text}\n---OCR---\n{ocr_text}"
        return combined_text
        
    except Exception as e:
        st.warning(f"OCR fehlgeschlagen: {str(e)}")
        return text

def create_name_variants(full_name: str) -> dict:
    """Erstellt alle möglichen Varianten eines Namens für besseres Matching"""
    parts = full_name.strip().split()
    if len(parts) < 2:
        return {}
    
    nachname, vorname = parts[0], parts[1]
    
    variants = {
        'original': full_name,
        'nachname': nachname,
        'vorname': vorname,
        'full_normalized': normalize_name(full_name),
        'nachname_normalized': normalize_name(nachname),
        'vorname_normalized': normalize_name(vorname),
        'ascii_full': de_ascii_normalize(full_name),
        'ascii_nachname': de_ascii_normalize(nachname),
        'ascii_vorname': de_ascii_normalize(vorname),
        'reversed': f"{vorname} {nachname}",  # Falls Vor- und Nachname vertauscht sind
        'initials': f"{nachname[0]}.{vorname[0]}." if nachname and vorname else "",
        'short_variants': []
    }
    
    # Kurze Varianten (z.B. "Max Mustermann" -> "Max M.", "M. Mustermann")
    if len(nachname) > 0 and len(vorname) > 0:
        variants['short_variants'].extend([
            f"{vorname} {nachname[0]}.",
            f"{vorname[0]}. {nachname}",
            f"{nachname}, {vorname}",  # Komma-Format
            f"{nachname},{vorname}",   # Ohne Leerzeichen
        ])
    
    return variants

def fuzzy_match_with_threshold(text: str, name_variants: dict, threshold: int = 85) -> int:
    """Fuzzy-Matching mit konfigurierbarem Threshold"""
    try:
        from fuzzywuzzy import fuzz
        
        max_score = 0
        text_upper = text.upper()
        
        # Teste alle Varianten
        for key, value in name_variants.items():
            if key == 'short_variants':
                for variant in value:
                    score = fuzz.partial_ratio(variant.upper(), text_upper)
                    max_score = max(max_score, score)
            elif isinstance(value, str) and value:
                score = fuzz.partial_ratio(value.upper(), text_upper)
                max_score = max(max_score, score)
        
        return max_score
    except ImportError:
        return 0

def advanced_name_matching(text: str, excel_names: list[str], pdf_name: str, debug: bool = False) -> list[str]:
    """Erweiterte Namenerkennung mit mehreren Strategien"""
    
    candidates = []
    debug_info = []
    
    # Erstelle Varianten für alle Namen
    name_variants_list = []
    for name in excel_names:
        variants = create_name_variants(name)
        name_variants_list.append(variants)
    
    text_normalized = normalize_name(text)
    text_ascii = de_ascii_normalize(text)
    text_words = text_normalized.split()
    text_lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    for variants in name_variants_list:
        original_name = variants['original']
        match_score = 0
        match_method = ""
        
        # Strategie 1: Exakte Wort-Matches
        if variants['nachname_normalized'] in text_words and variants['vorname_normalized'] in text_words:
            match_score = 100
            match_method = "Exakte Wörter"
        
        # Strategie 2: ASCII-normalisierter Text-Match
        elif variants['ascii_nachname'] in text_ascii and variants['ascii_vorname'] in text_ascii:
            match_score = 95
            match_method = "ASCII-normalisiert"
        
        # Strategie 3: Zeilen-basierte Suche (Namen stehen oft in einer Zeile)
        elif any(variants['ascii_nachname'] in de_ascii_normalize(line) and 
                variants['ascii_vorname'] in de_ascii_normalize(line) for line in text_lines):
            match_score = 90
            match_method = "Zeilen-Match"
        
        # Strategie 4: Fuzzy-Matching
        else:
            fuzzy_score = fuzzy_match_with_threshold(text, variants, threshold=80)
            if fuzzy_score >= 80:
                match_score = fuzzy_score
                match_method = f"Fuzzy ({fuzzy_score}%)"
        
        # Strategie 5: Kurze Varianten (Initialen, etc.)
        if match_score == 0:
            for short_variant in variants['short_variants']:
                if de_ascii_normalize(short_variant) in text_ascii:
                    match_score = 70
                    match_method = f"Kurzvariante: {short_variant}"
                    break
        
        if match_score > 0:
            candidates.append((original_name, match_score, match_method))
            debug_info.append(f"✅ {original_name}: {match_score}% ({match_method})")
        else:
            debug_info.append(f"❌ {original_name}: Nicht gefunden")
    
    if debug:
        for info in debug_info:
            st.write(info)
    
    # Sortiere nach Score und wähle besten Kandidaten
    candidates.sort(key=lambda x: x[1], reverse=True)
    
    if candidates:
        best_candidates = [c[0] for c in candidates if c[1] >= candidates[0][1] - 10]  # Ähnliche Scores
        return [choose_best_candidate(best_candidates, pdf_name) or candidates[0][0]]
    
    return []

def extract_names_enhanced(pdf_bytes: bytes, excel_names: list[str], pdf_name: str, use_ocr: bool = True) -> list[str]:
    """Verbesserte Hauptfunktion für Namenerkennung"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results: list[str] = []
    
    for page_idx, page in enumerate(doc, start=1):
        st.markdown(f"### Seite {page_idx}")
        
        # Text mit optionalem OCR extrahieren
        text = extract_text_with_ocr(page, use_ocr)
        text_with_filename = prepend_filename_to_text(text, pdf_name)
        
        # Zeige extrahierten Text (erste 500 Zeichen)
        with st.expander(f"📄 Extrahierter Text (Seite {page_idx})"):
            st.text(text_with_filename[:500] + "..." if len(text_with_filename) > 500 else text_with_filename)
        
        # Erweiterte Namenerkennung
        with st.expander(f"🔍 Matching-Details (Seite {page_idx})"):
            matched_names = advanced_name_matching(text_with_filename, excel_names, pdf_name, debug=True)
        
        chosen = matched_names[0] if matched_names else ""
        results.append(chosen)
        
        if chosen:
            st.success(f"**✅ Gefunden:** {chosen}")
        else:
            st.error("**❌ Kein Name erkannt**")
    
    doc.close()
    return results

# ──────────────────────────────────────────────────────────────────────────────
# PDF annotieren / mergen (bestehende Funktionen bleiben gleich)
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
# VERBESSERTE UI
# ──────────────────────────────────────────────────────────────────────────────

pdf_files = st.file_uploader("📑 PDFs hochladen", type=["pdf"], accept_multiple_files=True)
excel_file = st.file_uploader("📊 Tourplan-Excel hochladen", type=["xlsx", "xls", "xlsm"])

col1, col2 = st.columns(2)
with col1:
    use_ocr = st.checkbox("🔍 OCR für gescannte PDFs verwenden", value=True, 
                         help="Aktiviert OCR für bessere Texterkennung bei gescannten PDFs")
with col2:
    show_debug = st.checkbox("🐛 Debug-Informationen anzeigen", value=False,
                           help="Zeigt detaillierte Matching-Informationen")

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
        st.header(f"📄 **{pdf_file.name}**")
        pdf_bytes = pdf_file.read()

        # Verwende die verbesserte Namenerkennung
        ocr_names = extract_names_enhanced(pdf_bytes, excel_names, pdf_file.name, use_ocr)

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

    st.subheader("📊 Zusammenfassung")
    df_results = pd.DataFrame(display_rows)
    st.dataframe(df_results, use_container_width=True)
    
    # Erfolgsstatistik
    total_pages = len(df_results)
    successful_matches = len(df_results[df_results["Gefundener Name"] != "❌"])
    success_rate = (successful_matches / total_pages * 100) if total_pages > 0 else 0
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Gesamt Seiten", total_pages)
    with col2:
        st.metric("Erkannte Namen", successful_matches)
    with col3:
        st.metric("Erfolgsrate", f"{success_rate:.1f}%")

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
        
    # Verbesserungsvorschläge
    if success_rate < 80:
        st.warning("⚠️ **Niedrige Erkennungsrate - Verbesserungsvorschläge:**")
        st.write("• OCR aktivieren falls noch nicht geschehen")
        st.write("• PDF-Qualität prüfen (Auflösung, Kontrast)")
        st.write("• Namen in Excel-Datei auf Tippfehler prüfen")
        st.write("• Debug-Modus aktivieren für detaillierte Analyse")
