"""
Streamlit App: OCR‑gestützte PDF‑Annotation nach Namenssuche
-----------------------------------------------------------
• Lädt eine mehrseitige PDF‑Datei und eine Excel‑Tabelle hoch
• Liest jede PDF‑Seite per OCR (Tesseract) ein
• Sucht nach Namen aus einer gewählten Spalte der Excel‑Tabelle
• Wird ein Name gefunden, wird der Wert aus einer anderen gewählten Spalte
  (z. B. Abteilung, Personal‑ID o. Ä.) auf genau dieser PDF‑Seite ausgegeben
• Liefert eine kommentierte PDF zum Download

Benötigte Python‑Pakete
----------------------
streamlit, pandas, pdf2image, pillow, pytesseract, pymupdf (fitz), openpyxl

Wichtiger Hinweis
----------------
Tesseract‑OCR muss lokal installiert sein und der ausführbare Pfad ggf. in der
Variable pytesseract.pytesseract.tesseract_cmd gesetzt werden.
"""

import io
import re
from pathlib import Path

import streamlit as st
import pandas as pd
import pytesseract
import fitz  # PyMuPDF
from pdf2image import convert_from_bytes
from PIL import Image

# -----------------------------------------------------------------------------
# Konfiguration
# -----------------------------------------------------------------------------
st.set_page_config(page_title="PDF‑Namenssuche & Annotation", layout="centered")

st.title("🔎 PDF‑Namenssuche mit Excel‑Referenz und Annotation")

with st.expander("Anleitung", expanded=False):
    st.markdown(
        """
        1. **PDF hochladen**: Mehrseitige PDF, vorzugsweise gescannt oder als
           Bild‑PDF.
        2. **Excel‑Datei hochladen**: Enthält die Namen (z. B. Spalte
           `Name`) sowie die dazugehörige Information, die in die PDF
           geschrieben werden soll (z. B. Spalte `Abteilung`).
        3. **Spalten wählen**: Geben Sie an, welche Spalte den Namen enthält
           und welche Spalte die einzutragende Information.
        4. **Start**: Der Vorgang kann je nach Seitenzahl etwas dauern. Anschließend
           erhalten Sie die annotierte PDF zum Download.
        """
    )

# -----------------------------------------------------------------------------
# Datei‑Uploads
# -----------------------------------------------------------------------------
pdf_file = st.file_uploader("PDF hochladen", type=["pdf"], key="pdf")
excel_file = st.file_uploader("Excel‑Datei hochladen", type=["xlsx", "xls"], key="excel")

if pdf_file and excel_file:
    # Excel einlesen
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Excel konnte nicht eingelesen werden: {e}")
        st.stop()

    if df.empty:
        st.warning("Die Excel‑Datei enthält keine Daten.")
        st.stop()

    # Spaltenauswahl
    with st.sidebar:
        st.header("Spaltenauswahl")
        name_col = st.selectbox("Spalte mit Namen", df.columns)
        value_col = st.selectbox("Spalte mit auszugebender Information", df.columns, index=min(1, len(df.columns)-1))
        font_size = st.number_input("Schriftgröße (Pt)", 6, 48, value=12)
        y_position = st.number_input("Y‑Position in Pt (von oben)", 0, 800, value=50)
        x_position = st.number_input("X‑Position in Pt (von links)", 0, 600, value=50)
        case_sensitive = st.checkbox("Groß‑/Kleinschreibung beachten", value=False)

    # Name‑und‑Wert‑Dictionary vorbereiten
    name_map = {
        (str(row[name_col]) if case_sensitive else str(row[name_col]).lower()): row[value_col]
        for _, row in df.iterrows()
        if pd.notna(row[name_col])
    }

    if st.button("🚀 Starten"):
        with st.spinner("Verarbeite PDF… bitte warten"):
            pdf_bytes = pdf_file.read()

            # OCR: PDF in Bilder konvertieren
            try:
                images = convert_from_bytes(pdf_bytes)
            except Exception as e:
                st.error(f"PDF‑Konvertierung fehlgeschlagen: {e}")
                st.stop()

            # Ursprüngliche PDF erneut als bearbeitbares Dokument laden
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")

            # Seitenweise Verarbeitung
            for page_index, (pil_img, page) in enumerate(zip(images, doc), start=1):
                # OCR
                ocr_text = pytesseract.image_to_string(pil_img, lang="deu")
                search_space = ocr_text if case_sensitive else ocr_text.lower()

                # Suche aller Namen auf dieser Seite
                for search_name, value in name_map.items():
                    pattern = rf"\b{re.escape(search_name)}\b"
                    if re.search(pattern, search_space):
                        # Wert auf die PDF‑Seite schreiben
                        insertion_text = str(value)
                        # (x, y) Koordinate siehe Sidebar‑Input
                        page.insert_text(
                            (x_position, y_position),
                            insertion_text,
                            fontsize=font_size,
                            fontname="helv",
                            fill=(0, 0, 0),
                        )
                        # Sobald erster Treffer gesetzt, nächste Seite
                        break

            # Annotierte PDF speichern
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            doc.close()
            output_buffer.seek(0)

        st.success("Fertig! Die PDF ist annotiert.")
        st.download_button(
            label="📥 Annotierte PDF herunterladen",
            data=output_buffer,
            file_name="annotiert.pdf",
            mime="application/pdf",
        )

# -----------------------------------------------------------------------------
# Fußnote
# -----------------------------------------------------------------------------
st.markdown(
    """
    *Erstellt mit ❤️ und [PyMuPDF](https://pymupdf.readthedocs.io/) · Bitte
    stellen Sie sicher, dass [Tesseract‑OCR](https://github.com/tesseract-ocr)
    installiert und erreichbar ist.*
    """
)

