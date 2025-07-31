"""
Streamlit Utility: ROI-Ermittlung für OCR in einer PDF
====================================================
**Ziel**: Nur den Koordinaten-Ausschnitt (ROI) bestimmen, in dem später OCR
laufen soll. Keine Excel-Verarbeitung, keine Annotation – reine
Koordinaten-Erfassung.

### Funktionsübersicht
1. PDF hochladen (mehrseitig).  
2. Erste Seite als Bild anzeigen.  
3. Benutzer zieht einen Rahmen → ROI.  
4. App zeigt die Koordinaten (Pixel im Original) an und bietet einen
   JSON-Download (`roi.json`).

### Deploy-Hinweise (Streamlit Cloud)
* **requirements.txt**
  ```
  streamlit
  streamlit-drawable-canvas
  pymupdf  # fitz
  pillow
  ```
* **packages.txt** (optional, nur falls später OCR nötig ist)
  ```
  tesseract-ocr
  ```

> **Kein OCR-Paket nötig** – wir ermitteln nur Koordinaten.
"""

from __future__ import annotations

# -----------------------------------------------------------------------------
# Monkey-Patch für `streamlit_drawable_canvas` ↔ Streamlit ≥1.32
# -----------------------------------------------------------------------------
import base64
import io as _io
from PIL import Image as _PIL_Image
import streamlit.elements.image as _st_img

if not hasattr(_st_img, "image_to_url"):
    def _image_to_url(img, *args, **kwargs):  # noqa: D401,E501
        """Lightweight Ersatz: Bild (PIL/numpy) → data-URL (PNG)."""
        if isinstance(img, _PIL_Image.Image):
            pil_img = img
        else:
            try:
                import numpy as np
                if isinstance(img, np.ndarray):
                    pil_img = _PIL_Image.fromarray(img)
                else:
                    raise TypeError
            except Exception as exc:
                raise TypeError("Unsupported image type for image_to_url") from exc
        buf = _io.BytesIO()
        pil_img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode()
        return f"data:image/png;base64,{b64}"

    _st_img.image_to_url = _image_to_url  # type: ignore[attr-defined]

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
import json
from typing import Tuple

import fitz  # PyMuPDF
import streamlit as st
from PIL import Image
from streamlit_drawable_canvas import st_canvas

# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------
st.set_page_config(page_title="OCR-ROI-Finder", layout="centered")

st.title("📐 OCR-ROI in PDF bestimmen")

with st.expander("Anleitung", expanded=False):
    st.markdown(
        """
        1. **PDF hochladen** – idealerweise die finale Vorlage.
        2. Rahmen ziehen, wo auf jeder Seite OCR ausgeführt werden soll.
        3. Koordinaten kopieren oder als JSON herunterladen.
        """
    )

pdf_file = st.file_uploader("PDF hochladen", type=["pdf"], key="pdf")

if pdf_file:
    pdf_bytes = pdf_file.read()

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:
        st.error(f"PDF konnte nicht gelesen werden: {exc}")
        st.stop()

    # Erste Seite als Vorschau (150 dpi)
    first_pix = doc[0].get_pixmap(dpi=150)
    first_img = Image.frombytes("RGB", [first_pix.width, first_pix.height], first_pix.samples)

    st.subheader("1️⃣ ROI zeichnen")
    CANVAS_WIDTH = 600
    ratio = CANVAS_WIDTH / first_img.width

    canvas_result = st_canvas(
        fill_color="",
        stroke_width=3,
        stroke_color="#FF0000",
        background_image=first_img.resize((CANVAS_WIDTH, int(first_img.height * ratio))),
        update_streamlit=True,
        height=int(first_img.height * ratio),
        width=CANVAS_WIDTH,
        drawing_mode="rect",
        key="canvas",
    )

    roi: Tuple[int, int, int, int] | None = None
    if canvas_result.json_data and canvas_result.json_data.get("objects"):
        rect = canvas_result.json_data["objects"][-1]
        l_disp, t_disp = rect["left"], rect["top"]
        w_disp, h_disp = rect["width"], rect["height"]

        left, top = int(l_disp / ratio), int(t_disp / ratio)
        right = int((l_disp + w_disp) / ratio)
        bottom = int((t_disp + h_disp) / ratio)
        roi = (left, top, right, bottom)

    if roi:
        st.success(f"ROI: x={roi[0]}:{roi[2]}, y={roi[1]}:{roi[3]} (Pixel im Original)")

        roi_json = json.dumps({"left": roi[0], "top": roi[1], "right": roi[2], "bottom": roi[3]}, indent=2)
        st.download_button("📥 ROI als JSON herunterladen", roi_json, file_name="roi.json", mime="application/json")

        with st.expander("Vorschau des zugeschnittenen Bereichs"):
            # Quick Crop aus der Vorschau (nur zum visuellen Check)
            crop_preview = first_img.crop(roi)
            st.image(crop_preview, caption="Zur Kontrolle – nur Seite 1")
    else:
        st.info("Bitte einen Rahmen zeichnen, um die Koordinaten anzuzeigen.")
