#!/usr/bin/env python3
"""
Prickle Template (unchanged logic style): Streamlit app for PDF → images + text + OCR
- Uses pdf2image (Poppler) for rasterizing PDFs
- Uses pdfplumber for text/table extraction
- Uses pytesseract for OCR (for image-only PDFs)
- Keeps the original dependency choices and flow
"""

from __future__ import annotations
import os
import io
from typing import List

import streamlit as st
import streamlit.components.v1 as components  # kept for compatibility with original template
import pdfplumber
from pdf2image import convert_from_path, convert_from_bytes
import pytesseract
from PIL import Image
import pandas as pd

# Optional fuzzy search
try:
    from rapidfuzz import fuzz as rf_fuzz
    HAVE_RAPIDFUZZ = True
except Exception:
    HAVE_RAPIDFUZZ = False

st.set_page_config(page_title="Prickle PDF App", page_icon="📄", layout="wide")

# ---------- Configuration ----------
POPLER_PATH = os.environ.get("POPLER_PATH")  # set if your poppler bin dir isn't on PATH
TESSERACT_PATH = os.environ.get("TESSERACT_PATH")  # path to tesseract binary

if TESSERACT_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH  # e.g., "/usr/bin/tesseract"

# ---------- Helpers ----------

def pdf_to_images_from_bytes(pdf_bytes: bytes, dpi: int = 200) -> List[Image.Image]:
    kwargs = dict(dpi=dpi)
    if POPLER_PATH:
        kwargs["poppler_path"] = POPLER_PATH
    try:
        return convert_from_bytes(pdf_bytes, **kwargs)
    except Exception:
        # Fallback: write to a temp file and call convert_from_path
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tmp:
            tmp.write(pdf_bytes)
            tmp.flush()
            return convert_from_path(tmp.name, **kwargs)

def extract_text_by_page(pdf_bytes: bytes):
    texts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
    return texts

def ocr_images(images: list[Image.Image], lang: str = "eng"):
    out = []
    for img in images:
        out.append(pytesseract.image_to_string(img, lang=lang) or "")
    return out

def extract_tables(pdf_bytes: bytes):
    tables = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables() or []
            for tbl in page_tables:
                tables.append(pd.DataFrame(tbl))
    return tables

# ---------- UI ----------

st.title("📄 Prickle Template — Streamlit PDF Processor")
st.caption("Keeps original logic: pdf2image (Poppler), pdfplumber, pytesseract, rapidfuzz.")

col1, col2 = st.columns([2,1], gap="large")
with col1:
    uploaded = st.file_uploader("Upload a PDF", type=["pdf"])
with col2:
    dpi = st.slider("Rasterize DPI (pdf2image)", min_value=72, max_value=300, value=200, step=8)
    want_ocr = st.checkbox("Run OCR with Tesseract", value=True)
    show_text = st.checkbox("Show extracted text", value=True)
    show_tables = st.checkbox("Try to extract tables", value=False)
    do_search = st.checkbox("Fuzzy search (rapidfuzz)", value=False and HAVE_RAPIDFUZZ)
    if do_search and not HAVE_RAPIDFUZZ:
        st.info("Install `rapidfuzz` to enable fuzzy search. (Already included in requirements here.)")

if uploaded is None:
    st.info("Upload a PDF to begin.")
    st.stop()

pdf_bytes = uploaded.read()

# Rasterize pages
with st.spinner("Rendering pages with pdf2image…"):
    images = pdf_to_images_from_bytes(pdf_bytes, dpi=dpi)

st.subheader("Page Previews")
for i, im in enumerate(images, 1):
    st.image(im, caption=f"Page {i}", use_column_width=True)

# Extract text
page_text = extract_text_by_page(pdf_bytes) if show_text else []

# OCR if requested
ocr_text = []
if want_ocr:
    with st.spinner("Running OCR (pytesseract)…"):
        ocr_text = ocr_images(images)

# Display text blocks
if show_text:
    st.subheader("Extracted Text (pdfplumber)")
    for i, txt in enumerate(page_text, 1):
        with st.expander(f"Page {i} — extracted text"):
            st.code(txt or "(no text)", language="text")

if want_ocr:
    st.subheader("OCR Text (Tesseract)")
    for i, txt in enumerate(ocr_text, 1):
        with st.expander(f"Page {i} — OCR text"):
            st.code(txt or "(no text)", language="text")

# Tables
if show_tables:
    with st.spinner("Extracting tables…"):
        tables = extract_tables(pdf_bytes)
    if not tables:
        st.warning("No tables detected.")
    else:
        st.success(f"Found {len(tables)} table(s).")
        for idx, df in enumerate(tables, 1):
            st.markdown(f"**Table {idx}**")
            st.dataframe(df)

# Fuzzy search
if do_search and HAVE_RAPIDFUZZ:
    query = st.text_input("Search term")
    if query:
        st.subheader("Search hits")
        hits = []
        sources = [("extract", page_text), ("ocr", ocr_text if want_ocr else [])]
        for label, texts in sources:
            for i, txt in enumerate(texts, 1):
                if not txt:
                    continue
                score = rf_fuzz.partial_ratio(query, txt)
                if score >= 50:
                    loc = txt.lower().find(query.lower())
                    if loc == -1:
                        loc = 0
                    snippet = txt[max(0, loc-60): loc+60].replace("\n", " ")
                    hits.append((label, i, score, snippet))
        if not hits:
            st.info("No close matches found.")
        else:
            hits = sorted(hits, key=lambda x: (-x[2], x[1]))
            for label, page_no, score, snippet in hits[:20]:
                st.write(f"**Page {page_no} ({label}) — score {int(score)}**")
                st.code(snippet, language="text")

st.caption("This build sticks to the original dependency choices (pdf2image/pytesseract).")
