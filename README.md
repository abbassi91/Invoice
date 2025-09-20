# Invoice Extractor — Pro v6.3.1 (Prickle template, logic preserved)

This app keeps your exact logic stack:
- **pdf2image** (needs **Poppler**) for PDF → images
- **pdfplumber** for text & table extraction
- **pytesseract** (needs **Tesseract** binary) for OCR
- **rapidfuzz** for fuzzy search (optional)
- **streamlit-drawable-canvas** for rectangle templates
- **streamlit** UI

## Run locally
```bash
python -m venv .venv && source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Ubuntu/Debian:
sudo apt-get update && sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-eng

# macOS (Homebrew):
brew install poppler tesseract

streamlit run app_streamlit_cloud.py
```

## Deploy on Streamlit Community Cloud (step-by-step)
1. Push these files to the **root** of your GitHub repo:
   - `app_streamlit_cloud.py`
   - `requirements.txt`
   - `packages.txt`  ← installs Poppler & Tesseract at build time
2. In Streamlit Cloud → **New app**:
   - Select repo & branch
   - **Main file path**: `app_streamlit_cloud.py`
   - (Optional) Set Python to 3.12 or 3.13
3. Click **Deploy**. First build installs pip deps + system packages.
4. If the app can’t find Tesseract or Poppler, set **Secrets** (Settings → Secrets):
   ```
   TESSERACT_PATH="/usr/bin/tesseract"
   POPPLER_PATH="/usr/bin"
   ```
   (These paths are typical on the Streamlit Linux image.)
5. Reload the app, upload some PDFs, and go.

## Notes
- On Windows local dev, set the `Poppler bin path` in the sidebar to your Poppler `bin` folder.
- OCR languages default to `fra+eng+nld+ara` — adjust in the sidebar.
- Templates are stored in `templates_data/templates.json` at runtime.
