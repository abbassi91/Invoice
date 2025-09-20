# Prickle Template — Streamlit PDF Processor (original logic)

This project keeps your original logic stack:
- **pdf2image** (requires **Poppler**) for PDF → images
- **pdfplumber** for text/table extraction
- **pytesseract** (requires **Tesseract** binary) for OCR
- **rapidfuzz** for fuzzy search (optional)
- **streamlit** UI

## Run locally
```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# Ubuntu/Debian:
sudo apt-get update && sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-eng

# macOS (Homebrew):
brew install poppler tesseract

streamlit run app_streamlit_cloud.py
```

## Deploy to Streamlit Community Cloud
1. Push these files to the **root** of your GitHub repo:
   - `app_streamlit_cloud.py`
   - `requirements.txt`
   - `packages.txt`  ← installs Poppler & Tesseract
2. In Streamlit Cloud → **New app**:
   - Repo + branch
   - **Main file path**: `app_streamlit_cloud.py`
   - (Optional) Python: 3.12 or 3.13
3. Deploy. First build installs pip deps and system packages.
4. If needed, set secrets:
   - `TESSERACT_PATH="/usr/bin/tesseract"`
   - `POPLER_PATH="/usr/bin"`
5. Reload, upload a PDF, test.
