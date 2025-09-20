#!/usr/bin/env python3
# === App: Invoice Extractor — Pro v6.3.1 (Prickle template, logic preserved) ===
import os, io, json, tempfile, re, shutil, base64, time, glob
from pathlib import Path
from PIL import Image
import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
from pdf2image import convert_from_path
import pytesseract
import pandas as pd

st.set_page_config(page_title="Invoice Extractor — Pro v6.3.1", layout="wide")
st.markdown(
    """
    <style>
    .stat-card {border-radius:16px;padding:14px 16px;border:1px solid #e5e7eb;background:#fafafa}
    .stat-title {font-size:12px;color:#6b7280;margin:0}
    .stat-value {font-size:22px;font-weight:700;margin:0}
    </style>
    """, unsafe_allow_html=True
)
st.title("Invoice Extractor — Pro v6.3.1  •  Filter Fix • Folder Picker • Combined Excel")

ON_CLOUD = bool(os.environ.get("STREAMLIT_SERVER_ENABLED") or os.environ.get("STREAMLIT_RUNTIME"))
TESSERACT_AVAILABLE = shutil.which("tesseract") is not None
POPPLER_HINT = "Windows: install Poppler (bin) • macOS: brew install poppler • OCR: tesseract"

with st.sidebar:
    st.header("General")
    poppler_path = st.text_input("Poppler bin path (Windows)", value="")
    use_ocr = st.checkbox("Enable OCR (Tesseract)", value=False)
    tess_langs = st.text_input("OCR languages", value="fra+eng+nld+ara")
    if ON_CLOUD and not TESSERACT_AVAILABLE:
        st.info("Cloud: Tesseract unavailable → OCR off.")
        use_ocr = False

    st.header("Performance (thousands of PDFs)")
    fast_filter = st.checkbox("Fast filter (text only, no OCR)", value=True)
    skip_items = st.checkbox("Skip item-lines extraction (Smart)", value=True)
    batch_size = st.slider("Batch export size", 200, 5000, 1000, step=100)
    preview_dpi = st.slider("Preview DPI (templates)", 120, 220, 150, step=10)
    ocr_zone_dpi = st.slider("OCR DPI for zones (templates)", 120, 220, 160, step=10)
    st.caption(POPPLER_HINT)

DATA_DIR = Path(".").resolve() / "templates_data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_FILE = DATA_DIR / "templates.json"
if not TEMPLATES_FILE.exists():
    TEMPLATES_FILE.write_text(json.dumps([], ensure_ascii=False, indent=2), encoding="utf-8")
templates = json.loads(TEMPLATES_FILE.read_text(encoding="utf-8"))

tabs = st.tabs(["Smart Extract (Pro)", "Templates (full-page)", "PDF Preview", "Notes"])

HEADER_WORDS = set(["description","descriptions","articles","produits","product","products","items","détails","details","ligne","lignes"])

LABEL_SYNONYMS = {
    "Invoice_No": ["facture", "facture n", "facture no", "n° facture", "invoice", "invoice no", "inv no", "numéro facture", "ref", "reference", "référence", "document", "numero", "nr"],
    "Date": ["date", "date facture", "date d'émission", "issued", "datum"],
    "Customer": ["client", "billed to", "bill to", "customer", "klant", "à:", "adressé à"],
    "Vendor": ["fournisseur", "from", "vendor", "supplier", "leverancier", "société"],
    "Total_TTC": ["total ttc", "ttc", "montant ttc", "total à payer", "amount due", "total due", "grand total", "total", "totaal"],
    "Subtotal": ["sous-total", "subtotal", "total ht", "hors taxe", "ht"],
    "Tax": ["tva", "tax", "vat", "tva 20", "tva 10", "btw"],
    "PO_Number": ["bon de commande", "purchase order", "po", "po no", "commande n"],
    "IBAN": ["iban"],
    "TVA_Number": ["n° tva", "tva intracom", "vat number", "numéro de tva", "btw nummer"],
    "Address": ["adresse", "address", "adres"],
}

def pdf_page_to_image(path, page_num=0, dpi=200):
    try:
        return convert_from_path(str(path), dpi=dpi, poppler_path=poppler_path or None, first_page=page_num+1, last_page=page_num+1)[0]
    except Exception:
        return None

def ocr_on_image(img: Image.Image, langs="eng"):
    try:
        return pytesseract.image_to_string(img, lang=langs) or ""
    except Exception:
        return ""

def cleanup_multiline(text, join_separator=", ", drop_first_like_header=True) -> str:
    if text is None:
        return ""
    try:
        import math
        if isinstance(text, float) and math.isnan(text):
            return ""
    except Exception:
        pass
    s = str(text)
    if s.lower() == "nan":
        return ""
    lines = [ln.strip() for ln in s.splitlines() if ln and ln.strip()]
    if not lines:
        return ""
    if drop_first_like_header and len(lines) >= 2:
        first = re.sub(r"[^A-Za-zÀ-ÿ]", " ", lines[0]).strip().lower()
        token = first.split(" ")[0] if first else ""
        if token in HEADER_WORDS or first in HEADER_WORDS:
            lines = lines[1:]
    return join_separator.join(lines)

def normalize_amount(s):
    if not s: return None
    t = str(s).strip()
    t = re.sub(r"[€$£]|USD|CAD|EUR|GBP", "", t, flags=re.I)
    t = t.replace(" ", "").replace("\xa0","")
    if re.search(r"\d+\.\d{3},\d{2}$", t): t = t.replace(".", "").replace(",", ".")
    elif re.search(r"\d+,\d{3}\.\d{2}$", t): t = t.replace(",", "")
    t = t.replace(",", ".")
    m = re.search(r"(-?\d+(?:\.\d{1,2})?)", t)
    try: return float(m.group(1)) if m else None
    except Exception: return None

@st.cache_data(show_spinner=False)
def extract_page_texts_cached(pdf_path_str: str):
    pages = []
    try:
        with pdfplumber.open(pdf_path_str) as pdf:
            for p in pdf.pages:
                pages.append(p.extract_text() or "")
    except Exception:
        pass
    return pages

@st.cache_data(show_spinner=False)
def pdf_contains_text_cached(pdf_path_str: str, query: str):
    if not query: return True
    q = query.strip().lower()
    try:
        with pdfplumber.open(pdf_path_str) as pdf:
            for p in pdf.pages:
                t = (p.extract_text() or "").lower()
                if q in t: return True
    except Exception: pass
    return False

def smart_find_value(lines, keylist):
    key_regex = r"|".join([re.escape(k) for k in keylist])
    pattern_inline = re.compile(rf"(?i)\b({key_regex})\b\s*[:#-]?\s*(.+?)\s*$")
    for i, L in enumerate(lines):
        txt = " ".join(L.strip().split())
        m = pattern_inline.search(txt)
        if m:
            val = m.group(2).strip()
            if val and len(val) > 1:
                return val
        if re.search(rf"(?i)\b({key_regex})\b", txt) and i+1 < len(lines):
            nxt = " ".join(lines[i+1].strip().split())
            if nxt and len(nxt) > 1:
                return nxt
    return ""

def detect_tables_pdfplumber(pdf_path_str):
    tables = []
    try:
        with pdfplumber.open(pdf_path_str) as pdf:
            for page in pdf.pages:
                t1 = page.extract_table()
                if t1 and len(t1)>=2 and len(t1[0])>=3:
                    tables.append(t1)
                for settings in [{"vertical_strategy":"lines","horizontal_strategy":"lines"},
                                 {"vertical_strategy":"text","horizontal_strategy":"text"}]:
                    try:
                        ts = page.extract_tables(table_settings=settings)
                        for t in ts:
                            if t and len(t)>=2 and len(t[0])>=3:
                                tables.append(t)
                    except Exception: pass
    except Exception: pass
    return tables

def embed_pdf(path: Path, height=900):
    try:
        with open(path, "rb") as f:
            data = f.read()
        b64 = base64.b64encode(data).decode("utf-8")
        html = f'<iframe width="100%" height="{height}" src="data:application/pdf;base64,{b64}" type="application/pdf"></iframe>'
        components.html(html, height=height+10, scrolling=True)
    except Exception as e:
        st.error(f"Unable to preview PDF: {e}")

ITEMS_HEADER_RE = re.compile(
    r"#\s*DESCRIPTION\s+COLIS\s+QT[ÉE]\s+TOTAL\s+EN\s+U\.M\.\s+UNIT[ÉE]\s+P\.\s*U\.\s+TVA\s+TOTAL\s+HT",
    re.I,
)
TVA_BLOCK_RE = re.compile(r"TVA\s*%\s*Base", re.I)

def norm_num2(s):
    if s is None:
        return None
    t = str(s).replace("\xa0", " ").strip().replace(" ", "")
    if re.search(r"\d+\.\d{3},\d{2,4}$", t): t = t.replace(".", "").replace(",", ".")
    else: t = t.replace(",", ".")
    try: return float(t)
    except Exception: return None

def clean_text2(s):
    return re.sub(r"\s+", " ", (s or "")).strip()

def chakirs_extract_headers(full_txt):
    head = {}
    m = re.search(r"FACTURE\s*#\s*([A-Z0-9/\-]+)", full_txt, re.I)
    if m: 
        head["Invoice_No"] = m.group(1).strip()
        head["FACTURE"] = f"FACTURE # {head['Invoice_No']}"
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+D[ûu]e\s+(\d{2}/\d{2}/\d{4})", full_txt, re.I)
    if m: head["Date"], head["Due_Date"] = m.group(1), m.group(2)
    m = re.search(r"HALAL FOOD CHAKIRS BV", full_txt, re.I)
    head["Vendor"] = "HALAL FOOD CHAKIRS BV" if m else ""
    m = re.search(r"IBAN:\s*([A-Z]{2}\s?\d{2}\s?(?:\d{4}\s?){3}\d{3,4})", full_txt, re.I)
    if m: head["IBAN"] = clean_text2(m.group(1))
    m = re.search(r"BIC:\s*([A-Z0-9]+)", full_txt, re.I)
    if m: head["BIC"] = m.group(1).strip()
    m = re.search(r"N[ºo]\s*TVA\s+([A-Z0-9]+)", full_txt, re.I)
    if m: head["Vendor_VAT"] = m.group(1).strip()
    m = re.search(r"N[ºo]\s*Client\s+(\d+)", full_txt, re.I)
    if m: head["Customer_No"] = m.group(1).strip()
    bill = ""
    ship = ""
    mb = re.search(r"ADRESSE DE FACTURATION(.*?)ADRESSE DE LIVRAISON", full_txt, re.S | re.I)
    if mb: bill = clean_text2(mb.group(1))
    ms = re.search(r"ADRESSE DE LIVRAISON(.*?)(?:Geleegweg|N[ºo]\s*TVA|#\s*DESCRIPTION|TVA\s*%)", full_txt, re.S | re.I)
    if ms: ship = clean_text2(ms.group(1))
    head["ADRESSE DE FACTURATION"] = bill
    head["ADRESSE DE LIVRAISON"] = ship
    head["Billing_Address"] = bill
    head["Shipping_Address"] = ship
    for label, key in [
        (r"Total\s+Brut\s+HT\s+([\d\s.,]+)\s*€?", "Total_Brut_HT"),
        (r"Total\s+HT\s+([\d\s.,]+)\s*€?", "Total_HT"),
        (r"TVA\s+([\d\s.,]+)\s*€?", "TVA_EUR"),
        (r"Total\s+TTC\s+([\d\s.,]+)\s*€?", "Total_TTC"),
        (r"Solde\s+([\d\s.,]+)\s*€?", "Solde"),
    ]:
        m = re.search(label, full_txt, re.I)
        if m: head[key] = norm_num2(m.group(1))
    m = re.search(r"Produits\s+(\d+)\s+Colis\s+([\d\s]+)\s+Quantité Totale\s+([\d\s.,]+)\s+Poids Brut Total\s+([\d\s.,]+)\s*Kg\s+Poids Net Total\s+([\d\s.,]+)\s*Kg", full_txt, re.I)
    if m:
        head["Produits"] = int(m.group(1))
        head["Colis"] = int(m.group(2).replace(" ", ""))
        head["Quantite_Totale"] = norm_num2(m.group(3))
        head["Poids_Brut_Total_kg"] = norm_num2(m.group(4))
        head["Poids_Net_Total_kg"] = norm_num2(m.group(5))
    return head

def is_header_or_footer_line(txt_line):
    tl = txt_line.lower()
    if ITEMS_HEADER_RE.search(txt_line): return True
    bad_tokens = ["tva % base", "total ttc", "total brut", "total ht", "solde", "produits", "poids", "quantité totale"]
    return any(tok in tl for tok in bad_tokens)

def chakirs_parse_items_only_table(lines, invoice_no=None):
    items = []
    tail_re = re.compile(
        r"(?P<pack>\d+\s*x\s*\d+|\d+\s*x\s*\d+\s*x\s*\d+|\d+)\s+"
        r"(?P<qte>[\d\s]+)\s+"
        r"(?P<totum>[\d\s.,]+)\s+"
        r"(?P<unit>[A-Za-z]+)\s+"
        r"(?P<pu>[\d.,]+)\s+"
        r"(?P<tva>\d+%)\s+"
        r"(?P<totht>[\d\s.,]+)\s*$"
    )
    i = 0
    while i < len(lines):
        L = clean_text2(lines[i])
        if not L or L.lower().startswith("sh ") or is_header_or_footer_line(L):
            i += 1; continue
        if TVA_BLOCK_RE.search(L):
            break
        msku = re.match(r"^([A-Z0-9]+)\s+(.*)$", L)
        if not msku:
            i += 1; continue
        sku = msku.group(1); right = msku.group(2)
        mtail = tail_re.search(right)
        if mtail:
            desc = right[: mtail.start()].strip()
            pack = (mtail.group("pack") or "").strip()
            qte = norm_num2(mtail.group("qte"))
            totum = norm_num2(mtail.group("totum"))
            unit = mtail.group("unit")
            pu = norm_num2(mtail.group("pu"))
            tva = mtail.group("tva")
            totht = norm_num2(mtail.group("totht"))
            rec = {"Invoice_No": invoice_no, "SKU":sku,"Description":desc,"Colis":pack,"Qte":qte,"Total_UM":totum,"Unite":unit,"PU":pu,"TVA_pct":tva,"Total_HT":totht}
            items.append(rec)
        i += 1
    return items

def chakirs_process(pdf_path: Path):
    all_pages_txt = []
    lines_for_items = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for p in pdf.pages:
            txt = p.extract_text() or ""
            all_pages_txt.append(txt)
            lines_for_items.extend([ln for ln in txt.splitlines() if ln and ln.strip()])
    full_txt = "\n".join(all_pages_txt)
    headers = chakirs_extract_headers(full_txt)
    invoice_no = headers.get("Invoice_No")
    try:
        start_idx = next(i for i, L in enumerate(lines_for_items) if ITEMS_HEADER_RE.search(L))
    except StopIteration:
        start_idx = None
    items = []
    if start_idx is not None:
        sub = []
        for L in lines_for_items[start_idx + 1:]:
            if TVA_BLOCK_RE.search(L): break
            sub.append(L)
        items = chakirs_parse_items_only_table(sub, invoice_no=invoice_no)
    return pd.DataFrame([headers]), pd.DataFrame(items)

def is_header_or_footer(row):
    txt = " ".join([str(c or "").lower() for c in row])
    return (
        txt.strip() == "" or
        txt.startswith("description") or
        "total" in txt or
        "tva" in txt or
        "grand" in txt
    )

def clean_table_rows(table):
    rows = []
    for r in table[1:]:
        if not is_header_or_footer(r):
            rows.append(r)
    return rows

def generic_smart_extract(files, filter_text, decision):
    index_rows = []
    for p in files:
        contains = pdf_contains_text_cached(str(p), filter_text)
        if not contains and use_ocr and not fast_filter and filter_text:
            img = pdf_page_to_image(p, 0, dpi=160)
            if img:
                txt = ocr_on_image(img, tess_langs or "eng")
                contains = filter_text.lower() in (txt or "").lower()
        index_rows.append({"filename": p.name, "contains_filter": contains})
    df_index = pd.DataFrame(index_rows)
    st.subheader("Index (contains text?)")
    st.dataframe(df_index, use_container_width=True)

    if decision.startswith("Only") or decision.startswith("Seulement"):
        kept = [p for p in files if next((r["contains_filter"] for r in index_rows if r["filename"]==p.name), False)]
    else:
        kept = files[:]
    if not kept:
        st.warning("No PDF selected by your choice."); return None, None, df_index

    default_map = json.dumps(LABEL_SYNONYMS, ensure_ascii=False, indent=2)
    mapping_text = st.text_area("Fields & synonyms (JSON)", value=default_map, height=200, key="map_json")
    try:
        label_map = json.loads(mapping_text)
        if not isinstance(label_map, dict): raise ValueError
    except Exception:
        st.error("Invalid JSON; using default map.")
        label_map = LABEL_SYNONYMS

    include_items = not skip_items

    tmpdir = Path(tempfile.gettempdir())
    headers_csv = tmpdir / "headers_batch.csv"
    items_csv = tmpdir / "items_batch.csv"
    if headers_csv.exists(): headers_csv.unlink()
    if items_csv.exists(): items_csv.unlink()

    total = len(kept)
    all_rows_mem = []
    items_rows_mem = []
    progress = st.progress(0)
    status = st.empty()

    for i, pdf in enumerate(kept, start=1):
        page_texts = extract_page_texts_cached(str(pdf))
        if use_ocr and not fast_filter and (not any(page_texts) or len("".join(page_texts))<50):
            img = pdf_page_to_image(pdf, 0, dpi=170)
            if img:
                txt = ocr_on_image(img, tess_langs or "eng")
                page_texts = [txt]
        LINES = []
        for ptxt in page_texts:
            LINES.extend(ptxt.splitlines() if ptxt else [])
        row = {"filename": pdf.name,
               "contains_filter": next((r["contains_filter"] for r in index_rows if r["filename"]==p.name), False)}
        for field, synonyms in label_map.items():
            val = smart_find_value(LINES, synonyms)
            row[field] = val
        if "Total_TTC" in row and row["Total_TTC"]:
            num = normalize_amount(row["Total_TTC'])
            if num is not None: row["Total_TTC_num"] = num
        all_rows_mem.append(row)

        if include_items:
            tables = detect_tables_pdfplumber(str(pdf))
            for t in tables:
                headers = [str(h or "").strip() for h in t[0]]
                for r in clean_table_rows(t):
                    rec = {"filename": pdf.name, "contains_filter": row["contains_filter"]}
                    inv = row.get("Invoice_No") or row.get("invoice_no") or ""
                    rec["Invoice_No"] = inv
                    for ci, c in enumerate(r):
                        key = headers[ci] if ci < len(headers) else f"C{ci+1}"
                        rec[key] = cleanup_multiline(c)
                    items_rows_mem.append(rec)

        if len(all_rows_mem) >= batch_size:
            pd.DataFrame(all_rows_mem).to_csv(headers_csv, mode="a", header=not headers_csv.exists(), index=False)
            all_rows_mem = []
        if include_items and len(items_rows_mem) >= batch_size:
            pd.DataFrame(items_rows_mem).to_csv(items_csv, mode="a", header=not items_csv.exists(), index=False)
            items_rows_mem = []

        progress.progress(int(i*100/total))
        status.markdown(f"<div class='stat-card'><p class='stat-title'>Progress</p><p class='stat-value'>{i} / {total}</p></div>", unsafe_allow_html=True)

    if all_rows_mem:
        pd.DataFrame(all_rows_mem).to_csv(headers_csv, mode="a", header=not headers_csv.exists(), index=False)
    if include_items and items_rows_mem:
        pd.DataFrame(items_rows_mem).to_csv(items_csv, mode="a", header=not items_csv.exists(), index=False)

    df = pd.read_csv(headers_csv) if headers_csv.exists() else pd.DataFrame()
    dfi = pd.read_csv(items_csv) if items_csv.exists() and include_items else None
    return df, dfi, df_index

# -------- Tab 1: Smart --------
with tabs[0]:
    st.header("Smart Extract (Pro) — Folder Picker + Combined Excel (Filter Fix)")
    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        filter_text = st.text_input("Filter text (optional)", value="")
    with col2:
        decision = st.radio("Process:", ["Only PDFs that contain the text", "All PDFs"], index=0)
    with col3:
        use_chakirs = st.checkbox("✅ Use 'Chakirs' profile (template-specific)", value=True)

    st.markdown("### Choose input PDFs")
    uploaded = st.file_uploader("Upload PDFs (multiple)", type="pdf", accept_multiple_files=True)
    st.markdown("**— OR —**")
    folder = st.text_input("Folder path (processed recursively for *.pdf)", value="")
    load_btn = st.button("Load PDFs from folder")

    files = []
    if uploaded:
        tmp = Path(tempfile.gettempdir())
        for f in uploaded:
            p = tmp / f.name
            with open(p, "wb") as out: out.write(f.getbuffer())
            files.append(p)
    if load_btn and folder.strip():
        p = Path(folder).expanduser()
        if p.exists() and p.is_dir():
            for g in p.rglob("*.pdf"):
                files.append(g.resolve())
        else:
            st.error("Folder not found or not a directory.")

    run = st.button("Index then extract")

    if run:
        if not files:
            st.warning("No PDF selected."); st.stop()

        st.markdown("<div class='stat-card'><p class='stat-title'>Loaded files</p><p class='stat-value'>{}</p></div>".format(len(files)), unsafe_allow_html=True)

        if use_chakirs:
            index_rows = [{"filename": p.name, "contains_filter": pdf_contains_text_cached(str(p), filter_text)} for p in files]
            df_index = pd.DataFrame(index_rows)
            st.subheader("Index (contains text?)")
            st.dataframe(df_index, use_container_width=True)
            if decision.startswith("Only") or decision.startswith("Seulement"):
                kept = [p for p in files if next((r["contains_filter"] for r in index_rows if r["filename"]==p.name), False)]
            else:
                kept = files[:]
            if not kept:
                st.warning("No PDF selected by your choice."); st.stop()

            st.markdown("<div class='stat-card'><p class='stat-title'>To process</p><p class='stat-value'>{}</p></div>".format(len(kept)), unsafe_allow_html=True)

            allh = []; alli = []
            progress = st.progress(0)
            for i, pdf in enumerate(kept, start=1):
                dfh, dfi_local = chakirs_process(pdf)
                if "FACTURE" not in dfh.columns and "Invoice_No" in dfh.columns:
                    dfh["FACTURE"] = "FACTURE # " + dfh["Invoice_No"].fillna("")
                if "ADRESSE DE FACTURATION" not in dfh.columns and "Billing_Address" in dfh.columns:
                    dfh["ADRESSE DE FACTURATION"] = dfh["Billing_Address"]
                if "ADRESSE DE LIVRAISON" not in dfh.columns and "Shipping_Address" in dfh.columns:
                    dfh["ADRESSE DE LIVRAISON"] = dfh["Shipping_Address"]
                dfh["filename"] = pdf.name
                dfh["contains_filter"] = next((r["contains_filter"] for r in index_rows if r["filename"]==pdf.name), False)
                allh.append(dfh)
                if not dfi_local.empty:
                    dfi_local["filename"] = pdf.name
                    dfi_local["contains_filter"] = dfh["contains_filter"].iloc[0]
                    inv = dfh["Invoice_No"].iloc[0] if "Invoice_No" in dfh.columns else ""
                    dfi_local["Invoice_No"] = inv
                    alli.append(dfi_local)
                progress.progress(int(i*100/len(kept)))

            df = pd.concat(allh, ignore_index=True) if allh else pd.DataFrame()
            dfi = pd.concat(alli, ignore_index=True) if alli else None
        else:
            df, dfi, df_index = generic_smart_extract(files, filter_text, decision)

        # --- Display-only filters (FIX) ---
        if df is not None:
            st.subheader("Detected fields")
            hf = st.text_input("Filter in detected fields (substring)", value="", key="hdr_filter")
            df_show = df.copy()
            if hf:
                mask = pd.Series([False]*len(df_show))
                for c in df_show.columns:
                    mask = mask | df_show[c].astype(str).str.contains(hf, case=False, na=False)
                df_show = df_show[mask]
            st.dataframe(df_show, use_container_width=True)

            if dfi is not None:
                st.subheader("Item lines (table only)")
                if "Invoice_No" not in dfi.columns and "Invoice_No" in df.columns:
                    dfi["Invoice_No"] = df["Invoice_No"].iloc[0]
                if "Invoice_No" not in dfi.columns:
                    dfi["Invoice_No"] = ""
                ifilter = st.text_input("Filter in items (substring)", value="", key="itm_filter")
                dfi_show = dfi.copy()
                if ifilter:
                    mask2 = pd.Series([False]*len(dfi_show))
                    for c in dfi_show.columns:
                        mask2 = mask2 | dfi_show[c].astype(str).str.contains(ifilter, case=False, na=False)
                    dfi_show = dfi_show[mask2]
                st.dataframe(dfi_show, use_container_width=True)

            # --- Combined Excel (always full data, not filtered views) ---
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="headers", index=False)
                if dfi is not None:
                    dfi.to_excel(writer, sheet_name="items", index=False)
            bio.seek(0)
            st.download_button("⬇️ Download Combined Excel", data=bio, file_name="export-pro-v6_3_1_combined.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------- Tab 2: Templates --------
with tabs[1]:
    st.header("Templates (full-page)")
    st.caption("Page slider + width zoom. Draw rectangles; relative coordinates saved to templates.json.")
    try:
        from streamlit_drawable_canvas import st_canvas
        canvas_ok = True
    except Exception:
        canvas_ok = False
        st.warning("Canvas component unavailable (versions).")

    sample = st.file_uploader("Sample PDF", type="pdf", key="tmpl_sample_full")
    if sample is not None and canvas_ok:
        tmp = Path(tempfile.gettempdir()); sp = tmp / sample.name
        with open(sp, "wb") as out: out.write(sample.getbuffer())
        try:
            with pdfplumber.open(sp) as pdf: n_pages = len(pdf.pages)
        except Exception: n_pages = 1
        colp1, colp2 = st.columns([1,1])
        with colp1:
            page_index = st.number_input("Page (0-based)", min_value=0, max_value=max(0, n_pages-1), value=0, step=1)
        with colp2:
            fit_width = st.slider("Display width (px)", min_value=600, max_value=1600, value=1100, step=50)
        base_img = pdf_page_to_image(sp, page_num=int(page_index), dpi=preview_dpi)
        if base_img is None:
            st.error("Cannot render page. Poppler required.")
        else:
            w0, h0 = base_img.size; ratio = fit_width / float(w0); disp_h = int(h0 * ratio)
            img = base_img.resize((fit_width, disp_h))
            st.write(f"Page {int(page_index)+1}/{n_pages} — rendered {w0}x{h0} → displayed {fit_width}x{disp_h}")
            from streamlit_drawable_canvas import st_canvas
            canvas = st_canvas(
                fill_color="rgba(0,0,0,0)",
                stroke_width=2,
                stroke_color="#FF0000",
                background_image=img,
                height=disp_h,
                width=fit_width,
                drawing_mode="rect",
                key=f"canvas_full_{page_index}",
            )
            rects = []
            if canvas.json_data and "objects" in canvas.json_data:
                for o in canvas.json_data["objects"]:
                    if o.get("type") in ("rect","rectangle"):
                        left=o.get("left",0); top=o.get("top",0); rw=o.get("width",0); rh=o.get("height",0)
                        rects.append([left/fit_width, top/disp_h, rw/fit_width, rh/disp_h])
            st.info(f"Rectangles: {len(rects)}")
            names_txt = st.text_area("Field names (same order as rectangles)", value="FACTURE, ADRESSE DE FACTURATION, ADRESSE DE LIVRAISON")
            tmpl_name = st.text_input("Template name", value=f"Template_{len(templates)+1}")
            save_btn = st.button("Save/Update template")
            if save_btn:
                names=[n.strip() for n in names_txt.split(",") if n.strip()]
                if not rects:
                    st.error("Please draw at least one rectangle.")
                elif len(names)!=len(rects):
                    st.error("Number of names must match rectangles.")
                else:
                    zones=[{"field_name": nm, "page": int(page_index), "bbox": bbox, "method": "ocr" if use_ocr else "text"} for nm, bbox in zip(names, rects)]
                    exist = next((t for t in templates if t["name"]==tmpl_name), None)
                    if exist: exist["zones"]=zones; exist["langs"]=tess_langs
                    else: templates.append({"name": tmpl_name, "zones": zones, "langs": tess_langs})
                    TEMPLATES_FILE.write_text(json.dumps(templates, ensure_ascii=False, indent=2), encoding="utf-8")
                    st.success(f"Template '{tmpl_name}' saved ({len(zones)} zones).")
    st.subheader("Apply a template")
    if templates:
        name = st.selectbox("Template", [t["name"] for t in templates])
        uploaded2 = st.file_uploader("PDFs (multiple)", type="pdf", accept_multiple_files=True, key="tmpl_apply_full")
        run2 = st.button("Apply")
        if run2:
            tmpl = next(t for t in templates if t["name"]==name)
            zones = tmpl.get("zones", [])
            langs = tmpl.get("langs", tess_langs)
            files=[]; tmp = Path(tempfile.gettempdir())
            for f in uploaded2 or []:
                p = tmp / f.name
                with open(p,"wb") as out: out.write(f.getbuffer())
                files.append(p)
            if not files:
                st.warning("No PDF.")
            else:
                rows=[]
                for pdf in files:
                    row={"filename": pdf.name}
                    for z in zones:
                        img = pdf_page_to_image(pdf, page_num=int(z.get("page",0)), dpi=ocr_zone_dpi)
                        val = ""
                        if img:
                            x,y,w,h = z["bbox"]
                            crop = Image.Image.crop(img, (int(x*img.width), int(y*img.height), int((x+w)*img.width), int((y+h)*img.height)))
                            if use_ocr or z.get("method")=="ocr":
                                val = ocr_on_image(crop, tess_langs or "eng")
                            else:
                                with pdfplumber.open(str(pdf)) as doc:
                                    val = doc.pages[int(z.get("page",0))].extract_text() or ""
                        row[z["field_name"]] = cleanup_multiline(val)
                    rows.append(row)
                df = pd.DataFrame(rows)
                st.dataframe(df, use_container_width=True)
                bio = io.BytesIO()
                with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="template_apply", index=False)
                bio.seek(0)
                st.download_button("⬇️ Download Excel (Template)", data=bio, file_name="export-template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No template yet. Create one above.")

# -------- Tab 3: PDF Preview --------
with tabs[2]:
    st.header("PDF Preview (entire document)")
    up = st.file_uploader("Upload one or more PDFs for preview", type="pdf", accept_multiple_files=True, key="preview_full")
    if up:
        tmp = Path(tempfile.gettempdir())
        for i,f in enumerate(up, start=1):
            p = tmp / f.name
            with open(p,"wb") as out: out.write(f.getbuffer())
            st.write(f"**{i}.** {p.name}")
            with open(p, "rb") as rb:
                b64 = base64.b64encode(rb.read()).decode("utf-8")
            components.html(f'<iframe width="100%" height="900" src="data:application/pdf;base64,{b64}" type="application/pdf"></iframe>', height=910, scrolling=True)
            st.divider()

# -------- Tab 4: Notes --------
with tabs[3]:
    st.header("Notes")
    st.markdown("- **Filter fix**: headers/items DataFrames remain intact; filters only affect on-screen views.")
    st.markdown("- **Folder picker**: give a folder path; app reads all `*.pdf` recursively.")
    st.markdown("- **Combined Excel**: single workbook with `headers` + `items` for all PDFs processed.")
    st.markdown("- **Detected fields include**: FACTURE, ADRESSE DE FACTURATION, ADRESSE DE LIVRAISON (Chakirs).")
    st.markdown("- **Items always include Invoice_No** for grouping/joining.")
    st.markdown("- **Scale**: index first + batch CSVs + progress → handles thousands of PDFs with low RAM.")
    st.markdown("- **Tip**: For text PDFs (not scanned), keep OCR off for maximal speed.")
