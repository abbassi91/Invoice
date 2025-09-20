# # invoice_extractor_sidebar_streamlit_tables.py
# import io, re, shutil, tempfile
# import streamlit as st
# import pdfplumber
# import pandas as pd

# # ---------- Optional OCR imports (safe import) ----------
# try:
#     from pdf2image import convert_from_path
#     import pytesseract
#     OCR_LIBS = True
# except Exception:
#     OCR_LIBS = False

# # ================= I18N (UI only; data NOT translated) =================
# TRANSLATIONS = {
#     "fr": {
#         "app_title": "CHATAR — Extracteur de Factures",
#         "app_caption": "Téléversez des PDF → Pré-filtre rapide → Extraction en-têtes + lignes.",
#         "upload_label": "Téléverser un ou plusieurs PDF",
#         "filter_label": "Filtre global (mots-clés, séparés par des virgules). Vide = traiter tous",
#         "extract_btn": "▶️ Extraire",
#         "kpi_headers": "Lignes en-têtes",
#         "kpi_items": "Lignes articles",
#         "warn_upload": "Veuillez téléverser au moins un PDF.",
#         "warn_no_match": "Aucun PDF ne correspond au filtre.",
#         "headers_title": "En-têtes",
#         "items_title": "Articles",
#         "items_empty": "Aucune ligne d’articles détectée (tableau introuvable).",
#         "export_btn": "⬇️ Exporter en Excel",
#         "preview_title": "Aperçu du premier PDF (optionnel)",
#         "previewing": "Aperçu : **{name}**",
#         "lang_label": "Langue de l’interface",
#         "language_fr": "Français",
#         "language_en": "English",
#         "language_nl": "Nederlands",
#         "language_ar": "العربية",
#         "filter_info": "Insensible à la casse. « huile, sucre » ⇒ correspond si l’un apparaît.",
#         "ocr_label": "Utiliser l’OCR (Tesseract) pour PDFs non recherchables",
#         "ocr_langs": "Langues OCR (ex. fra+eng+nld+ara)",
#         "ocr_hint": "Activez l’OCR si certains PDFs sont des scans.",
#         "ocr_missing": "OCR indisponible : installez pdf2image + pytesseract & Tesseract/Poppler, ou désactivez l’OCR.",
#         "not_searchable_found": "PDF non recherchables (ignorés car OCR est OFF) :",
#         "prefilter_summary": "Pré-filtre : {kept} gardés, {skipped} ignorés.",
#         "headers_filter": "Filtrer les en-têtes (toutes colonnes)",
#         "items_filter": "Filtrer les articles (toutes colonnes)",
#         "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
#         "toggle_sidebar": "Afficher / masquer la barre latérale",
#     },
#     "en": {
#         "app_title": "CHATAR — Invoice Extractor",
#         "app_caption": "Upload PDFs → Fast prefilter → Extract headers + items.",
#         "upload_label": "Upload one or more PDFs",
#         "filter_label": "Global filter (keywords, comma-separated). Empty = process all",
#         "extract_btn": "▶️ Extract",
#         "kpi_headers": "Header rows",
#         "kpi_items": "Item rows",
#         "warn_upload": "Please upload at least one PDF.",
#         "warn_no_match": "No PDF matched your filter.",
#         "headers_title": "Headers",
#         "items_title": "Items",
#         "items_empty": "No items detected (table body not found).",
#         "export_btn": "⬇️ Export to Excel",
#         "preview_title": "Preview first PDF (optional)",
#         "previewing": "Previewing: **{name}**",
#         "lang_label": "Interface language",
#         "language_fr": "Français",
#         "language_en": "English",
#         "language_nl": "Nederlands",
#         "language_ar": "العربية",
#         "filter_info": "Case-insensitive. ‘oil, sugar’ ⇒ match if any appears.",
#         "ocr_label": "Use OCR (Tesseract) for non-searchable PDFs",
#         "ocr_langs": "OCR languages (e.g. fra+eng+nld+ara)",
#         "ocr_hint": "Enable OCR if some PDFs are scans.",
#         "ocr_missing": "OCR not available: install pdf2image + pytesseract & Tesseract/Poppler, or turn OCR off.",
#         "not_searchable_found": "Non-searchable PDFs (skipped because OCR is OFF):",
#         "prefilter_summary": "Prefilter: {kept} kept, {skipped} skipped.",
#         "headers_filter": "Filter headers (all columns)",
#         "items_filter": "Filter items (all columns)",
#         "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
#         "toggle_sidebar": "Show / hide sidebar",
#     },
#     "nl": {
#         "app_title": "CHATAR — Factuur Extractor",
#         "app_caption": "Upload PDF’s → Snelle voorfilter → Koppen + regels extraheren.",
#         "upload_label": "Upload één of meerdere PDF's",
#         "filter_label": "Globaal filter (komma-gescheiden). Leeg = alles verwerken",
#         "extract_btn": "▶️ Extractie",
#         "kpi_headers": "Kopregels",
#         "kpi_items": "Artikelregels",
#         "warn_upload": "Upload minstens één PDF.",
#         "warn_no_match": "Geen PDF komt overeen met het filter.",
#         "headers_title": "Koppen",
#         "items_title": "Artikelen",
#         "items_empty": "Geen artikelen gedetecteerd (tabel niet gevonden).",
#         "export_btn": "⬇️ Exporteren naar Excel",
#         "preview_title": "Voorbeeld eerste PDF (optioneel)",
#         "previewing": "Voorbeeld: **{name}**",
#         "lang_label": "Taal van de interface",
#         "language_fr": "Français",
#         "language_en": "English",
#         "language_nl": "Nederlands",
#         "language_ar": "العربية",
#         "filter_info": "Niet-hoofdlettergevoelig. ‘olie, suiker’ ⇒ matcht als een voorkomt.",
#         "ocr_label": "OCR gebruiken (Tesseract) voor niet-doorzoekbare PDF’s",
#         "ocr_langs": "OCR-talen (bijv. fra+eng+nld+ara)",
#         "ocr_hint": "Schakel OCR in als sommige PDF’s scans zijn.",
#         "ocr_missing": "OCR niet beschikbaar: installeer pdf2image + pytesseract & Tesseract/Poppler, of zet OCR uit.",
#         "not_searchable_found": "Niet-doorzoekbare PDF’s (overgeslagen omdat OCR UIT staat):",
#         "prefilter_summary": "Voorfilter: {kept} behouden, {skipped} overgeslagen.",
#         "headers_filter": "Koppen filteren (alle kolommen)",
#         "items_filter": "Artikelen filteren (alle kolommen)",
#         "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
#         "toggle_sidebar": "Zijbalk tonen / verbergen",
#     },
#     "ar": {
#         "app_title": "شاطر — أداة استخراج الفواتير",
#         "app_caption": "ارفع ملفات PDF → فلترة سريعة → استخراج العناوين والبنود.",
#         "upload_label": "ارفع ملف/ملفات PDF",
#         "filter_label": "مرشح عام (كلمات مفصولة بفواصل). فارغ = معالجة الجميع",
#         "extract_btn": "▶️ استخراج",
#         "kpi_headers": "صفوف العناوين",
#         "kpi_items": "صفوف البنود",
#         "warn_upload": "الرجاء رفع ملف PDF واحد على الأقل.",
#         "warn_no_match": "لا يوجد ملف يطابق المرشح.",
#         "headers_title": "العناوين",
#         "items_title": "البنود",
#         "items_empty": "لم يتم العثور على بنود (الجدول غير موجود).",
#         "export_btn": "⬇️ تصدير إلى Excel",
#         "preview_title": "معاينة أول ملف (اختياري)",
#         "previewing": "المعاينة: **{name}**",
#         "lang_label": "لغة الواجهة",
#         "language_fr": "Français",
#         "language_en": "English",
#         "language_nl": "Nederlands",
#         "language_ar": "العربية",
#         "filter_info": "مطابقة غير حساسة لحالة الأحرف. « زيت، سكر » ⇒ تطابق إذا وُجد أي منها.",
#         "ocr_label": "استخدم OCR (Tesseract) للملفات غير القابلة للبحث",
#         "ocr_langs": "لغات OCR (مثال: fra+eng+nld+ara)",
#         "ocr_hint": "فعّل OCR إذا كانت بعض الملفات صورًا ممسوحة.",
#         "ocr_missing": "‏OCR غير متاح: ثبّت pdf2image + pytesseract و Tesseract/Poppler، أو عطّله.",
#         "not_searchable_found": "ملفات غير قابلة للبحث (تخطيناها لأن OCR مُعطّل):",
#         "prefilter_summary": "الفلترة المسبقة: {kept} تم الاحتفاظ بها، {skipped} تم تجاوزها.",
#         "headers_filter": "تصفية العناوين (كل الأعمدة)",
#         "items_filter": "تصفية البنود (كل الأعمدة)",
#         "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
#         "toggle_sidebar": "إظهار/إخفاء الشريط الجانبي",
#     },
# }
# def t(key: str, lang: str, **kwargs) -> str:
#     s = TRANSLATIONS.get(lang, TRANSLATIONS["fr"]).get(key, key)
#     return s.format(**kwargs) if kwargs else s

# # ================== Page config ==================
# st.set_page_config(page_title="CHATAR — Invoice Extractor", layout="wide")

# # ---------- Session state init ----------
# for k, v in {
#     "df_headers": None,
#     "df_items": None,
#     "sidebar_collapsed": False,
#     "lang": "fr",
# }.items():
#     st.session_state.setdefault(k, v)

# # ==== TOP BAR: toggle first, THEN maybe inject CSS (so one click works) ====
# col_toggle, col_title = st.columns([0.12, 0.88])
# with col_toggle:
#     if st.button("☰", help=t("toggle_sidebar", st.session_state["lang"]), use_container_width=True, key="toggle_sidebar_btn"):
#         st.session_state["sidebar_collapsed"] = not st.session_state["sidebar_collapsed"]

# def apply_sidebar_collapse():
#     if st.session_state.get("sidebar_collapsed", False):
#         st.markdown("""
#         <style>
#         section[data-testid="stSidebar"],
#         div[data-testid="stSidebar"]{
#             width:0 !important; min-width:0 !important; max-width:0 !important;
#             overflow:hidden !important; padding:0 !important;
#         }
#         div[data-testid="stSidebarNav"]{ display:none !important; }
#         </style>
#         """, unsafe_allow_html=True)
# apply_sidebar_collapse()

# with col_title:
#     lang_for_title = st.session_state["lang"]
#     st.markdown(f"**{t('basmala', lang_for_title)}**")
#     st.markdown(f"## {t('app_title', lang_for_title)}")
#     st.caption(t("app_caption", lang_for_title))

# # -------- Sidebar controls --------
# with st.sidebar:
#     lang = st.selectbox(
#         TRANSLATIONS["fr"]["lang_label"],
#         options=[
#             ("fr", "🇫🇷 " + TRANSLATIONS["fr"]["language_fr"]),
#             ("en", "🇬🇧 " + TRANSLATIONS["en"]["language_en"]),
#             ("nl", "🇳🇱 " + TRANSLATIONS["nl"]["language_nl"]),
#             ("ar", "🇲🇦 " + TRANSLATIONS["ar"]["language_ar"]),
#         ],
#         index=["fr","en","nl","ar"].index(st.session_state["lang"]),
#         format_func=lambda opt: opt[1],
#         key="lang_picker"
#     )[0]
#     st.session_state["lang"] = lang

#     files = st.file_uploader(t("upload_label", lang), type="pdf", accept_multiple_files=True, key="uploader")
#     filter_text = st.text_input(t("filter_label", lang), value="", key="prefilter_text")
#     st.caption("ℹ️ " + t("filter_info", lang))

#     use_ocr = st.checkbox(t("ocr_label", lang), value=False, key="use_ocr")
#     ocr_langs = st.text_input(t("ocr_langs", lang), value="fra+eng+nld+ara", key="ocr_langs")
#     st.caption("🧠 " + t("ocr_hint", lang))

#     go = st.button(t("extract_btn", lang), use_container_width=True, key="extract_btn")

# # ---------------- Helpers ----------------
# def parse_terms(s: str):
#     return [x.strip().lower() for x in (s or "").split(",") if x.strip()]

# def pdf_text_allpages(file_bytes: bytes):
#     texts = []
#     try:
#         with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
#             for p in pdf.pages:
#                 txt = p.extract_text() or ""
#                 if txt:
#                     texts.append(txt)
#     except Exception:
#         pass
#     return "\n".join(texts)

# def pdf_ocr_first_pages(file_bytes: bytes, langs="eng", pages=2):
#     if not (OCR_LIBS and shutil.which("tesseract")):
#         return ""
#     try:
#         imgs = convert_from_path(io.BytesIO(file_bytes), first_page=1, last_page=pages, dpi=180)
#     except TypeError:
#         with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tmp:
#             tmp.write(file_bytes); tmp.flush()
#             imgs = convert_from_path(tmp.name, first_page=1, last_page=pages, dpi=180)
#     except Exception:
#         return ""
#     out = []
#     for im in imgs:
#         try:
#             out.append(pytesseract.image_to_string(im, lang=langs) or "")
#         except Exception:
#             continue
#     return "\n".join(out)

# ITEMS_HEADER_RE = re.compile(
#     r"#\s*DESCRIPTION\s+COLIS\s+QT[ÉE]\s+TOTAL\s+EN\s+U\.M\.\s+UNIT[ÉE]\s+P\.\s*U\.\s+TVA\s+TOTAL\s+HT",
#     re.I
# )
# TVA_BLOCK_RE = re.compile(r"TVA\s*%\s*Base", re.I)

# def norm_num2(s):
#     if s is None: return None
#     t = str(s).replace("\xa0"," ").strip().replace(" ","")
#     if re.search(r"\d+\.\d{3},\d{2,4}$", t): t = t.replace(".","").replace(",",".")
#     else: t = t.replace(",",".")
#     try: return float(t)
#     except Exception: return None

# COUNTRIES = {
#     "canada","france","belgium","belgique","belgië","netherlands","pays-bas","nederland",
#     "morocco","maroc","algeria","algérie","tunisia","tunisie","spain","espagne","italy","italie",
#     "germany","allemagne","luxembourg","suisse","switzerland","united kingdom","uk","angleterre",
#     "usa","united states","etats-unis","états-unis"
# }
# def last_token_country(s: str) -> str:
#     if not s: return ""
#     tail = s.split("|")[-1].strip().lower()
#     words = re.split(r"[,\s]+", tail)
#     for w in reversed(words):
#         if w in COUNTRIES:
#             return w.title() if len(w) > 2 else w.upper()
#     if tail in COUNTRIES: return tail.title()
#     return ""

# def chakirs_extract_headers(full_txt: str) -> dict:
#     head = {}
#     m = re.search(r"FACTURE\s*#\s*([A-Z0-9/\-]+)", full_txt, re.I)
#     if m:
#         head["Invoice_No"] = m.group(1).strip()
#     m = re.search(r"(\d{2}/\d{2}/\d{4})\s+D[ûu]e\s+(\d{2}/\d{2}/\d{4})", full_txt, re.I)
#     if m:
#         head["Date"], head["Due_Date"] = m.group(1), m.group(2)

#     def parse_block(full_txt, title_re, stops):
#         mm = re.search(title_re, full_txt, re.I)
#         if not mm: return "", ""
#         start = mm.end(); end = len(full_txt)
#         for sr in stops:
#             ms = re.search(sr, full_txt[start:], re.I)
#             if ms:
#                 end = start + ms.start(); break
#         chunk = full_txt[start:end]
#         lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]
#         if not lines: return "", ""
#         name = re.sub(r"\s+"," ",lines[0]).strip()
#         addr = "\n".join(lines[1:]) if len(lines)>1 else ""
#         return name, addr

#     ship_name, ship_addr = parse_block(
#         full_txt, r"ADRESSE DE LIVRAISON\s*",
#         [r"#\s*DESCRIPTION", r"TVA\s*%\s*Base", r"N[ºo]\s*TVA"]
#     )
#     head["Shipping_Name"]    = ship_name
#     head["Shipping_Address"] = ship_addr.strip()
#     head["Shipping_Country"] = last_token_country(ship_addr)
#     return head

# def chakirs_parse_items(lines, invoice_no=None):
#     items = []
#     tail_re = re.compile(
#         r"(?P<pack>\d+\s*x\s*\d+|\d+\s*x\s*\d+\s*x\s*\d+|\d+)\s+"
#         r"(?P<qte>[\d\s]+)\s+"
#         r"(?P<totum>[\d\s.,]+)\s+"
#         r"(?P<unit>[A-Za-z]+)\s+"
#         r"(?P<pu>[\d.,]+)\s+"
#         r"(?P<tva>\d+%)\s+"
#         r"(?P<totht>[\d\s.,]+)\s*$"
#     )
#     i = 0
#     while i < len(lines):
#         L = re.sub(r"\s+", " ", lines[i]).strip()
#         if not L or L.lower().startswith("sh "):
#             i += 1; continue
#         if TVA_BLOCK_RE.search(L): break
#         msku = re.match(r"^([A-Z0-9]+)\s+(.*)$", L)
#         if not msku:
#             i += 1; continue
#         sku = msku.group(1); right = msku.group(2)
#         mtail = tail_re.search(right)
#         if mtail:
#             desc  = right[: mtail.start()].strip()
#             pack  = (mtail.group("pack") or "").strip()
#             qte   = norm_num2(mtail.group("qte"))
#             totum = norm_num2(mtail.group("totum"))
#             unit  = mtail.group("unit")
#             pu    = norm_num2(mtail.group("pu"))
#             tva   = mtail.group("tva")
#             totht = norm_num2(mtail.group("totht"))
#             items.append({
#                 "Invoice_No": invoice_no, "SKU": sku, "Description": desc, "Colis": pack,
#                 "Qte": qte, "Total_UM": totum, "Unite": unit, "PU": pu, "TVA_pct": tva, "Total_HT": totht
#             })
#         i += 1
#     return items

# def extract_one_pdf(name: str, raw_bytes: bytes, filter_text: str):
#     contains = True
#     all_pages = []
#     try:
#         with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
#             for p in pdf.pages:
#                 t = p.extract_text() or ""
#                 all_pages.append(t)
#     except Exception:
#         all_pages = []
#     joined = "\n".join(all_pages)
#     if filter_text:
#         contains = filter_text.lower() in joined.lower()

#     head = chakirs_extract_headers(joined)
#     m = re.search(r"FACTURE\s*#\s*([A-Z0-9/\-]+)", joined, re.I)
#     if m: head["Invoice_No"] = m.group(1).strip()
#     m = re.search(r"(\d{2}/\d{2}/\d{4})\s+D[ûu]e\s+(\d{2}/\d{2}/\d{4})", joined, re.I)
#     if m:
#         head["Date"], head["Due_Date"] = m.group(1), m.group(2)
#     head["filename"] = name
#     head["contains_filter"] = contains

#     lines = []
#     for t in all_pages:
#         lines.extend([ln for ln in t.splitlines() if ln and ln.strip()])
#     items_df = pd.DataFrame()
#     try:
#         start_idx = next(i for i, L in enumerate(lines) if ITEMS_HEADER_RE.search(L))
#     except StopIteration:
#         start_idx = None
#     if start_idx is not None:
#         sub = []
#         for L in lines[start_idx+1:]:
#             if TVA_BLOCK_RE.search(L): break
#             sub.append(L)
#         rows = chakirs_parse_items(sub, invoice_no=head.get("Invoice_No"))
#         if rows:
#             items_df = pd.DataFrame(rows)
#     return contains, pd.DataFrame([head]), items_df

# # ================== EXTRACT (only on button press) ==================
# if go:
#     if not files:
#         st.warning(t("warn_upload", st.session_state["lang"]))
#         st.stop()

#     # Collapse sidebar immediately after clicking Extract
#     st.session_state["sidebar_collapsed"] = True
#     apply_sidebar_collapse()

#     # Phase A: Prefilter (with visible progress)
#     with st.spinner("Indexing PDFs..."):
#         terms = parse_terms(st.session_state.get("prefilter_text", ""))
#         non_searchable = []
#         kept_files, skipped_files = [], []
#         progress = st.progress(0)
#         status = st.empty()

#         for i, f in enumerate(files, start=1):
#             fb = f.getbuffer()

#             if not terms:
#                 kept_files.append(f)
#             else:
#                 plain = pdf_text_allpages(fb)
#                 searchable = bool(plain.strip())
#                 found = False

#                 if searchable:
#                     pl = plain.lower()
#                     for term in terms:
#                         if term in pl:
#                             found = True; break

#                 if not found and not searchable:
#                     non_searchable.append(f.name)
#                     if st.session_state.get("use_ocr") and OCR_LIBS and shutil.which("tesseract"):
#                         ocr_txt = pdf_ocr_first_pages(fb, langs=st.session_state.get("ocr_langs") or "eng", pages=2)
#                         ol = (ocr_txt or "").lower()
#                         for term in terms:
#                             if term in ol:
#                                 found = True; break

#                 if found or not terms:
#                     kept_files.append(f)
#                 elif terms:
#                     skipped_files.append(f)

#             progress.progress(int(i * 100 / len(files)))
#             status.write(f"Prefiltered {i}/{len(files)}")

#     st.info(t("prefilter_summary", st.session_state["lang"], kept=len(kept_files), skipped=len(skipped_files)))
#     if parse_terms(st.session_state.get("prefilter_text", "")) and non_searchable and not st.session_state.get("use_ocr"):
#         st.warning("⚠️ " + t("not_searchable_found", st.session_state["lang"]) + " " + ", ".join(non_searchable))

#     if parse_terms(st.session_state.get("prefilter_text", "")) and not kept_files:
#         st.warning(t("warn_no_match", st.session_state["lang"]))
#         st.session_state["df_headers"] = pd.DataFrame()
#         st.session_state["df_items"] = pd.DataFrame()
#         st.stop()

#     # Phase B: Extract (with visible progress)
#     with st.spinner("Extracting data..."):
#         all_headers, all_items = [], []
#         progress2 = st.progress(0)
#         status2 = st.empty()

#         for i, f in enumerate(kept_files, start=1):
#             _, dfh, dfi = extract_one_pdf(f.name, f.getbuffer(), st.session_state.get("prefilter_text", ""))
#             all_headers.append(dfh)
#             if not dfi.empty:
#                 dfi["filename"] = f.name
#                 all_items.append(dfi)
#             progress2.progress(int(i * 100 / len(kept_files)))
#             status2.write(f"Processed {i}/{len(kept_files)}")

#     df_full = pd.concat(all_headers, ignore_index=True) if all_headers else pd.DataFrame()
#     dfi = pd.concat(all_items,   ignore_index=True) if all_items   else pd.DataFrame()

#     # Keep only requested header columns
#     keep_cols = ["Invoice_No", "Date", "Due_Date", "Shipping_Name", "Shipping_Country"]
#     df_headers = df_full[[c for c in keep_cols if c in df_full.columns]].copy()

#     # Save to session state (persistent across reruns)
#     st.session_state["df_headers"] = df_headers
#     st.session_state["df_items"] = dfi

# # ================== RENDER RESULTS FROM STATE ==================
# lang = st.session_state["lang"]
# df_headers = st.session_state.get("df_headers")
# df_items = st.session_state.get("df_items")

# if isinstance(df_headers, pd.DataFrame):
#     items_cnt = 0 if df_items is None or df_items.empty else len(df_items)
#     st.write(f"**{t('kpi_headers', lang)}:** {len(df_headers)}  |  **{t('kpi_items', lang)}:** {items_cnt}")

#     # ---- Headers with native Streamlit table + global filter
#     st.subheader(t("headers_title", lang))
#     hdr_filter = st.text_input(t("headers_filter", lang), value="", key="hdr_filter_input")
#     dfh_show = df_headers.copy()
#     if hdr_filter.strip():
#         mask = pd.Series([False] * len(dfh_show))
#         for c in dfh_show.columns:
#             mask |= dfh_show[c].astype(str).str.contains(hdr_filter, case=False, na=False)
#         dfh_show = dfh_show[mask]
#     st.dataframe(dfh_show, use_container_width=True, hide_index=True)

#     # ---- Items with native Streamlit features + MATCH/⚑ columns
#     st.subheader(t("items_title", lang))
#     if df_items is None or df_items.empty:
#         st.info(t("items_empty", lang))
#     else:
#         itm_filter = st.text_input(t("items_filter", lang), value="", key="itm_filter_input")
#         dfi_show = df_items.copy()

#         # global substring filter across all columns
#         if itm_filter.strip():
#             m = pd.Series([False] * len(dfi_show))
#             for c in dfi_show.columns:
#                 m |= dfi_show[c].astype(str).str.contains(itm_filter, case=False, na=False)
#             dfi_show = dfi_show[m]

#         # MATCH logic based on prefilter terms
#         def _terms(s: str):
#             return [x.strip().lower() for x in (s or "").split(",") if x.strip()]
#         terms_list = _terms(st.session_state.get("prefilter_text", ""))

#         if terms_list:
#             def row_has_term(row):
#                 for v in row.astype(str).fillna(""):
#                     lv = v.lower()
#                     for term in terms_list:
#                         if term in lv:
#                             return True
#                 return False
#             dfi_show["MATCH"] = dfi_show.apply(row_has_term, axis=1)
#             dfi_show["⚑"] = dfi_show["MATCH"].map(lambda x: "✅" if x else "")
#             cols = ["⚑"] + [c for c in dfi_show.columns if c != "⚑"]
#             dfi_show = dfi_show[cols]

#         st.dataframe(dfi_show, use_container_width=True, hide_index=True)

#     # ---- Export
#     bio = io.BytesIO()
#     with pd.ExcelWriter(bio, engine="openpyxl") as writer:
#         df_headers.to_excel(writer, sheet_name="headers", index=False)
#         if df_items is not None and not df_items.empty:
#             df_items.to_excel(writer, sheet_name="items", index=False)
#     bio.seek(0)
#     st.download_button(
#         t("export_btn", lang),
#         data=bio,
#         file_name="invoice_extract_selected_headers.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#     )


# invoice_extractor_streamlit_full.py
import io, re, shutil, tempfile
import streamlit as st
import pdfplumber
import pandas as pd

# ------- Optional OCR libs (safe import) -------
try:
    from pdf2image import convert_from_path
    import pytesseract
    OCR_LIBS = True
except Exception:
    OCR_LIBS = False

# ================= I18N (UI only; data NOT translated) =================
TRANSLATIONS = {
    "fr": {
        "app_title": "CHATAR — Extracteur de Factures",
        "app_caption": "Téléversez des PDF → Pré-filtre rapide → Extraction en-têtes + lignes.",
        "upload_label": "Téléverser un ou plusieurs PDF",
        "filter_label": "Filtre global (mots-clés séparés par des virgules). Vide = traiter tous",
        "extract_btn": "▶️ Extraire",
        "kpi_headers": "Lignes en-têtes",
        "kpi_items": "Lignes articles",
        "warn_upload": "Veuillez téléverser au moins un PDF.",
        "warn_no_match": "Aucun PDF ne correspond au filtre.",
        "headers_title": "En-têtes",
        "items_title": "Articles",
        "items_empty": "Aucune ligne d’articles détectée (tableau introuvable).",
        "export_btn": "⬇️ Exporter en Excel",
        "lang_label": "Langue de l’interface",
        "language_fr": "Français",
        "language_en": "English",
        "language_nl": "Nederlands",
        "language_ar": "العربية",
        "filter_info": "Insensible à la casse. « huile, sucre » ⇒ correspond si l’un apparaît.",
        "ocr_label": "Utiliser l’OCR (Tesseract) pour PDFs non recherchables",
        "ocr_langs": "Langues OCR (ex. fra+eng+nld+ara)",
        "ocr_hint": "Activez l’OCR si certains PDFs sont des scans.",
        "not_searchable_found": "PDF non recherchables (ignorés car OCR est OFF) :",
        "prefilter_summary": "Pré-filtre : {kept} gardés, {skipped} ignorés.",
        "headers_filter": "Filtrer les en-têtes (toutes colonnes)",
        "items_filter": "Filtrer les articles (toutes colonnes)",
        "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
    },
    "en": {
        "app_title": "CHATAR — Invoice Extractor",
        "app_caption": "Upload PDFs → Fast prefilter → Extract headers + items.",
        "upload_label": "Upload one or more PDFs",
        "filter_label": "Global filter (keywords, comma-separated). Empty = process all",
        "extract_btn": "▶️ Extract",
        "kpi_headers": "Header rows",
        "kpi_items": "Item rows",
        "warn_upload": "Please upload at least one PDF.",
        "warn_no_match": "No PDF matched your filter.",
        "headers_title": "Headers",
        "items_title": "Items",
        "items_empty": "No items detected (table body not found).",
        "export_btn": "⬇️ Export to Excel",
        "lang_label": "Interface language",
        "language_fr": "Français",
        "language_en": "English",
        "language_nl": "Nederlands",
        "language_ar": "العربية",
        "filter_info": "Case-insensitive. ‘oil, sugar’ ⇒ match if any appears.",
        "ocr_label": "Use OCR (Tesseract) for non-searchable PDFs",
        "ocr_langs": "OCR languages (e.g. fra+eng+nld+ara)",
        "ocr_hint": "Enable OCR if some PDFs are scans.",
        "not_searchable_found": "Non-searchable PDFs (skipped because OCR is OFF):",
        "prefilter_summary": "Prefilter: {kept} kept, {skipped} skipped.",
        "headers_filter": "Filter headers (all columns)",
        "items_filter": "Filter items (all columns)",
        "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
    },
    "nl": {
        "app_title": "CHATAR — Factuur Extractor",
        "app_caption": "Upload PDF’s → Snelle voorfilter → Koppen + regels extraheren.",
        "upload_label": "Upload één of meerdere PDF's",
        "filter_label": "Globaal filter (komma-gescheiden). Leeg = alles verwerken",
        "extract_btn": "▶️ Extractie",
        "kpi_headers": "Kopregels",
        "kpi_items": "Artikelregels",
        "warn_upload": "Upload minstens één PDF.",
        "warn_no_match": "Geen PDF komt overeen met het filter.",
        "headers_title": "Koppen",
        "items_title": "Artikelen",
        "items_empty": "Geen artikelen gedetecteerd (tabel niet gevonden).",
        "export_btn": "⬇️ Exporteren naar Excel",
        "lang_label": "Taal van de interface",
        "language_fr": "Français",
        "language_en": "English",
        "language_nl": "Nederlands",
        "language_ar": "العربية",
        "filter_info": "Niet-hoofdlettergevoelig. ‘olie, suiker’ ⇒ matcht als één voorkomt.",
        "ocr_label": "OCR gebruiken (Tesseract) voor niet-doorzoekbare PDF’s",
        "ocr_langs": "OCR-talen (bijv. fra+eng+nld+ara)",
        "ocr_hint": "Schakel OCR in als sommige PDF’s scans zijn.",
        "not_searchable_found": "Niet-doorzoekbare PDF’s (overgeslagen omdat OCR UIT staat):",
        "prefilter_summary": "Voorfilter: {kept} behouden, {skipped} overgeslagen.",
        "headers_filter": "Koppen filteren (alle kolommen)",
        "items_filter": "Artikelen filteren (alle kolommen)",
        "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
    },
    "ar": {
        "app_title": "شاطر — أداة استخراج الفواتير",
        "app_caption": "ارفع ملفات PDF → فلترة سريعة → استخراج العناوين والبنود.",
        "upload_label": "ارفع ملف/ملفات PDF",
        "filter_label": "مرشح عام (كلمات مفصولة بفواصل). فارغ = معالجة الجميع",
        "extract_btn": "▶️ استخراج",
        "kpi_headers": "صفوف العناوين",
        "kpi_items": "صفوف البنود",
        "warn_upload": "الرجاء رفع ملف PDF واحد على الأقل.",
        "warn_no_match": "لا يوجد ملف يطابق المرشح.",
        "headers_title": "العناوين",
        "items_title": "البنود",
        "items_empty": "لم يتم العثور على بنود (الجدول غير موجود).",
        "export_btn": "⬇️ تصدير إلى Excel",
        "lang_label": "لغة الواجهة",
        "language_fr": "Français",
        "language_en": "English",
        "language_nl": "Nederlands",
        "language_ar": "العربية",
        "filter_info": "مطابقة غير حساسة لحالة الأحرف. « زيت، سكر » ⇒ تطابق إذا وُجد أي منها.",
        "ocr_label": "استخدم OCR (Tesseract) للملفات غير القابلة للبحث",
        "ocr_langs": "لغات OCR (مثال: fra+eng+nld+ara)",
        "ocr_hint": "فعّل OCR إذا كانت بعض الملفات صورًا ممسوحة.",
        "not_searchable_found": "ملفات غير قابلة للبحث (تخطيناها لأن OCR مُعطّل):",
        "prefilter_summary": "الفلترة المسبقة: {kept} تم الاحتفاظ بها، {skipped} تم تجاوزها.",
        "headers_filter": "تصفية العناوين (كل الأعمدة)",
        "items_filter": "تصفية البنود (كل الأعمدة)",
        "basmala": "بِسْمِ اللهِ الرَّحْمٰنِ الرَّحِيمِ",
    },
}
def t(key: str, lang: str, **kwargs) -> str:
    s = TRANSLATIONS.get(lang, TRANSLATIONS["fr"]).get(key, key)
    return s.format(**kwargs) if kwargs else s

# ================== Page config ==================
st.set_page_config(page_title="CHATAR — Invoice Extractor", layout="wide")

# ---------- Session state init ----------
for k, v in {"df_headers": None, "df_items": None, "lang": "fr"}.items():
    st.session_state.setdefault(k, v)

# ===== App header (no sidebar toggle; no auto-collapse) =====
st.markdown(f"**{t('basmala', st.session_state['lang'])}**")
st.markdown(f"## {t('app_title', st.session_state['lang'])}")
st.caption(t("app_caption", st.session_state["lang"]))

# -------- Sidebar controls --------
with st.sidebar:
    lang = st.selectbox(
        TRANSLATIONS["fr"]["lang_label"],
        options=[
            ("fr", "🇫🇷 " + TRANSLATIONS["fr"]["language_fr"]),
            ("en", "🇬🇧 " + TRANSLATIONS["en"]["language_en"]),
            ("nl", "🇳🇱 " + TRANSLATIONS["nl"]["language_nl"]),
            ("ar", "🇲🇦 " + TRANSLATIONS["ar"]["language_ar"]),
        ],
        index=["fr","en","nl","ar"].index(st.session_state["lang"]),
        format_func=lambda opt: opt[1],
        key="lang_picker"
    )[0]
    st.session_state["lang"] = lang

    files = st.file_uploader(t("upload_label", lang), type="pdf", accept_multiple_files=True, key="uploader")
    filter_text = st.text_input(t("filter_label", lang), value="", key="prefilter_text")
    st.caption("ℹ️ " + t("filter_info", lang))

    use_ocr = st.checkbox(t("ocr_label", lang), value=False, key="use_ocr")
    ocr_langs = st.text_input(t("ocr_langs", lang), value="fra+eng+nld+ara", key="ocr_langs")
    st.caption("🧠 " + t("ocr_hint", lang))

    go = st.button(t("extract_btn", lang), use_container_width=True, key="extract_btn")

# ---------------- Helpers ----------------
def parse_terms(s: str):
    return [x.strip().lower() for x in (s or "").split(",") if x.strip()]

def pdf_text_allpages(file_bytes: bytes):
    texts = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for p in pdf.pages:
                txt = p.extract_text() or ""
                if txt:
                    texts.append(txt)
    except Exception:
        pass
    return "\n".join(texts)

def pdf_ocr_first_pages(file_bytes: bytes, langs="eng", pages=2):
    if not (OCR_LIBS and shutil.which("tesseract")):
        return ""
    try:
        imgs = convert_from_path(io.BytesIO(file_bytes), first_page=1, last_page=pages, dpi=180)
    except TypeError:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tmp:
            tmp.write(file_bytes); tmp.flush()
            imgs = convert_from_path(tmp.name, first_page=1, last_page=pages, dpi=180)
    except Exception:
        return ""
    out = []
    for im in imgs:
        try:
            out.append(pytesseract.image_to_string(im, lang=langs) or "")
        except Exception:
            continue
    return "\n".join(out)

# --- Template-specific parsing (Chakirs-like) ---
ITEMS_HEADER_RE = re.compile(
    r"#\s*DESCRIPTION\s+COLIS\s+QT[ÉE]\s+TOTAL\s+EN\s+U\.M\.\s+UNIT[ÉE]\s+P\.\s*U\.\s+TVA\s+TOTAL\s+HT",
    re.I
)
TVA_BLOCK_RE = re.compile(r"TVA\s*%\s*Base", re.I)

def norm_num2(s):
    if s is None: return None
    t = str(s).replace("\xa0"," ").strip().replace(" ","")
    if re.search(r"\d+\.\d{3},\d{2,4}$", t): t = t.replace(".","").replace(",",".")
    else: t = t.replace(",",".")
    try: return float(t)
    except Exception: return None

COUNTRIES = {
    "canada","france","belgium","belgique","belgië","netherlands","pays-bas","nederland",
    "morocco","maroc","algeria","algérie","tunisia","tunisie","spain","espagne","italy","italie",
    "germany","allemagne","luxembourg","suisse","switzerland","united kingdom","uk","angleterre",
    "usa","united states","etats-unis","états-unis"
}
def last_token_country(s: str) -> str:
    if not s: return ""
    tail = s.split("|")[-1].strip().lower()
    words = re.split(r"[,\s]+", tail)
    for w in reversed(words):
        if w in COUNTRIES:
            return w.title() if len(w) > 2 else w.upper()
    if tail in COUNTRIES: return tail.title()
    return ""

def chakirs_extract_headers(full_txt: str) -> dict:
    head = {}
    m = re.search(r"FACTURE\s*#\s*([A-Z0-9/\-]+)", full_txt, re.I)
    if m: head["Invoice_No"] = m.group(1).strip()
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+D[ûu]e\s+(\d{2}/\d{2}/\d{4})", full_txt, re.I)
    if m: head["Date"], head["Due_Date"] = m.group(1), m.group(2)

    def parse_block(full_txt, title_re, stops):
        mm = re.search(title_re, full_txt, re.I)
        if not mm: return "", ""
        start = mm.end(); end = len(full_txt)
        for sr in stops:
            ms = re.search(sr, full_txt[start:], re.I)
            if ms:
                end = start + ms.start(); break
        chunk = full_txt[start:end]
        lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]
        if not lines: return "", ""
        name = re.sub(r"\s+"," ",lines[0]).strip()
        addr = "\n".join(lines[1:]) if len(lines)>1 else ""
        return name, addr

    ship_name, ship_addr = parse_block(
        full_txt, r"ADRESSE DE LIVRAISON\s*",
        [r"#\s*DESCRIPTION", r"TVA\s*%\s*Base", r"N[ºo]\s*TVA"]
    )
    head["Shipping_Name"]    = ship_name
    head["Shipping_Address"] = ship_addr.strip()
    head["Shipping_Country"] = last_token_country(ship_addr)
    return head

def chakirs_parse_items(lines, invoice_no=None):
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
        L = re.sub(r"\s+", " ", lines[i]).strip()
        if not L or L.lower().startswith("sh "):
            i += 1; continue
        if TVA_BLOCK_RE.search(L): break
        msku = re.match(r"^([A-Z0-9]+)\s+(.*)$", L)
        if not msku:
            i += 1; continue
        sku = msku.group(1); right = msku.group(2)
        mtail = tail_re.search(right)
        if mtail:
            desc  = right[: mtail.start()].strip()
            pack  = (mtail.group("pack") or "").strip()
            qte   = norm_num2(mtail.group("qte"))
            totum = norm_num2(mtail.group("totum"))
            unit  = mtail.group("unit")
            pu    = norm_num2(mtail.group("pu"))
            tva   = mtail.group("tva")
            totht = norm_num2(mtail.group("totht"))
            items.append({
                "Invoice_No": invoice_no, "SKU": sku, "Description": desc, "Colis": pack,
                "Qte": qte, "Total_UM": totum, "Unite": unit, "PU": pu, "TVA_pct": tva, "Total_HT": totht
            })
        i += 1
    return items

def extract_one_pdf(name: str, raw_bytes: bytes, filter_text: str):
    contains = True
    all_pages = []
    try:
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                all_pages.append(t)
    except Exception:
        all_pages = []
    joined = "\n".join(all_pages)
    if filter_text:
        contains = filter_text.lower() in joined.lower()

    head = chakirs_extract_headers(joined)
    m = re.search(r"FACTURE\s*#\s*([A-Z0-9/\-]+)", joined, re.I)
    if m: head["Invoice_No"] = m.group(1).strip()
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+D[ûu]e\s+(\d{2}/\d{2}/\d{4})", joined, re.I)
    if m: head["Date"], head["Due_Date"] = m.group(1), m.group(2)
    head["filename"] = name
    head["contains_filter"] = contains

    lines = []
    for t in all_pages:
        lines.extend([ln for ln in t.splitlines() if ln and ln.strip()])
    items_df = pd.DataFrame()
    try:
        start_idx = next(i for i, L in enumerate(lines) if ITEMS_HEADER_RE.search(L))
    except StopIteration:
        start_idx = None
    if start_idx is not None:
        sub = []
        for L in lines[start_idx+1:]:
            if TVA_BLOCK_RE.search(L): break
            sub.append(L)
        rows = chakirs_parse_items(sub, invoice_no=head.get("Invoice_No"))
        if rows:
            items_df = pd.DataFrame(rows)
    return contains, pd.DataFrame([head]), items_df

# ================== EXTRACT (button) ==================
if go:
    if not files:
        st.warning(t("warn_upload", st.session_state["lang"]))
        st.stop()

    # Prefilter
    with st.spinner("Indexing PDFs..."):
        terms = parse_terms(st.session_state.get("prefilter_text", ""))
        non_searchable = []
        kept_files, skipped_files = [], []
        progress = st.progress(0); status = st.empty()

        for i, f in enumerate(files, start=1):
            fb = f.getbuffer()
            if not terms:
                kept_files.append(f)
            else:
                plain = pdf_text_allpages(fb)
                searchable = bool(plain.strip())
                found = False
                if searchable:
                    low = plain.lower()
                    found = any(term in low for term in terms)
                if not found and not searchable:
                    non_searchable.append(f.name)
                    if st.session_state.get("use_ocr") and OCR_LIBS and shutil.which("tesseract"):
                        ocr_txt = pdf_ocr_first_pages(fb, langs=st.session_state.get("ocr_langs") or "eng", pages=2)
                        low = (ocr_txt or "").lower()
                        found = any(term in low for term in terms)
                if found or not terms:
                    kept_files.append(f)
                else:
                    skipped_files.append(f)
            progress.progress(int(i*100/len(files))); status.write(f"Prefiltered {i}/{len(files)}")

    st.info(t("prefilter_summary", st.session_state["lang"], kept=len(kept_files), skipped=len(skipped_files)))
    if parse_terms(st.session_state.get("prefilter_text", "")) and non_searchable and not st.session_state.get("use_ocr"):
        st.warning("⚠️ " + t("not_searchable_found", st.session_state["lang"]) + " " + ", ".join(non_searchable))
    if parse_terms(st.session_state.get("prefilter_text", "")) and not kept_files:
        st.warning(t("warn_no_match", st.session_state["lang"]))
        st.session_state["df_headers"] = pd.DataFrame(); st.session_state["df_items"] = pd.DataFrame()
        st.stop()

    # Extract
    with st.spinner("Extracting data..."):
        all_headers, all_items = [], []
        progress2 = st.progress(0); status2 = st.empty()
        for i, f in enumerate(kept_files, start=1):
            _, dfh, dfi = extract_one_pdf(f.name, f.getbuffer(), st.session_state.get("prefilter_text", ""))
            all_headers.append(dfh)
            if not dfi.empty:
                dfi["filename"] = f.name
                all_items.append(dfi)
            progress2.progress(int(i*100/len(kept_files))); status2.write(f"Processed {i}/{len(kept_files)}")

    df_full = pd.concat(all_headers, ignore_index=True) if all_headers else pd.DataFrame()
    dfi = pd.concat(all_items, ignore_index=True) if all_items else pd.DataFrame()

    # Headers: only these columns
    keep_cols = ["Invoice_No", "Date", "Due_Date", "Shipping_Name", "Shipping_Country"]
    df_headers = df_full[[c for c in keep_cols if c in df_full.columns]].copy()

    st.session_state["df_headers"] = df_headers
    st.session_state["df_items"] = dfi

# ================== RENDER RESULTS FROM STATE ==================
lang = st.session_state["lang"]
df_headers = st.session_state.get("df_headers")
df_items = st.session_state.get("df_items")

if isinstance(df_headers, pd.DataFrame):
    items_cnt = 0 if df_items is None or df_items.empty else len(df_items)
    st.write(f"**{t('kpi_headers', lang)}:** {len(df_headers)}  |  **{t('kpi_items', lang)}:** {items_cnt}")

    # Headers table + global filter
    st.subheader(t("headers_title", lang))
    hdr_filter = st.text_input(t("headers_filter", lang), value="", key="hdr_filter_input")
    dfh_show = df_headers.copy()
    if hdr_filter.strip():
        mask = pd.Series([False]*len(dfh_show))
        for c in dfh_show.columns:
            mask |= dfh_show[c].astype(str).str.contains(hdr_filter, case=False, na=False)
        dfh_show = dfh_show[mask]

    st.dataframe(
        dfh_show,
        use_container_width=True,
        hide_index=True
    )

    # Items table + global filter + MATCH/⚑ column; keep Streamlit features
    st.subheader(t("items_title", lang))
    if df_items is None or df_items.empty:
        st.info(t("items_empty", lang))
    else:
        itm_filter = st.text_input(t("items_filter", lang), value="", key="itm_filter_input")
        dfi_show = df_items.copy()

        if itm_filter.strip():
            m = pd.Series([False]*len(dfi_show))
            for c in dfi_show.columns:
                m |= dfi_show[c].astype(str).str.contains(itm_filter, case=False, na=False)
            dfi_show = dfi_show[m]

        # Build MATCH flag from prefilter terms
        def _terms(s: str):
            return [x.strip().lower() for x in (s or "").split(",") if x.strip()]
        terms_list = _terms(st.session_state.get("prefilter_text", ""))

        if terms_list:
            def row_has_term(row):
                for v in row.astype(str).fillna(""):
                    lv = v.lower()
                    for term in terms_list:
                        if term in lv:
                            return True
                return False
            dfi_show["MATCH"] = dfi_show.apply(row_has_term, axis=1)
            dfi_show["⚑"] = dfi_show["MATCH"].map(lambda x: "✅" if x else "")
            ordered_cols = ["⚑"] + [c for c in dfi_show.columns if c != "⚑"]
            dfi_show = dfi_show[ordered_cols]

        # Nice number formatting
        column_cfg = {
            "Qte": st.column_config.NumberColumn(format="%,d"),
            "PU": st.column_config.NumberColumn(format="%.2f"),
            "Total_HT": st.column_config.NumberColumn(format="%.2f"),
            "Total_UM": st.column_config.NumberColumn(format="%.0f"),
        }

        st.dataframe(
            dfi_show,
            use_container_width=True,
            hide_index=True,
            column_config=column_cfg
        )

    # Export both sheets
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_headers.to_excel(writer, sheet_name="headers", index=False)
        if df_items is not None and not df_items.empty:
            df_items.to_excel(writer, sheet_name="items", index=False)
    bio.seek(0)
    st.download_button(
        t("export_btn", lang),
        data=bio,
        file_name="invoice_extract.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )