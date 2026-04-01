import streamlit as st
import pandas as pd
import re
import io
import random
from faker import Faker
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import copy

# ─────────────────────────────────────────────
#  Setup
# ─────────────────────────────────────────────
st.set_page_config(page_title="DataMask", page_icon="🔒", layout="wide")

fake_en = Faker('en_GB')
fake_sv = Faker('sv_SE')

# ─────────────────────────────────────────────
#  UI TRANSLATIONS
# ─────────────────────────────────────────────
T = {
    "en": {
        "title": "🔒 Veilix — Excel Sensitive Data Cleaner",
        "subtitle": "Upload an Excel or CSV file. We'll scan it, mask the sensitive data, and give you a clean file to download.",
        "settings": "⚙️ Settings",
        "what_to_mask": "What to Mask",
        "masking_style": "Masking Style",
        "mask_mode_label": "How should we replace sensitive data?",
        "mask_modes": ["Asterisks (*****)", "Fake Realistic Data", "Token (e.g. EMAIL_001)"],
        "tip": "💡 'Fake Realistic Data' keeps your file usable for testing while removing all real personal info.",
        "upload_label": "📂 Upload your file (.xlsx or .csv)",
        "upload_help": "Your file never leaves your browser session. Nothing is stored.",
        "file_loaded": "✅ File loaded: **{name}** — {rows} rows, {cols} columns",
        "choose_cols": "📋 Choose which columns to scan",
        "cols_caption": "By default all columns are scanned. Untick any you want to leave untouched.",
        "cols_label": "Columns to scan:",
        "preview_label": "👀 Preview Original Data (first 10 rows)",
        "run_button": "🚀 Run Masking",
        "warn_no_patterns": "Please select at least one pattern to mask.",
        "spinner": "Scanning and masking your data...",
        "done": "✅ Done! **{n}** values were masked across your file.",
        "original_sample": "🔴 Original (sample)",
        "masked_sample": "🟢 Masked (sample)",
        "report_expander": "📊 View Masking Report ({n} changes)",
        "download_header": "📥 Download Your Clean File",
        "download_button": "⬇️ Download Masked Excel File",
        "download_caption": "The downloaded file contains 3 tabs: Original Data, Masked Data, and a Masking Report.",
        "landing_info": "👈 Upload a file to get started. Use the sidebar to configure what gets masked.",
        "landing_table_header": "### What does this tool detect?",
        "landing_col_type": "Type",
        "landing_col_example": "Example",
        "landing_col_masked": "Masked As",
        "file_error": "Couldn't read that file. Error: {e}",    
    },
    "sv": {
        "title": "🔒 Velix — Rensa känslig data i Excel",
        "subtitle": "Ladda upp en Excel- eller CSV-fil. Vi skannar den, maskerar känsliga uppgifter och ger dig en ren fil att ladda ner.",
        "settings": "⚙️ Inställningar",
        "what_to_mask": "Vad ska maskeras",
        "masking_style": "Maskeringsstil",
        "mask_mode_label": "Hur ska vi ersätta känsliga uppgifter?",
        "mask_modes": ["Asterisker (*****)", "Falsk realistisk data", "Token (t.ex. EMAIL_001)"],
        "tip": "💡 'Falsk realistisk data' håller filen användbar för testning och tar bort all riktig personlig information.",
        "upload_label": "📂 Ladda upp din fil (.xlsx eller .csv)",
        "upload_help": "Din fil lämnar aldrig din webbläsarsession. Ingenting lagras.",
        "file_loaded": "✅ Fil inläst: **{name}** — {rows} rader, {cols} kolumner",
        "choose_cols": "📋 Välj vilka kolumner som ska skannas",
        "cols_caption": "Som standard skannas alla kolumner. Avmarkera de du vill lämna orörda.",
        "cols_label": "Kolumner att skanna:",
        "preview_label": "👀 Förhandsgranska originaldata (första 10 raderna)",
        "run_button": "🚀 Kör maskering",
        "warn_no_patterns": "Välj minst ett mönster att maskera.",
        "spinner": "Skannar och maskerar din data...",
        "done": "✅ Klart! **{n}** värden maskerades i din fil.",
        "original_sample": "🔴 Original (urval)",
        "masked_sample": "🟢 Maskerad (urval)",
        "report_expander": "📊 Visa maskeringsrapport ({n} ändringar)",
        "download_header": "📥 Ladda ner din rena fil",
        "download_button": "⬇️ Ladda ner maskerad Excel-fil",
        "download_caption": "Den nedladdade filen innehåller 3 flikar: Originaldata, Maskerad data och Maskeringsrapport.",
        "landing_info": "👈 Ladda upp en fil för att komma igång. Använd sidopanelen för att konfigurera vad som maskeras.",
        "landing_table_header": "### Vad identifierar det här verktyget?",
        "landing_col_type": "Typ",
        "landing_col_example": "Exempel",
        "landing_col_masked": "Maskeras som",
        "file_error": "Kunde inte läsa filen. Fel: {e}",
    },
}

# ─────────────────────────────────────────────
#  REGEX PATTERNS — what we scan for
# ─────────────────────────────────────────────
PATTERNS = [
    ("Email Addresses",       r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'),
    ("Swedish Personnummer",  r'\b(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])[-+]?\d{4}\b'),
    ("Credit Card Numbers",   r'\b(?:4\d{3}|5[1-5]\d{2}|3[47]\d{2}|6(?:011|5\d{2}))[\s-]?\d{4}[\s-]?\d{4}[\s-]?\d{3,4}\b'),
    ("National Insurance",    r'\b[A-Z]{2}\s?\d{2}\s?\d{2}\s?\d{2}\s?[A-Z]\b'),
    ("UK Phone Numbers",      r'(\+44\s?|0)7\d{3}[\s-]?\d{6}'),
    ("US Phone Numbers",      r'\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}'),
    ("Swedish Phone Numbers", r'(\+46\s?|0)7[02369]\s?-?\s?\d{3}\s?\d{2}\s?\d{2}'),
    ("Dates of Birth",        r'\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b'),
    ("IP Addresses",          r'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'),
    ("UK Postcodes",          r'\b[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}\b'),
    ("Swedish Postcodes",     r'\b\d{3}\s\d{2}\b'),
    ("Salary / Currency",     r'£[\d,]+|€[\d,]+|\$[\d,]+'),
    ("SEK Currency",          r'\b\d[\d\s]{2,}\s*kr\b'),
]

# Quick lookup set for categories (used in sidebar)
PATTERN_NAMES = [p[0] for p in PATTERNS]

# Bilingual display names for each pattern
PATTERN_LABELS = {
    "en": {
        "Email Addresses":       "Email Addresses",
        "UK Phone Numbers":      "UK Phone Numbers",
        "US Phone Numbers":      "US Phone Numbers",
        "UK Postcodes":          "UK Postcodes",
        "National Insurance":    "National Insurance (NI)",
        "Swedish Personnummer":  "Swedish Personnummer",
        "Swedish Phone Numbers": "Swedish Phone Numbers",
        "Swedish Postcodes":     "Swedish Postcodes",
        "Credit Card Numbers":   "Credit Card Numbers",
        "Dates of Birth":        "Dates of Birth",
        "IP Addresses":          "IP Addresses",
        "Salary / Currency":     "Salary / Currency (£€$)",
        "SEK Currency":          "SEK Currency (kr)",
    },
    "sv": {
        "Email Addresses":       "E-postadresser",
        "UK Phone Numbers":      "Brittiska telefonnummer",
        "US Phone Numbers":      "Amerikanska telefonnummer",
        "UK Postcodes":          "Brittiska postnummer",
        "National Insurance":    "NI-nummer (UK)",
        "Swedish Personnummer":  "Personnummer",
        "Swedish Phone Numbers": "Svenska mobilnummer",
        "Swedish Postcodes":     "Svenska postnummer",
        "Credit Card Numbers":   "Kreditkortsnummer",
        "Dates of Birth":        "Födelsedatum",
        "IP Addresses":          "IP-adresser",
        "Salary / Currency":     "Lön / Valuta (£€$)",
        "SEK Currency":          "SEK-belopp (kr)",
    },
}

# ─────────────────────────────────────────────
#  FAKE DATA REPLACEMENTS
# ─────────────────────────────────────────────
# ─────────────────────────────────────────────
#  FAKE DATA REPLACEMENTS
#  — uses a cache dict so the same original
#    value always maps to the same fake value
# ─────────────────────────────────────────────
def build_fake_generator(lang):
    """Returns a function that generates fake data with a
    per-session cache so identical originals → identical fakes."""
    fake = fake_sv if lang == "sv" else fake_en
    cache = {}  # {(category, original_text): fake_replacement}
 
    def fake_uk_phone():
        prefix = f"07{random.randint(100, 999)}"
        suffix = f"{random.randint(100000, 999999)}"
        return f"{prefix} {suffix}"
 
    def fake_us_phone():
        return f"({random.randint(200,999)}) {random.randint(200,999)}-{random.randint(1000,9999)}"
 
    def fake_personnummer():
        dob = fake_sv.date_of_birth(minimum_age=18, maximum_age=80)
        last4 = f"{random.randint(1000, 9999)}"
        return dob.strftime('%Y%m%d') + '-' + last4
 
    def fake_swedish_phone():
        prefixes = ['070', '072', '073', '076', '079']
        prefix = random.choice(prefixes)
        number = f"{random.randint(100, 999)} {random.randint(10, 99)} {random.randint(10, 99)}"
        return f"{prefix}-{number}"
 
    def fake_swedish_postcode():
        return f"{random.randint(100, 999)} {random.randint(10, 99):02d}"
 
    def fake_sek():
        amount = random.randint(10000, 120000)
        formatted = f"{amount:,}".replace(",", " ")
        return f"{formatted} kr"
 
    def fake_credit_card():
        prefix = random.choice(['4', '5', '37', '6011'])
        remaining = 16 - len(prefix)
        digits = prefix + ''.join([str(random.randint(0, 9)) for _ in range(remaining)])
        return ' '.join([digits[i:i+4] for i in range(0, len(digits), 4)])
 
    generators = {
        "Email Addresses":       fake.email,
        "UK Phone Numbers":      fake_uk_phone,
        "US Phone Numbers":      fake_us_phone,
        "UK Postcodes":          fake_en.postcode,
        "National Insurance":    lambda: f"{fake_en.lexify('??').upper()}{fake_en.numerify('######')}{fake_en.lexify('?').upper()}",
        "Swedish Personnummer":  fake_personnummer,
        "Swedish Phone Numbers": fake_swedish_phone,
        "Swedish Postcodes":     fake_swedish_postcode,
        "Credit Card Numbers":   fake_credit_card,
        "Dates of Birth":        lambda: fake.date_of_birth(minimum_age=18, maximum_age=80).strftime('%d/%m/%Y'),
        "IP Addresses":          fake.ipv4,
        "Salary / Currency":     lambda: f"£{fake_en.random_int(20000, 120000):,}",
        "SEK Currency":          fake_sek,
    }
 
    def get_fake(category, original_text):
        key = (category, original_text)
        if key not in cache:
            fn = generators.get(category)
            cache[key] = fn() if fn else "****"
        return cache[key]
 
    return get_fake

# ─────────────────────────────────────────────
#  CORE MASKING LOGIC — collect-then-replace
#  Finds all matches first, resolves overlaps,
#  then replaces in reverse order so spans
#  stay valid.
# ─────────────────────────────────────────────
def mask_cell(value, selected_patterns, mask_mode, token_counters, get_fake):
    """Mask a single cell value. Uses span-based replacement to
    avoid double-matching problems."""
    text = str(value)
 
    # Step 1: Collect all matches with their spans
    all_matches = []  # list of (start, end, category, matched_text)
    for category, pattern in PATTERNS:
        if category not in selected_patterns:
            continue
        for m in re.finditer(pattern, text):
            all_matches.append((m.start(), m.end(), category, m.group()))
 
    if not all_matches:
        return text
 
    # Step 2: Sort by start position, then longest match first (to prefer specific)
    all_matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))
 
    # Step 3: Remove overlapping matches — keep the first (most specific due to pattern order)
    kept = []
    last_end = -1
    for start, end, category, matched in all_matches:
        if start >= last_end:
            kept.append((start, end, category, matched))
            last_end = end
 
    # Step 4: Replace in reverse order so string indices stay correct
    for start, end, category, matched in reversed(kept):
        if mask_mode in ("Asterisks (*****)", "Asterisker (*****)"):
            replacement = '*' * len(matched)
        elif mask_mode in ("Fake Realistic Data", "Falsk realistisk data"):
            replacement = get_fake(category, matched)
        elif mask_mode in ("Token (e.g. EMAIL_001)", "Token (t.ex. EMAIL_001)"):
            key = category.replace(" ", "_").upper()
            token_counters[key] = token_counters.get(key, 0) + 1
            replacement = f"{key}_{token_counters[key]:03d}"
        else:
            replacement = '*' * len(matched)
 
        text = text[:start] + replacement + text[end:]
 
    return text
 
 
def process_dataframe(df, selected_patterns, mask_mode, selected_columns, lang, progress_bar=None):
    """Applies masking to a dataframe. Returns masked df + a report.
    Optionally updates a Streamlit progress bar."""
    masked_df = df.copy()
    token_counters = {}
    report = []
    get_fake = build_fake_generator(lang)
 
    cols_to_process = selected_columns if selected_columns else df.columns.tolist()
 
    # Calculate total work units for progress bar
    total_cells = sum(len(df[col].dropna()) for col in cols_to_process if col in df.columns)
    processed = 0
 
    for col in cols_to_process:
        if col not in df.columns:
            continue
        for idx, value in df[col].items():
            if pd.isna(value):
                continue
            original = str(value)
            cleaned = mask_cell(original, selected_patterns, mask_mode, token_counters, get_fake)
            if cleaned != original:
                masked_df.at[idx, col] = cleaned
                report.append({
                    "Row": idx + 2,
                    "Column": col,
                    "Original": original,
                    "Masked As": cleaned,
                })
            processed += 1
            if progress_bar and total_cells > 0:
                progress_bar.progress(
                    min(processed / total_cells, 1.0),
                    text=f"Processing column **{col}** — {processed:,}/{total_cells:,} cells"
                )
 
    if progress_bar:
        progress_bar.empty()
 
    return masked_df, pd.DataFrame(report)
 
 
# ─────────────────────────────────────────────
#  EXPORT TO EXCEL (with three tabs)
# ─────────────────────────────────────────────
def export_to_excel(original_df, masked_df, report_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        original_df.to_excel(writer, sheet_name='Original Data', index=False)
        masked_df.to_excel(writer, sheet_name='Masked Data', index=False)
        if not report_df.empty:
            report_df.to_excel(writer, sheet_name='Masking Report', index=False)
 
        wb = writer.book
        ws = wb['Masked Data']
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        for cell in ws[1]:
            cell.fill = green_fill
            cell.font = Font(bold=True)
 
    output.seek(0)
    return output
 
 
def export_to_csv(masked_df):
    """Export masked data as CSV."""
    return masked_df.to_csv(index=False).encode('utf-8')
 
 
# ─────────────────────────────────────────────
#  UI — STREAMLIT INTERFACE
# ─────────────────────────────────────────────
 
# ── Language toggle in main header area ──
header_left, header_right = st.columns([4, 1])
with header_right:
    lang = st.selectbox(
        "🌐",
        options=["en", "sv"],
        format_func=lambda x: "🇬🇧 English" if x == "en" else "🇸🇪 Svenska",
        label_visibility="collapsed",
    )
 
ui = T[lang]
labels = PATTERN_LABELS[lang]
 
with header_left:
    st.title(ui["title"])
 
st.markdown(ui["subtitle"])
st.divider()
 
# ── Sidebar: settings ──
with st.sidebar:
    st.header(ui["settings"])
 
    st.subheader(ui["what_to_mask"])
    selected_patterns = []
    for category in PATTERN_NAMES:
        if st.checkbox(labels[category], value=True, key=f"chk_{category}"):
            selected_patterns.append(category)
 
    st.subheader(ui["masking_style"])
    mask_mode = st.radio(
        ui["mask_mode_label"],
        ui["mask_modes"],
        index=0,
    )
 
    st.divider()
    st.caption(ui["tip"])
 
# --- File Upload ---
uploaded_file = st.file_uploader(
    ui["upload_label"],
    type=["xlsx", "csv"],
    help=ui["upload_help"],
)
 
if uploaded_file:
    # Store df in session_state so re-runs don't reload
    file_key = f"df_{uploaded_file.name}_{uploaded_file.size}"
    if file_key not in st.session_state:
        try:
            if uploaded_file.name.endswith(".csv"):
                st.session_state[file_key] = pd.read_csv(uploaded_file)
            else:
                st.session_state[file_key] = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(ui["file_error"].format(e=e))
            st.stop()
 
    df = st.session_state[file_key]
 
    st.success(ui["file_loaded"].format(name=uploaded_file.name, rows=len(df), cols=len(df.columns)))
    st.divider()
 
    st.subheader(ui["choose_cols"])
    st.caption(ui["cols_caption"])
    all_cols = df.columns.tolist()
    selected_columns = st.multiselect(
        ui["cols_label"],
        options=all_cols,
        default=all_cols,
    )
 
    with st.expander(ui["preview_label"]):
        st.dataframe(df.head(10), use_container_width=True)
 
    st.divider()
 
    if st.button(ui["run_button"], type="primary", use_container_width=True):
        if not selected_patterns:
            st.warning(ui["warn_no_patterns"])
        else:
            progress_bar = st.progress(0, text="Starting…")
 
            masked_df, report_df = process_dataframe(
                df, selected_patterns, mask_mode, selected_columns, lang,
                progress_bar=progress_bar,
            )
 
            # Store results in session state
            st.session_state['masked_df'] = masked_df
            st.session_state['report_df'] = report_df
 
    # ── Show results if they exist ──
    if 'masked_df' in st.session_state and 'report_df' in st.session_state:
        masked_df = st.session_state['masked_df']
        report_df = st.session_state['report_df']
 
        st.success(ui["done"].format(n=len(report_df)))
        st.divider()
 
        if not report_df.empty:
            # Show only the rows that actually changed
            changed_row_indices = report_df["Row"].unique() - 2  # convert back to 0-indexed
            changed_row_indices = sorted(set(changed_row_indices) & set(df.index))
            sample_indices = changed_row_indices[:15]  # show up to 15 changed rows
 
            col1, col2 = st.columns(2)
            with col1:
                st.subheader(ui["original_sample"])
                st.dataframe(df.loc[sample_indices], use_container_width=True)
            with col2:
                st.subheader(ui["masked_sample"])
                st.dataframe(masked_df.loc[sample_indices], use_container_width=True)
 
            with st.expander(ui["report_expander"].format(n=len(report_df))):
                st.dataframe(report_df, use_container_width=True)
        else:
            st.info(ui["no_changes"])
 
        st.divider()
 
        # ── Download section with both Excel and CSV ──
        st.subheader(ui["download_header"])
        original_name = uploaded_file.name.replace('.xlsx', '').replace('.csv', '')
 
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            excel_output = export_to_excel(df, masked_df, report_df)
            st.download_button(
                label=ui["download_excel_button"],
                data=excel_output,
                file_name=f"{original_name}_MASKED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        with dl_col2:
            csv_output = export_to_csv(masked_df)
            st.download_button(
                label=ui["download_csv_button"],
                data=csv_output,
                file_name=f"{original_name}_MASKED.csv",
                mime="text/csv",
                use_container_width=True,
            )
        st.caption(ui["download_caption"])
 
else:
    st.info(ui["landing_info"])
    st.markdown(ui["landing_table_header"])
 
    # Dynamic landing table that reflects the selected masking mode
    EXAMPLE_DATA = [
        ("Email Addresses",       "Email",            "E-post",                   "john@company.com",        "kalle@foretag.se"),
        ("UK Phone Numbers",      "Phone (UK)",       "Brittiskt telefonnr",      "07891 234567",            "07891 234567"),
        ("Swedish Phone Numbers", "Phone (SWE)",      "Svenskt mobilnr",         "070-123 45 67",           "070-123 45 67"),
        ("UK Postcodes",          "Postcode (UK)",    "Brittiskt postnr",         "SW1A 1AA",                "SW1A 1AA"),
        ("Swedish Postcodes",     "Postcode (SWE)",   "Svenskt postnr",          "113 45",                  "113 45"),
        ("National Insurance",    "NI Number",        "NI-nummer (UK)",           "AB123456C",               "AB123456C"),
        ("Swedish Personnummer",  "Personnummer",     "Personnummer",             "19850312-1234",           "19850312-1234"),
        ("Salary / Currency",     "Salary",           "Lön",                      "£45,000",                 "£45,000"),
        ("SEK Currency",          "SEK Amount",       "SEK-belopp",              "45 000 kr",               "45 000 kr"),
        ("Credit Card Numbers",   "Credit Card",      "Kreditkort",              "4111 1111 1111 1111",     "4111 1111 1111 1111"),
        ("Dates of Birth",        "Date of Birth",    "Födelsedatum",            "12/05/1987",              "12/05/1987"),
        ("IP Addresses",          "IP Address",       "IP-adress",               "192.168.1.1",             "192.168.1.1"),
    ]
 
    get_fake_landing = build_fake_generator(lang)
    token_counters_landing = {}
 
    col_type = ui["landing_col_type"]
    col_ex   = ui["landing_col_example"]
    col_mask = ui["landing_col_masked"]
 
    table = f"| {col_type} | {col_ex} | {col_mask} |\n|---|---|---|\n"
    for category, label_en, label_sv, example_en, example_sv in EXAMPLE_DATA:
        label   = label_sv if lang == "sv" else label_en
        example = example_sv if lang == "sv" else example_en
 
        if mask_mode in ("Asterisks (*****)", "Asterisker (*****)"):
            masked = "*" * len(example)
        elif mask_mode in ("Fake Realistic Data", "Falsk realistisk data"):
            masked = get_fake_landing(category, example)
        elif mask_mode in ("Token (e.g. EMAIL_001)", "Token (t.ex. EMAIL_001)"):
            key = category.replace(" ", "_").upper()
            token_counters_landing[key] = token_counters_landing.get(key, 0) + 1
            masked = f"{key}_{token_counters_landing[key]:03d}"
        else:
            masked = "*" * len(example)
 
        table += f"| {label} | `{example}` | `{masked}` |\n"
 
    st.markdown(table)
