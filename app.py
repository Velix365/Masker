
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
st.set_page_config(
    page_title="DataMask | Professional Data Anonymization",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Professional CSS Styling
st.markdown("""
<style>
    /* Main color scheme - Professional blues and greens */
    :root {
        --primary-color: #0066CC;
        --secondary-color: #00AA88;
        --background-color: #F8F9FA;
        --text-color: #2C3E50;
        --border-color: #E1E8ED;
    }

    /* Header styling */
    h1 {
        color: var(--primary-color) !important;
        font-weight: 700 !important;
        letter-spacing: -0.5px !important;
        margin-bottom: 0.5rem !important;
    }

    h2, h3 {
        color: var(--text-color) !important;
        font-weight: 600 !important;
    }

    /* Subtitle styling */
    .subtitle {
        font-size: 1.1rem;
        color: #5A6C7D;
        margin-bottom: 2rem;
    }

    /* Security banner */
    .security-banner {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1.5rem 0;
        text-align: center;
        font-weight: 500;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }

    /* Feature boxes */
    .feature-box {
        background: white;
        border: 2px solid var(--border-color);
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }

    .feature-box:hover {
        border-color: var(--primary-color);
        box-shadow: 0 4px 12px rgba(0,102,204,0.15);
        transform: translateY(-2px);
    }

    /* Buttons */
    .stButton > button {
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 0.5rem 2rem !important;
        transition: all 0.3s ease !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
    }

    /* File uploader */
    .stFileUploader {
        background: white;
        border: 2px dashed var(--border-color);
        border-radius: 10px;
        padding: 2rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #F8F9FA;
    }

    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: var(--primary-color) !important;
    }

    /* Success/Info boxes */
    .stSuccess, .stInfo {
        border-radius: 8px !important;
    }

    /* Dataframe styling */
    .dataframe {
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
    }

    /* Trust badges */
    .trust-badge {
        display: inline-block;
        background: white;
        border: 1px solid var(--border-color);
        border-radius: 20px;
        padding: 0.4rem 1rem;
        margin: 0.3rem;
        font-size: 0.9rem;
        font-weight: 500;
        color: var(--text-color);
    }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Custom footer */
    .custom-footer {
        text-align: center;
        padding: 2rem 0;
        color: #7F8C8D;
        font-size: 0.9rem;
        border-top: 1px solid var(--border-color);
        margin-top: 3rem;
    }
</style>
""", unsafe_allow_html=True)

fake_en = Faker('en_GB')
fake_sv = Faker('sv_SE')

# ─────────────────────────────────────────────
#  NLP — lazy-loaded spaCy models
# ─────────────────────────────────────────────
@st.cache_resource
def load_nlp_models():
    try:
        import spacy
        try:
            nlp_en = spacy.load("en_core_web_sm")
        except OSError:
            # Try to download if not found
            import subprocess
            subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"], check=True)
            nlp_en = spacy.load("en_core_web_sm")

        try:
            nlp_sv = spacy.load("sv_core_news_sm")
        except OSError:
            # Try to download if not found
            import subprocess
            subprocess.run(["python", "-m", "spacy", "download", "sv_core_news_sm"], check=True)
            nlp_sv = spacy.load("sv_core_news_sm")

        return nlp_en, nlp_sv
    except Exception as e:
        st.warning(f"Error loading NLP models: {e}")
        return None, None

def detect_lang(text):
    """Returns 'sv' or 'en' based on text content. Falls back to 'en'."""
    try:
        from langdetect import detect
        result = detect(text)
        return "sv" if result == "sv" else "en"
    except Exception:
        return "en"

def mask_names_ner(text, mask_mode, token_counters, lang, nlp_en, nlp_sv):
    """
    Uses spaCy NER to find PERSON entities and mask them.
    Tries both English and Swedish models to catch all names.
    """
    if not nlp_en or not nlp_sv:
        return text  # NLP models not available

    if len(text.strip()) < 2:
        return text  # too short

    # Try both models and combine results
    persons_found = set()

    for nlp in [nlp_en, nlp_sv]:
        doc = nlp(text)
        for ent in doc.ents:
            if ent.label_ in ("PER", "PERSON"):
                persons_found.add((ent.start_char, ent.end_char, ent.text))

    if not persons_found:
        return text

    # Sort by position (reverse order for replacement)
    persons = sorted(persons_found, key=lambda x: x[0], reverse=True)

    # Replace from end to start so character offsets stay valid
    result = text
    for start, end, name in persons:
        if mask_mode in ("Asterisks (*****)", "Asterisker (*****)"):
            replacement = '*' * len(name)
        elif mask_mode in ("Fake Realistic Data", "Falsk realistisk data"):
            # Use Swedish faker for Swedish-sounding names, English otherwise
            fake = fake_sv if any(c in name.lower() for c in ['å', 'ä', 'ö']) else fake_en
            replacement = fake.name()
        else:  # Token
            token_counters["PERSON"] = token_counters.get("PERSON", 0) + 1
            replacement = f"PERSON_{token_counters['PERSON']:03d}"
        result = result[:start] + replacement + result[end:]

    return result

# ─────────────────────────────────────────────
#  UI TRANSLATIONS
# ─────────────────────────────────────────────
T = {
    "en": {
        "title": "🔒 DataMask — Excel Sensitive Data Cleaner",
        "subtitle": "Upload an Excel or CSV file. We'll scan it, mask the sensitive data, and give you a clean file to download.",
        "settings": "⚙️ Settings",
        "what_to_mask": "What to Mask",
        "masking_style": "Masking Style",
        "mask_mode_label": "How should we replace sensitive data?",
        "mask_modes": ["Asterisks (*****)", "Fake Realistic Data", "Token (e.g. EMAIL_001)"],
        "tip": "💡 'Fake Realistic Data' keeps your file usable for testing while removing all real personal info.\n\n✨ **Person names are automatically detected** using AI (spaCy NER) in both English and Swedish.",
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
        "landing_table_rows": [
            ("Person Name (NLP) ✨", "Erik Johansson", "James Smith",        "both"),
            ("Email",                "john@company.com","fake@email.com",    "both"),
            ("UK Phone",             "07891 234567",    "07234 567890",      "en"),
            ("UK Postcode",          "SW1A 1AA",        "M4 3AB",            "en"),
            ("NI Number",            "AB123456C",       "XY654321D",         "en"),
            ("Salary (£€$)",         "£45,000",         "£67,000",           "en"),
            ("Personnummer",         "19850312-1234",   "19920814-5678",     "sv"),
            ("Swedish Phone",        "070-123 45 67",   "073-456 78 90",     "sv"),
            ("Swedish Postcode",     "113 45",          "211 56",            "sv"),
            ("SEK Amount",           "45 000 kr",       "67 000 kr",         "sv"),
            ("Credit Card",          "4111 1111 1111 1111","5412 7534 2341 9876","both"),
            ("Date of Birth",        "12/05/1987",      "23/08/1994",        "both"),
            ("IP Address",           "192.168.1.1",     "83.21.45.7",        "both"),
        ],
        "file_error": "Couldn't read that file. Error: {e}",
        "landing_col_type": "Type",
        "landing_col_example": "Example",
        "landing_col_masked": "Masked",
        "no_changes": "No sensitive data was found in the selected columns.",
        "download_excel_button": "⬇️ Download Excel (.xlsx)",
        "download_csv_button": "⬇️ Download CSV",
    },
    "sv": {
        "title": "🔒 DataMask — Rensa känslig data i Excel",
        "subtitle": "Ladda upp en Excel- eller CSV-fil. Vi skannar den, maskerar känsliga uppgifter och ger dig en ren fil att ladda ner.",
        "settings": "⚙️ Inställningar",
        "what_to_mask": "Vad ska maskeras",
        "masking_style": "Maskeringsstil",
        "mask_mode_label": "Hur ska vi ersätta känsliga uppgifter?",
        "mask_modes": ["Asterisker (*****)", "Falsk realistisk data", "Token (t.ex. EMAIL_001)"],
        "tip": "💡 'Falsk realistisk data' håller filen användbar för testning och tar bort all riktig personlig information.\n\n✨ **Personnamn identifieras automatiskt** med AI (spaCy NER) på både engelska och svenska.",
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
        "landing_table_rows": [
            ("Personnamn (NLP) ✨",  "Erik Johansson",          "Lars Svensson",             "both"),
            ("E-post",               "kalle@foretag.se",        "falsk@epost.se",             "both"),
            ("Personnummer",         "19850312-1234",           "19920814-5678",              "sv"),
            ("Svenskt mobilnummer",  "070-123 45 67",           "073-456 78 90",              "sv"),
            ("Svenskt postnummer",   "113 45",                  "211 56",                     "sv"),
            ("SEK-belopp",           "45 000 kr",               "67 000 kr",                  "sv"),
            ("Brittiskt telefon",    "07891 234567",            "07234 567890",               "en"),
            ("Brittiskt postnummer", "SW1A 1AA",                "M4 3AB",                     "en"),
            ("NI-nummer (UK)",       "AB123456C",               "XY654321D",                  "en"),
            ("Lön (£€$)",            "£45,000",                 "£67,000",                    "en"),
            ("Kreditkort",           "4111 1111 1111 1111",     "5412 7534 2341 9876",        "both"),
            ("Födelsedatum",         "12/05/1987",              "23/08/1994",                 "both"),
            ("IP-adress",            "192.168.1.1",             "83.21.45.7",                 "both"),
        ],
        "file_error": "Kunde inte läsa filen. Fel: {e}",
        "landing_col_type": "Typ",
        "landing_col_example": "Exempel",
        "landing_col_masked": "Maskerad",
        "no_changes": "Ingen känslig data hittades i de valda kolumnerna.",
        "download_excel_button": "⬇️ Ladda ner Excel (.xlsx)",
        "download_csv_button": "⬇️ Ladda ner CSV",
    },
}

# ─────────────────────────────────────────────
#  REGEX PATTERNS — what we scan for
# ─────────────────────────────────────────────
PATTERNS = {
    "Person Names (NLP)":     None,  # Special case - handled by NER, not regex
    "Email Addresses":        r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+',
    "UK Phone Numbers":       r'(?:\+44\s?|0)7\d{3}[\s-]?\d{3}[\s-]?\d{3}|(?:\+44\s?|0)7\d{3}[\s-]?\d{6}',
    "US Phone Numbers":       r'\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}',
    "UK Postcodes":           r'[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}',
    "National Insurance":     r'[A-Z]{2}\s?\d{2}\s?\d{2}\s?\d{2}\s?[A-Z]',
    "Swedish Personnummer":   r'(?:19|20)?\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])[-\s]?\d{4}',
    "Swedish Phone Numbers":  r'(?:\+46[\s-]?|0)?7[0-9][\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}',
    "Swedish Postcodes":      r'\d{3}\s?\d{2}',
    "Credit Card Numbers":    r'\d{4}[\s-]?\d{4}[\s-]?\d{4}[\s-]?\d{4}',
    "Dates of Birth":         r'\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}',
    "IP Addresses":           r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}',
    "Salary / Currency":      r'[£€$]\s?\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})?',
    "SEK Currency":           r'\d{1,3}(?:\s\d{3})+\s*kr',
}

# Which language each pattern belongs to: 'en', 'sv', or 'both'
PATTERN_LANG = {
    "Person Names (NLP)":    "both",
    "Email Addresses":       "both",
    "UK Phone Numbers":      "both",
    "US Phone Numbers":      "both",
    "UK Postcodes":          "both",
    "National Insurance":    "both",
    "Swedish Personnummer":  "both",
    "Swedish Phone Numbers": "both",
    "Swedish Postcodes":     "both",
    "Credit Card Numbers":   "both",
    "Dates of Birth":        "both",
    "IP Addresses":          "both",
    "Salary / Currency":     "both",
    "SEK Currency":          "both",
}

# Bilingual display names for each pattern
PATTERN_LABELS = {
    "en": {
        "Person Names (NLP)":    "Person Names (NLP) ✨",
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
        "Person Names (NLP)":    "Personnamn (NLP) ✨",
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
def replace_with_fake(category, lang):
    """Returns realistic fake data for a given category and language."""
    fake = fake_sv if lang == "sv" else fake_en

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

    replacements = {
        "Email Addresses":       fake.email,
        "UK Phone Numbers":      fake_en.phone_number,
        "US Phone Numbers":      fake_en.phone_number,
        "UK Postcodes":          fake_en.postcode,
        "National Insurance":    lambda: f"{fake_en.lexify('??').upper()}{fake_en.numerify('######')}{fake_en.lexify('?').upper()}",
        "Swedish Personnummer":  fake_personnummer,
        "Swedish Phone Numbers": fake_swedish_phone,
        "Swedish Postcodes":     fake_swedish_postcode,
        "Credit Card Numbers":   lambda: fake.credit_card_number(card_type=None),
        "Dates of Birth":        lambda: fake.date_of_birth(minimum_age=18, maximum_age=80).strftime('%d/%m/%Y'),
        "IP Addresses":          fake.ipv4,
        "Salary / Currency":     lambda: f"£{fake_en.random_int(20000, 120000):,}",
        "SEK Currency":          fake_sek,
    }
    fn = replacements.get(category)
    return fn() if fn else "****"

# ─────────────────────────────────────────────
#  CORE MASKING LOGIC
# ─────────────────────────────────────────────
def mask_cell(value, selected_patterns, mask_mode, token_counters, lang,
              use_nlp=False, nlp_lang="auto", nlp_en=None, nlp_sv=None):
    text = str(value)
    original_text = text  # Store original for debug

    # Check if Person Names (NLP) is enabled
    use_nlp_for_names = "Person Names (NLP)" in selected_patterns

    for category, pattern in PATTERNS.items():
        if category not in selected_patterns:
            continue

        # Skip the NLP pattern in regex loop (handled separately below)
        if category == "Person Names (NLP)":
            continue

        # Debug: Check if pattern matches before applying
        before_text = text
        if mask_mode in ("Asterisks (*****)", "Asterisker (*****)"):
            text = re.sub(pattern, lambda m: '*' * len(m.group()), text)
        elif mask_mode in ("Fake Realistic Data", "Falsk realistisk data"):
            text = re.sub(pattern, lambda m, c=category: replace_with_fake(c, lang), text)
        elif mask_mode in ("Token (e.g. EMAIL_001)", "Token (t.ex. EMAIL_001)"):
            def token_replace(m, c=category):
                key = c.replace(" ", "_").upper()
                token_counters[key] = token_counters.get(key, 0) + 1
                return f"{key}_{token_counters[key]:03d}"
            text = re.sub(pattern, token_replace, text)

        # Debug logging
        if text != before_text:
            print(f"[MASK] {category}: '{original_text}' -> '{text}'")

    # Apply NLP for person names if enabled
    if use_nlp_for_names and nlp_en and nlp_sv:
        text = mask_names_ner(text, mask_mode, token_counters, nlp_lang, nlp_en, nlp_sv)

    return text


def process_dataframe(df, selected_patterns, mask_mode, selected_columns, lang,
                      use_nlp=False, nlp_lang="auto"):
    """Applies masking to a dataframe. Returns masked df + a report."""
    masked_df = df.copy()
    token_counters = {}
    report = []

    # Load NLP models if Person Names (NLP) is selected
    nlp_en, nlp_sv = (None, None)
    if "Person Names (NLP)" in selected_patterns:
        nlp_en, nlp_sv = load_nlp_models()

    cols_to_process = selected_columns if selected_columns else df.columns.tolist()

    for col in cols_to_process:
        if col not in df.columns:
            continue
        for idx, value in df[col].items():
            if pd.isna(value):
                continue
            original = str(value)
            cleaned = mask_cell(
                original, selected_patterns, mask_mode, token_counters, lang,
                use_nlp=True, nlp_lang=nlp_lang, nlp_en=nlp_en, nlp_sv=nlp_sv
            )
            if cleaned != original:
                masked_df.at[idx, col] = cleaned
                report.append({
                    "Row": idx + 2,
                    "Column": col,
                    "Original": original,
                    "Masked As": cleaned,
                })

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


# ─────────────────────────────────────────────
#  UI — STREAMLIT INTERFACE
# ─────────────────────────────────────────────

# Initialize session state
if 'masked_df' not in st.session_state:
    st.session_state.masked_df = None
if 'report_df' not in st.session_state:
    st.session_state.report_df = None
if 'original_df' not in st.session_state:
    st.session_state.original_df = None

# Language selector at top of sidebar (BEFORE settings)
with st.sidebar:
    lang = st.radio(
        "Language / Språk",
        options=["en", "sv"],
        format_func=lambda x: "🇬🇧 English" if x == "en" else "🇸🇪 Svenska",
        horizontal=True,
        key="language_selector"
    )
    st.divider()

ui = T[lang]
labels = PATTERN_LABELS[lang]

# Professional Header
st.title(ui["title"])
st.markdown(f'<p class="subtitle">{ui["subtitle"]}</p>', unsafe_allow_html=True)

# Security Trust Banner
security_message = "🛡️ 100% Secure | 🔒 Client-Side Processing | ✅ GDPR Compliant | 🚫 Zero Data Storage" if lang == "en" else "🛡️ 100% Säker | 🔒 Klientbaserad bearbetning | ✅ GDPR-kompatibel | 🚫 Ingen datalagring"
st.markdown(f'<div class="security-banner">{security_message}</div>', unsafe_allow_html=True)

st.divider()

with st.sidebar:
    st.header(ui["settings"])

    st.subheader(ui["what_to_mask"])
    selected_patterns = []
    for category in PATTERNS:
        if PATTERN_LANG[category] not in (lang, "both"):
            continue
        if st.checkbox(labels[category], value=True, key=f"chk_{category}"):
            selected_patterns.append(category)

    st.subheader(ui["masking_style"])
    mask_mode = st.radio(
        ui["mask_mode_label"],
        ui["mask_modes"],
        index=1,
        key="mask_mode_radio"
    )

    # Clear results when mask mode changes
    if 'previous_mask_mode' not in st.session_state:
        st.session_state.previous_mask_mode = mask_mode
    elif st.session_state.previous_mask_mode != mask_mode:
        st.session_state.masked_df = None
        st.session_state.report_df = None
        st.session_state.original_df = None
        st.session_state.previous_mask_mode = mask_mode

    st.divider()
    st.caption(ui["tip"])

# --- File Upload ---
uploaded_file = st.file_uploader(
    ui["upload_label"],
    type=["xlsx", "csv"],
    help=ui["upload_help"],
)

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(ui["file_error"].format(e=e))
        st.stop()

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
            # Debug: Print selected patterns
            print(f"\n[DEBUG] Selected patterns: {selected_patterns}")
            print(f"[DEBUG] Mask mode: {mask_mode}")

            # Check if Person Names (NLP) is selected
            if "Person Names (NLP)" in selected_patterns:
                with st.spinner("Loading AI models for name detection..."):
                    nlp_en, nlp_sv = load_nlp_models()
                    if not nlp_en or not nlp_sv:
                        st.error("⚠️ NLP models not available. Person names will not be masked. Please install: `python -m spacy download en_core_web_sm` and `python -m spacy download sv_core_news_sm`")
                    else:
                        st.info("✅ AI models loaded successfully")

            with st.spinner(ui["spinner"]):
                st.session_state.masked_df, st.session_state.report_df = process_dataframe(
                    df, selected_patterns, mask_mode, selected_columns, lang,
                    use_nlp=True, nlp_lang="auto"
                )
                st.session_state.original_df = df

    # Display results if they exist in session state
    if st.session_state.masked_df is not None:
        st.success(ui["done"].format(n=len(st.session_state.report_df)))
        st.divider()

        col1, col2 = st.columns(2)
        with col1:
            st.subheader(ui["original_sample"])
            st.dataframe(st.session_state.original_df.head(10), use_container_width=True)
        with col2:
            st.subheader(ui["masked_sample"])
            st.dataframe(st.session_state.masked_df.head(10), use_container_width=True)

        if not st.session_state.report_df.empty:
            with st.expander(ui["report_expander"].format(n=len(st.session_state.report_df))):
                st.dataframe(st.session_state.report_df, use_container_width=True)

        st.divider()

        st.subheader(ui["download_header"])
        excel_output = export_to_excel(st.session_state.original_df, st.session_state.masked_df, st.session_state.report_df)
        original_name = uploaded_file.name.replace('.xlsx', '').replace('.csv', '')

        st.download_button(
            label=ui["download_button"],
            data=excel_output,
            file_name=f"{original_name}_MASKED.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption(ui["download_caption"])

else:
    st.info(ui["landing_info"])
    st.markdown(ui["landing_table_header"])

    # Generate dynamic examples based on selected mask_mode
    col_header = "Type" if lang == "en" else "Typ"
    ex_header  = "Example" if lang == "en" else "Exempel"
    as_header  = "Masked As" if lang == "en" else "Maskeras som"

    # Mapping of display labels to pattern categories
    label_to_category = {
        "Person Name (NLP) ✨": "PERSON",
        "Personnamn (NLP) ✨": "PERSON",
        "Email": "Email Addresses",
        "E-post": "Email Addresses",
        "UK Phone": "UK Phone Numbers",
        "Brittiskt telefon": "UK Phone Numbers",
        "UK Postcode": "UK Postcodes",
        "Brittiskt postnummer": "UK Postcodes",
        "NI Number": "National Insurance",
        "NI-nummer (UK)": "National Insurance",
        "Salary (£€$)": "Salary / Currency",
        "Lön (£€$)": "Salary / Currency",
        "Personnummer": "Swedish Personnummer",
        "Swedish Phone": "Swedish Phone Numbers",
        "Svenskt mobilnummer": "Swedish Phone Numbers",
        "Swedish Postcode": "Swedish Postcodes",
        "Svenskt postnummer": "Swedish Postcodes",
        "SEK Amount": "SEK Currency",
        "SEK-belopp": "SEK Currency",
        "Credit Card": "Credit Card Numbers",
        "Kreditkort": "Credit Card Numbers",
        "Date of Birth": "Dates of Birth",
        "Födelsedatum": "Dates of Birth",
        "IP Address": "IP Addresses",
        "IP-adress": "IP Addresses",
    }

    rows = ui["landing_table_rows"]
    table_md = f"| {col_header} | {ex_header} | {as_header} |\n|---|---|---|\n"

    token_counters_demo = {}
    for label, example, _, row_lang in rows:
        if row_lang in (lang, "both"):
            # Generate masked version based on current mask_mode
            category = label_to_category.get(label)

            if category == "PERSON":
                # Handle person names with NER logic
                if mask_mode in ("Asterisks (*****)", "Asterisker (*****)"):
                    masked_example = '*' * len(example)
                elif mask_mode in ("Fake Realistic Data", "Falsk realistisk data"):
                    fake = fake_sv if lang == "sv" else fake_en
                    masked_example = fake.name()
                else:  # Token
                    token_counters_demo["PERSON"] = token_counters_demo.get("PERSON", 0) + 1
                    masked_example = f"PERSON_{token_counters_demo['PERSON']:03d}"
            elif category and category in PATTERNS:
                # Use the mask_cell logic for other patterns
                masked_example = mask_cell(example, [category], mask_mode, token_counters_demo, lang)
            else:
                masked_example = example

            table_md += f"| {label} | {example} | {masked_example} |\n"

    st.markdown(table_md)

    # Trust badges section
    st.markdown("---")
    if lang == "en":
        st.markdown("### Why Choose DataMask?")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown('<div class="feature-box">🔒 <b>100% Client-Side</b><br>Your data never leaves your browser. Zero uploads to servers.</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="feature-box">🤖 <b>AI-Powered Detection</b><br>Advanced NLP recognizes names in English & Swedish automatically.</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="feature-box">⚡ <b>Instant Processing</b><br>Mask thousands of rows in seconds. Download immediately.</div>', unsafe_allow_html=True)
    else:
        st.markdown("### Varför välja DataMask?")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown('<div class="feature-box">🔒 <b>100% Klientbaserad</b><br>Din data lämnar aldrig din webbläsare. Inga uppladdningar till servrar.</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="feature-box">🤖 <b>AI-driven identifiering</b><br>Avancerad NLP känner igen namn på engelska och svenska automatiskt.</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="feature-box">⚡ <b>Omedelbar bearbetning</b><br>Maskera tusentals rader på sekunder. Ladda ner direkt.</div>', unsafe_allow_html=True)

# Professional Footer
footer_text = "Made with ❤️ | DataMask © 2025 | Privacy-First Data Anonymization" if lang == "en" else "Skapad med ❤️ | DataMask © 2025 | Integritetsfokuserad dataanonymisering"
st.markdown(f'<div class="custom-footer">{footer_text}</div>', unsafe_allow_html=True)
