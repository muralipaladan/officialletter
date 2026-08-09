import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import datetime

st.set_page_config(
    page_title="ഔദ്യോഗിക രേഖ നിർമ്മാതാവ്",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Styles ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&family=Noto+Sans+Malayalam:wght@400;500;600&display=swap');

body, .stApp { background: #F5F0E8; }

.main-header {
    background: linear-gradient(135deg, #8C2F39, #6E2229);
    color: white;
    padding: 20px 28px;
    border-radius: 8px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 16px;
}
.main-header h1 { margin:0; font-size:1.6rem; font-family:'Noto Serif Malayalam',serif; }
.main-header p  { margin:0; opacity:.85; font-size:.9rem; }
.badge {
    background: rgba(255,255,255,.2);
    border-radius: 50%;
    width: 56px; height: 56px;
    display: flex; align-items:center; justify-content:center;
    font-size: 1.6rem; font-weight:700; flex-shrink:0;
}

.section-card {
    background: white;
    border-radius: 6px;
    padding: 20px;
    margin-bottom: 16px;
    border-left: 4px solid #8C2F39;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
}
.section-title {
    font-family:'Noto Serif Malayalam',serif;
    color: #6E2229;
    font-size: 1rem;
    font-weight: 700;
    margin-bottom: 12px;
    display: flex; align-items:center; gap:8px;
}

/* Letter preview sheet */
.letter-sheet {
    background: white;
    border: 1px solid #D9D0BC;
    border-radius: 4px;
    padding: 48px 52px;
    font-family: 'Noto Serif Malayalam', serif;
    font-size: 1.05rem;
    line-height: 2;
    color: #1a1a1a;
    min-height: 400px;
    box-shadow: 0 4px 20px rgba(0,0,0,.1);
    white-space: pre-wrap;
    position: relative;
}
.letter-sheet::before {
    content: '';
    position: absolute; top:0; left:0; right:0; height:4px;
    background: linear-gradient(90deg, #8C2F39, #A9812F);
    border-radius: 4px 4px 0 0;
}

.empty-sheet {
    background: white;
    border: 2px dashed #D9D0BC;
    border-radius: 4px;
    padding: 60px;
    text-align: center;
    color: #9a8f80;
    font-family:'Noto Sans Malayalam',sans-serif;
    font-size: .95rem;
}

.doc-type-badge {
    background: #FFF3F4;
    border: 1px solid #EFC9C9;
    color: #8C2F39;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: .78rem;
    font-weight: 600;
    font-family: 'Noto Sans Malayalam', sans-serif;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #1E2A3A;
}
[data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] label {
    color: #E8E0D0 !important;
}
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
    color: white !important;
}

@media print {
    .stSidebar, .stButton, [data-testid="stToolbar"],
    .main-header, .section-card, header, footer { display:none !important; }
    .letter-sheet { box-shadow:none; border:none; padding: 0; }
}
</style>
""", unsafe_allow_html=True)

# ── Document Types ───────────────────────────────────────────────────────
DOC_TYPES = {
    "കത്തുകൾ": {
        "letter": "ഔദ്യോഗിക കത്ത് (Official Letter)",
        "do_letter": "അർദ്ധ-ഔദ്യോഗിക കത്ത് (D.O. Letter)",
        "forwarding": "അയക്കൽ കത്ത് (Forwarding Letter)",
        "reminder": "ഓർമ്മപ്പെടുത്തൽ കത്ത്",
        "invitation": "ക്ഷണക്കത്ത്",
    },
    "ഉത്തരവുകൾ": {
        "order": "ഉത്തരവ് / നടപടിക്രമം",
        "sanction": "അനുമതി ഉത്തരവ്",
        "transfer": "സ്ഥലംമാറ്റ ഉത്തരവ്",
    },
    "അറിയിപ്പുകൾ": {
        "circular": "സർക്കുലർ",
        "public_notice": "പൊതു അറിയിപ്പ്",
        "tender_notice": "ടെണ്ടർ അറിയിപ്പ്",
        "show_cause": "കാരണം കാണിക്കൽ നോട്ടീസ്",
    },
    "മറ്റുള്ളവ": {
        "note": "നോട്ട് (File Note)",
        "memo": "മെമ്മോറാണ്ടം",
        "certificate": "സർട്ടിഫിക്കറ്റ്",
        "rti_reply": "വിവരാവകാശ മറുപടി (RTI)",
    },
    "അപേക്ഷകൾ (Public → Office)": {
        "app_general": "പൊതു അപേക്ഷ",
        "app_income": "വരുമാന സർട്ടിഫിക്കറ്റ് അപേക്ഷ",
        "app_nativity": "ജനന/നാട്ടുകാർ സർട്ടിഫിക്കറ്റ് അപേക്ഷ",
        "app_residence": "താമസ സർട്ടിഫിക്കറ്റ് അപേക്ഷ",
        "app_caste": "ജാതി/കമ്മ്യൂണിറ്റി സർട്ടിഫിക്കറ്റ് അപേക്ഷ",
        "app_noc": "NOC അപേക്ഷ",
        "app_building": "കെട്ടിട അനുമതി അപേക്ഷ",
        "app_trade": "വ്യാപാര ലൈസൻസ് അപേക്ഷ",
        "app_pension": "പെൻഷൻ/ആനുകൂല്യ അപേക്ഷ",
        "app_land": "ഭൂമി അനുബന്ധ അപേക്ഷ",
        "app_complaint": "പരാതി / Grievance",
        "app_leave": "അവധി അപേക്ഷ",
        "app_scholarship": "സ്കോളർഷിപ്പ് അപേക്ഷ",
        "app_water": "കുടിവെള്ള/ഡ്രെയിനേജ് കണക്ഷൻ",
        "app_road": "റോഡ്/ഇൻഫ്രാ ആവശ്യ അപേക്ഷ",
    }
}

APP_TYPES = list(DOC_TYPES["അപേക്ഷകൾ (Public → Office)"].keys())
NO_RECIPIENT = ['note', 'order', 'sanction', 'memo']
ALL_DOC_TYPES = {k: v for group in DOC_TYPES.values() for k, v in group.items()}

COMMON_RULES = """
പൊതു നിയമങ്ങൾ:
- ലളിതവും സ്പഷ്ടവുമായ ഭരണമലയാളം ഉപയോഗിക്കുക.
- വലിയ വിഷയങ്ങൾ ഖണ്ഡികകളായി തിരിക്കുക.
- ആവർത്തനം ഒഴിവാക്കുക, വസ്തുതകൾ മാത്രം.
- അന്തിമ ഔട്ട്പുട്ടിൽ കത്ത്/രേഖ മാത്രം നൽകുക — Markdown symbols (** ## -- ) ഇടരുത്.
- ശരിയായ വരി ഇടവേളകളോടെ format ചെയ്യുക."""

FORMAT_GUIDES = {
    "letter": "ഓഫീസ് ഹെഡർ, നമ്പർ/തീയതി, സ്വീകർത്താവ് വിലാസം, 'വിഷയം:', 'സൂചന:' (ആവശ്യമെങ്കിൽ), ഖണ്ഡികകൾ, 'വിശ്വസ്തതയോടെ', ഒപ്പ്/പദവി.",
    "do_letter": "D.O. Letter: 'പ്രിയപ്പെട്ട ശ്രീ. [പേര്],' – സൗഹാർദ്ദ professional ഭാഷ, 'സ്നേഹപൂർവ്വം' closing.",
    "forwarding": "ഹെഡർ, 'വിഷയം:', 'സൂചന:', 'മേൽ സൂചിപ്പിച്ച രേഖ ഇതോടൊപ്പം അയക്കുന്നു, ആവശ്യമായ നടപടി സ്വീകരിക്കണം' – ഹ്രസ്വം.",
    "reminder": "'സൂചന:' ൽ മുൻ കത്ത്, മാന്യഭാഷ, വേഗം നടപടി ആവശ്യം.",
    "invitation": "ചടങ്ങ്/തീയതി/സമയം/സ്ഥലം, ഊഷ്മള ഭാഷ, പങ്കെടുക്കണം request.",
    "order": "'പരാമർശം', വസ്തുത വിശദീകരണം, 'ഇതിനാൽ ഉത്തരവാകുന്നു.', 'പകർപ്പ്'.",
    "sanction": "ആവശ്യം, ബാധകമായ ചട്ടം, തുക/നിബന്ധനകൾ, 'സാങ്ഷൻ ചെയ്ത് ഉത്തരവാകുന്നു.'",
    "transfer": "ജീവനക്കാരൻ, നിലവിലെ/പുതിയ ഓഫീസ്, പ്രാബല്യ തീയതി, ചാർജ് transfer നിർദ്ദേശം.",
    "circular": "നിർദ്ദേശം/അറിയിപ്പ്, ബാധകമായവർ, സമയപരിധി (ഉണ്ടെങ്കിൽ).",
    "public_notice": "ആരെ ബാധിക്കുന്നു, കാര്യം, നടപടി/അവസാന തീയതി – ലളിത ഭാഷ.",
    "tender_notice": "പണിയുടെ പേര്, estimate, യോഗ്യത, tender submit തീയതി/സ്ഥലം.",
    "show_cause": "ആരോപണം/വീഴ്ച, ബാധകമായ ചട്ടം, X ദിവസത്തിനകം മറുപടി, 'ഇല്ലെങ്കിൽ തുടർനടപടി'.",
    "note": "'വായിക്കുക:', ഹ്രസ്വ വിവരണം, ചട്ടം, ശുപാർശ – objective ഭാഷ.",
    "memo": "Internal communication, ഹ്രസ്വം, നേരിട്ടുള്ള ഭാഷ.",
    "certificate": "'ഇതിനാൽ സാക്ഷ്യപ്പെടുത്തുന്നത്...', കൃത്യം, ഒപ്പ്/സീൽ സ്ഥലം.",
    "rti_reply": """RTI Act 2005 / Kerala RTI Rules 2006 പ്രകാരം SPIO reply:
1. ഹെഡർ: SPIO ഓഫീസ്, നമ്പർ, തീയതി.
2. 'വിഷയം: RTI 2005 – മറുപടി'
3. 'സൂചന:' – അപേക്ഷ തീയതി/നമ്പർ.
4. ഓരോ ചോദ്യത്തിനും 'ചോദ്യം N:' ചോദ്യം ആവർത്തിക്കാതെ 'ഉത്തരം:' മാത്രം.
5. ഒടുവിൽ 30 ദിവസത്തിനകം First Appellate Authority-ക്ക് appeal (Section 19(1)) paragraph.
6. SPIO ഒപ്പ്.""",
    "app_general": "അപേക്ഷ format: 'മഹോദയ/മഹോദയേ,' – അപേക്ഷകൻ intro, ആവശ്യം, grounds, request, 'അപേക്ഷകൻ/അപേക്ഷക' ഒപ്പ്, വിലാസം, തീയതി.",
    "app_income": "വരുമാന source/ഏകദേശ തുക, ഉദ്ദേശ്യം, village officer verify ആവശ്യം. സത്യസന്ധ declaration.",
    "app_nativity": "ജനനം/നിരന്തര residence confirm, ഉദ്ദേശ്യം, ജനനതീയതി/ജന്മഗ്രാമം/ഇപ്പോഴത്തെ മേൽ‌വിലാസം.",
    "app_residence": "X വർഷം/തീയതി മുതൽ residence, ഉദ്ദേശ്യം, ID proof reference.",
    "app_caste": "ജാതി/community/list, ഉദ്ദേശ്യം, revenue authority ദ്വാര verify ആവശ്യം.",
    "app_noc": "ഏത് ആവശ്യം, activity/location, ബന്ധപ്പെട്ട dept-ൽ submit, objection ഇല്ലെന്ന് confirm.",
    "app_building": "Plot/Survey No., ഉദ്ദേശ്യം, ഫ്ലോർ area, Building Rules compliance, permit request.",
    "app_trade": "ബിസിനസ് പേര്/trade/Ward, owner info, NOC ready, license/renewal request.",
    "app_pension": "ഏത് pension scheme, അർഹത, proof list, bank account, direct transfer request.",
    "app_land": "Survey No., owner, ആവശ്യ change, supporting docs, Tahsildar/RI approval request.",
    "app_complaint": "ആര്/വിഷയം/സംഭവം, prior complaint ഉണ്ടോ, ആഗ്രഹിക്കുന്ന remedy – objective ഭാഷ.",
    "app_leave": "Designation/office, leave type/dates, കാരണം, alternate arrangement, balance reference.",
    "app_scholarship": "Scheme, qualification/marks, income, community, bank account, docs attached.",
    "app_water": "Connection type/address/Ward, present source, pipeline availability, fee paid reference.",
    "app_road": "Location/Ward, problem description, affected count, estimate/survey ഉണ്ടോ, priority reason.",
}

def build_prompt(data: dict) -> str:
    is_app = data['doc_type'] in APP_TYPES
    
    applicant_block = ""
    if is_app:
        applicant_block = f"""
അപേക്ഷകന്റെ വിവരങ്ങൾ:
- പേര്: {data.get('app_name') or '(നൽകിയിട്ടില്ല)'}
- വയസ്സ്/ജനനതീയതി: {data.get('app_age') or '(നൽകിയിട്ടില്ല)'}
- വിലാസം: {data.get('app_addr') or '(നൽകിയിട്ടില്ല)'}
- ഫോൺ: {data.get('app_phone') or '(നൽകിയിട്ടില്ല)'}
- ID/Proof: {data.get('app_id') or '(നൽകിയിട്ടില്ല)'}"""

    system_ctx = (
        "നീ ഒരു document expert ആണ്. Kerala-യിലെ ഒരു പൗരൻ government/local body office-ൽ ഔദ്യോഗിക അപേക്ഷ നൽകേണ്ടതുണ്ട്. നൽകിയ വിവരങ്ങൾ ഉപയോഗിച്ച് ഭരണമലയാള ശൈലിയിൽ അപേക്ഷ തയ്യാറാക്കുക."
        if is_app else
        f"നീ Kerala സർക്കാർ/തദ്ദേശ സ്ഥാപന ഫയൽ എഴുത്തിൽ വിദഗ്ധനായ ഒരു സീനിയർ ഉദ്യോഗസ്ഥനാണ്. ശരിയായ ഭരണമലയാള ശൈലിയിൽ ഒരു {ALL_DOC_TYPES[data['doc_type']]} തയ്യാറാക്കുക."
    )

    return f"""{system_ctx}

Format നിർദ്ദേശം:
{FORMAT_GUIDES[data['doc_type']]}
{COMMON_RULES}

നൽകിയ വിവരങ്ങൾ:
- {'Authority/ഓഫീസ്' if is_app else 'ഓഫീസ്'}: {data.get('office_name') or '(നൽകിയിട്ടില്ല)'}
- വിലാസം: {data.get('office_addr') or '(നൽകിയിട്ടില്ല)'}
- {'Ref No. (ഐച്ഛികം)' if is_app else 'ഫയൽ നം.'}: {data.get('file_no') or '(നൽകിയിട്ടില്ല)'}
- തീയതി: {data.get('date_str') or '(നൽകിയിട്ടില്ല)'}
- {'ആർക്ക്' if not is_app else 'ഏത് ഓഫീസ്/അധികാരി'}: {data.get('to_whom') or '(ബാധകമല്ല)'}
- വിഷയം: {data.get('subject') or '(നൽകിയിട്ടില്ല)'}
- സൂചന: {data.get('reference') or '(ഇല്ല)'}{applicant_block}
- {'ആവശ്യം / Details' if is_app else 'പ്രധാന കാര്യങ്ങൾ (rough notes)'}:
{data.get('points') or '(നൽകിയിട്ടില്ല)'}
- {'അപേക്ഷകൻ' if is_app else 'ഒപ്പ്'}: {data.get('sign_name') or ''} {('(' + data['sign_desig'] + ')') if data.get('sign_desig') else ''}

നിർദ്ദേശം: rough notes ഔദ്യോഗിക ഖണ്ഡികകളാക്കുക. Output-ൽ കത്ത്/അപേക്ഷ മാത്രം, Markdown symbols ഇടരുത്."""


def generate_letter(api_key: str, model_name: str, data: dict) -> str:
    # പുതിയ Google GenAI SDK Client ഉപയോഗിക്കുന്നു
    client = genai.Client(api_key=api_key)
    prompt = build_prompt(data)
    
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.4
        )
    )
    return response.text.strip()


def make_docx(text: str, doc_type_label: str) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2.5)

    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(f"[{doc_type_label}]")
    title_run.font.size = Pt(9)
    title_run.font.color.rgb = RGBColor(0x8C, 0x2F, 0x39)
    title_run.font.bold = True
    title_run.font.name = "Noto Serif Malayalam"
    pPr = title_para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), '8C2F39')
    pBdr.append(bottom)
    pPr.append(pBdr)

    doc.add_paragraph()

    for line in text.split('\n'):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(line)
        run.font.size = Pt(11)
        run.font.name = "Noto Serif Malayalam"
        pPr2 = p._p.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:line'), '360')
        spacing.set(qn('w:lineRule'), 'auto')
        pPr2.append(spacing)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Session State Init ───────────────────────────────────────────────────
for key, default in {
    'generated_text': '',
    'edited_text': '',
    'edit_mode': False,
    'api_key': '',
    'model': 'gemini-3.1-flash',
    'office_name': '',
    'office_addr': '',
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# Streamlit Secrets-ൽ നിന്ന് Key സ്വയം ലോഡ് ചെയ്യുന്നു
secret_key = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else ""

# ── Sidebar ──────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔑 Gemini API")

    with st.container():
        input_key = st.text_input(
            "API Key",
            type="password",
            value=st.session_state.api_key or secret_key,
            placeholder="AIza...",
            help="Secrets-ൽ നൽകിയിട്ടുണ്ടെങ്കിൽ ഇവിടെ നേരിട്ട് ലോഡ് ആകും."
        )
        
        # 3.1 ഉം അതിനു മുകളിലുള്ള മോഡലുകളും മാത്രം ഉൾപ്പെടുത്തിയിരിക്കുന്നു
        model = st.selectbox(
            "Model",
            ["gemini-3.1-flash", "gemini-3.1-pro", "gemini-3.2-flash", "gemini-3.4-flash"],
            index=0
        )
        
        active_key = input_key or secret_key
        if active_key:
            st.session_state.api_key = active_key
            st.session_state.model = model
            st.success("✓ API Key Ready", icon="✅")

    st.divider()
    st.markdown("## 🏢 ഓഫീസ് Defaults")
    st.session_state.office_name = st.text_input(
        "ഓഫീസ് പേര്",
        value=st.session_state.office_name,
        placeholder="ഉദാ: നിലമ്പൂർ ഗ്രാമ പഞ്ചായത്ത്"
    )
    st.session_state.office_addr = st.text_input(
        "വിലാസം",
        value=st.session_state.office_addr,
        placeholder="ഉദാ: നിലമ്പൂർ, മലപ്പുറം - 679329"
    )

    st.divider()
    st.markdown("## ℹ️ Help")
    st.caption("""
**Secrets Setup**:
`.streamlit/secrets.toml` ഫയലിൽ `GEMINI_API_KEY` ചേർക്കുക.
    """)


# ── Main Header ──────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <div class="badge">കേ</div>
  <div>
    <h1>ഔദ്യോഗിക രേഖ നിർമ്മാതാവ്</h1>
    <p>Kerala Government · ഭരണമലയാളത്തിൽ കത്ത്, ഉത്തരവ്, അപേക്ഷ, RTI — AI ഉപയോഗിച്ച് seconds-ൽ</p>
  </div>
</div>
""", unsafe_allow_html=True)

col_form, col_preview = st.columns([1, 1], gap="large")

# ── Left Column: Form ────────────────────────────────────────────────────
with col_form:

    st.markdown('<div class="section-card"><div class="section-title">📋 രേഖയുടെ തരം</div>', unsafe_allow_html=True)
    
    group_options = list(DOC_TYPES.keys())
    selected_group = st.selectbox("Category", group_options, label_visibility="collapsed")
    
    type_options = DOC_TYPES[selected_group]
    doc_type = st.selectbox(
        "Type",
        options=list(type_options.keys()),
        format_func=lambda x: type_options[x],
        label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    is_app = doc_type in APP_TYPES

    if is_app:
        st.markdown('<div class="section-card"><div class="section-title">👤 അപേക്ഷകന്റെ വിവരങ്ങൾ</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        app_name = c1.text_input("പേര് *", placeholder="ഉദാ: രാജേഷ് കുമാർ കെ.")
        app_age  = c2.text_input("വയസ്സ് / ജനനതീയതി", placeholder="ഉദാ: 42 / 15-03-1982")
        app_addr = st.text_area("വിലാസം", placeholder="ഉദാ: 'ശ്രീനിലയം', ആനക്കയം P.O., നിലമ്പൂർ", height=80)
        c3, c4 = st.columns(2)
        app_phone = c3.text_input("ഫോൺ", placeholder="9876543210")
        app_id    = c4.text_input("ID / Proof", placeholder="Voter ID: ABC1234567")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        app_name = app_age = app_addr = app_phone = app_id = ""

    st.markdown(f'<div class="section-card"><div class="section-title">🏢 {"Authority / ഓഫീസ്" if is_app else "ഓഫീസ് വിവരങ്ങൾ"}</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    office_name = c1.text_input(
        "ഓഫീസ് / Authority",
        value=st.session_state.office_name,
        placeholder="ഉദാ: നിലമ്പൂർ ഗ്രാമ പഞ്ചായത്ത്"
    )
    office_addr = c2.text_input(
        "വിലാസം",
        value=st.session_state.office_addr,
        placeholder="ഉദാ: നിലമ്പൂർ, മലപ്പുറം - 679329"
    )
    c3, c4 = st.columns(2)
    file_no = c3.text_input("ഫയൽ നം. / Ref No.", placeholder="ഉദാ: B2-1234/2026")
    date_val = c4.date_input("തീയതി", value=datetime.date.today())
    st.markdown('</div>', unsafe_allow_html=True)

    if doc_type not in NO_RECIPIENT:
        to_whom_labels = {
            "letter": "ആർക്ക് (പദവി/ഓഫീസ്)",
            "app_general": "ആർക്ക് / ഏത് ഓഫീസ്",
            "rti_reply": "അപേക്ഷകന്റെ പേര്/വിലാസം",
        }
        to_label = to_whom_labels.get(doc_type, "ആർക്ക് / Authority")
        st.markdown(f'<div class="section-card"><div class="section-title">📬 {to_label}</div>', unsafe_allow_html=True)
        to_whom = st.text_input(
            to_label,
            placeholder="ഉദാ: സെക്രട്ടറി, നിലമ്പൂർ ഗ്രാമ പഞ്ചായത്ത്",
            label_visibility="collapsed"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        to_whom = ""

    st.markdown('<div class="section-card"><div class="section-title">📝 വിഷയവും വിവരങ്ങളും</div>', unsafe_allow_html=True)
    subject = st.text_input("വിഷയം *", placeholder="ഉദാ: വാർഡ് 12-ലെ റോഡ് അറ്റകുറ്റപ്പണി സംബന്ധിച്ച്")
    reference = st.text_area(
        "സൂചന / Reference (ഐച്ഛികം)",
        placeholder="ഉദാ: 1. ശ്രീ. XXX-ന്റെ അപേക്ഷ dt. 01-06-2026",
        height=70
    )
    points_label = (
        "വിവരാവകാശ ചോദ്യങ്ങൾ (നമ്പറിട്ട്)" if doc_type == "rti_reply"
        else "ആവശ്യം / Details" if is_app
        else "പ്രധാന കാര്യങ്ങൾ (rough notes)"
    )
    points = st.text_area(
        f"{points_label} *",
        placeholder=(
            "ഉദാ:\n1. ചോദ്യം 1...\n2. ചോദ്യം 2..." if doc_type == "rti_reply"
            else "ഉദാ:\n- ആവശ്യത്തിന്റെ കാരണം\n- Survey No., Ward, dates...\n- Attached documents" if is_app
            else "ഉദാ:\n- കുഴി ഉണ്ട്, അപകടം\n- അടിയന്തിരം\n- estimate ആവശ്യം"
        ),
        height=140
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f'<div class="section-card"><div class="section-title">✍️ {"അപേക്ഷകൻ" if is_app else "ഒപ്പ്"}</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    sign_name  = c1.text_input("പേര്", placeholder="ഉദാ: കെ. രാജൻ")
    sign_desig = c2.text_input(
        "പദവി" if not is_app else "തൊഴിൽ/നില (ഐച്ഛികം)",
        placeholder="ഉദാ: സെക്രട്ടറി" if not is_app else "ഉദാ: കർഷകൻ"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    btn_label = "⚡ അപേക്ഷ തയ്യാറാക്കുക" if is_app else "⚡ കത്ത് / രേഖ തയ്യാറാക്കുക"
    generate_clicked = st.button(btn_label, type="primary", use_container_width=True)

    if generate_clicked:
        current_api_key = st.session_state.api_key or secret_key
        if not current_api_key:
            st.error("⚠️ Sidebar-ലോ Streamlit secrets.toml-ലോ API Key നൽകുക.")
        elif not subject or not points:
            st.error("⚠️ വിഷയവും details-ഉം നിർബന്ധമാണ്.")
        elif is_app and not app_name:
            st.error("⚠️ അപേക്ഷകന്റെ പേര് നൽകുക.")
        else:
            data = {
                'doc_type': doc_type,
                'office_name': office_name,
                'office_addr': office_addr,
                'file_no': file_no,
                'date_str': date_val.strftime('%d/%m/%Y'),
                'to_whom': to_whom,
                'subject': subject,
                'reference': reference,
                'points': points,
                'sign_name': sign_name,
                'sign_desig': sign_desig,
                'app_name': app_name,
                'app_age': app_age,
                'app_addr': app_addr,
                'app_phone': app_phone,
                'app_id': app_id,
            }
            with st.spinner("AI രേഖ തയ്യാറാക്കുന്നു..."):
                try:
                    text = generate_letter(current_api_key, st.session_state.model, data)
                    st.session_state.generated_text = text
                    st.session_state.edited_text = text
                    st.session_state.edit_mode = False
                    st.success("✅ തയ്യാറായി! വലതുവശം കാണുക.")
                except Exception as e:
                    st.error(f"❌ Error: {e}")


# ── Right Column: Preview + Edit + Download ──────────────────────────────
with col_preview:
    st.markdown("### 📄 തയ്യാറാക്കിയ രേഖ")

    has_content = bool(st.session_state.generated_text)
    doc_label = ALL_DOC_TYPES.get(doc_type, "")

    if not has_content:
        st.markdown("""
        <div class="empty-sheet">
            <div style="font-size:2.5rem;margin-bottom:12px">📋</div>
            <div>ഇടതുവശത്ത് വിവരങ്ങൾ നൽകി<br><b>Generate</b> click ചെയ്യുക</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        col_e, col_d, col_p, col_n = st.columns([1.2, 1.5, 1, 1])
        
        edit_label = "✏️ Edit Mode" if not st.session_state.edit_mode else "👁️ Preview"
        if col_e.button(edit_label, use_container_width=True):
            st.session_state.edit_mode = not st.session_state.edit_mode
            st.rerun()

        docx_bytes = make_docx(st.session_state.edited_text, doc_label)
        col_d.download_button(
            "⬇️ Download .docx",
            data=docx_bytes,
            file_name=f"letter-{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

        if col_p.button("🖨️ Print", use_container_width=True):
            st.session_state['trigger_print'] = True

        if col_n.button("🆕 പുതിയത്", use_container_width=True):
            st.session_state.generated_text = ''
            st.session_state.edited_text = ''
            st.session_state.edit_mode = False
            st.rerun()

        st.markdown(f'<span class="doc-type-badge">{doc_label}</span>', unsafe_allow_html=True)
        st.markdown("")

        if st.session_state.edit_mode:
            st.caption("✏️ നേരിട്ട് edit ചെയ്യാം — changes auto-save ആകും")
            edited = st.text_area(
                "Edit",
                value=st.session_state.edited_text,
                height=600,
                label_visibility="collapsed",
                key="editor_area"
            )
            if edited != st.session_state.edited_text:
                st.session_state.edited_text = edited
        else:
            safe_text = (st.session_state.edited_text
                        .replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;"))
            
            print_js = ""
            if st.session_state.get('trigger_print'):
                print_js = "<script>window.print();</script>"
                st.session_state['trigger_print'] = False

            st.components.v1.html(f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ background:#f5f0e8; padding:16px; }}
  .sheet {{
    background: white;
    border: 1px solid #D9D0BC;
    border-radius: 4px;
    padding: 48px 52px;
    font-family: 'Noto Serif Malayalam', serif;
    font-size: 14px;
    line-height: 2;
    color: #1a1a1a;
    white-space: pre-wrap;
    box-shadow: 0 4px 20px rgba(0,0,0,.1);
    position: relative;
  }}
  .sheet::before {{
    content: '';
    position: absolute; top:0; left:0; right:0; height:4px;
    background: linear-gradient(90deg, #8C2F39, #A9812F);
    border-radius: 4px 4px 0 0;
  }}
  @media print {{
    body {{ background: white; padding: 0; }}
    .sheet {{ box-shadow: none; border: none; border-radius:0; }}
    .sheet::before {{ display:none; }}
  }}
</style>
</head>
<body>
<div class="sheet">{safe_text}</div>
{print_js}
</body>
</html>
""", height=650, scrolling=True)

        word_count = len(st.session_state.edited_text.split())
        st.caption(f"📊 {word_count} words · {len(st.session_state.edited_text)} characters")

# ── Footer ───────────────────────────────────────────────────────────────
st.divider()
st.caption("🔒 Kerala Government Document Tool · Gemini GenAI SDK Powered")