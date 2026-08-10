import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import datetime
import time

st.set_page_config(
    page_title="ഔദ്യോഗിക അപേക്ഷ",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&family=Noto+Sans+Malayalam:wght@400;500;600&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #f0ebe0 !important;
    font-family: 'Noto Sans Malayalam', sans-serif;
}

/* Hide streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stToolbar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }

/* App wrapper */
.block-container {
    max-width: 780px !important;
    padding: 2rem 1.5rem 4rem !important;
}

/* App header */
.app-header {
    background: linear-gradient(135deg, #7B1C28 0%, #5C1520 100%);
    border-radius: 10px;
    padding: 22px 28px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 4px 20px rgba(123,28,40,.25);
}
.app-header-icon {
    font-size: 2rem;
    background: rgba(255,255,255,.15);
    width: 52px; height: 52px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
}
.app-header h1 {
    margin: 0; color: white;
    font-family: 'Noto Serif Malayalam', serif;
    font-size: 1.3rem; font-weight: 700; line-height: 1.4;
}
.app-header p { margin: 4px 0 0; color: rgba(255,255,255,.75); font-size: .82rem; }

/* Form card */
.form-card {
    background: white;
    border-radius: 10px;
    padding: 24px 28px;
    margin-bottom: 20px;
    box-shadow: 0 2px 12px rgba(0,0,0,.07);
    border-top: 3px solid #7B1C28;
}
.form-section-title {
    font-family: 'Noto Serif Malayalam', serif;
    font-size: .95rem;
    font-weight: 700;
    color: #7B1C28;
    margin-bottom: 14px;
    padding-bottom: 8px;
    border-bottom: 1px solid #f0e8e8;
    display: flex; align-items: center; gap: 8px;
}

/* Streamlit input overrides */
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] > div > div {
    border-radius: 6px !important;
    border-color: #ddd !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .92rem !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
    border-color: #7B1C28 !important;
    box-shadow: 0 0 0 2px rgba(123,28,40,.12) !important;
}

/* Generate button */
[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #7B1C28, #9E2535) !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    height: 52px !important;
    box-shadow: 0 4px 14px rgba(123,28,40,.3) !important;
    transition: all .2s !important;
}
[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(123,28,40,.4) !important;
}
[data-testid="stButton"] > button[kind="secondary"] {
    border-radius: 6px !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .88rem !important;
}

/* AI loading animation */
.ai-loader {
    background: white;
    border-radius: 10px;
    padding: 36px 28px;
    text-align: center;
    box-shadow: 0 2px 12px rgba(0,0,0,.07);
    border-top: 3px solid #7B1C28;
    margin-bottom: 20px;
}
.ai-loader-title {
    font-family: 'Noto Serif Malayalam', serif;
    color: #7B1C28;
    font-size: 1.1rem;
    font-weight: 700;
    margin-bottom: 6px;
}
.ai-loader-sub {
    color: #888;
    font-size: .82rem;
    margin-bottom: 24px;
}
.typing-dots {
    display: inline-flex;
    gap: 8px;
    margin-bottom: 20px;
}
.typing-dots span {
    width: 10px; height: 10px;
    background: #7B1C28;
    border-radius: 50%;
    display: inline-block;
    animation: bounce 1.2s infinite ease-in-out;
}
.typing-dots span:nth-child(2) { animation-delay: .2s; background: #b04455; }
.typing-dots span:nth-child(3) { animation-delay: .4s; background: #d4747f; }
@keyframes bounce {
    0%, 60%, 100% { transform: translateY(0); opacity:.6; }
    30% { transform: translateY(-10px); opacity:1; }
}
.ai-steps {
    display: flex;
    flex-direction: column;
    gap: 8px;
    text-align: left;
    max-width: 340px;
    margin: 0 auto;
}
.ai-step {
    display: flex; align-items: center; gap: 10px;
    font-size: .85rem;
    color: #555;
    padding: 8px 12px;
    border-radius: 6px;
    background: #fdf8f8;
    border: 1px solid #f0e8e8;
    transition: all .3s;
}
.ai-step.active {
    color: #7B1C28;
    background: #fff0f1;
    border-color: #f0b8be;
    font-weight: 600;
}
.ai-step.done {
    color: #2e7d32;
    background: #f0faf0;
    border-color: #a5d6a7;
}
.step-icon { font-size: 1rem; }

/* Output letter */
.letter-output {
    background: white;
    border-radius: 10px;
    padding: 44px 52px;
    box-shadow: 0 4px 24px rgba(0,0,0,.1);
    font-family: 'Noto Serif Malayalam', serif;
    font-size: 1rem;
    line-height: 2.1;
    color: #111;
    white-space: pre-wrap;
    position: relative;
    margin-bottom: 16px;
    border-top: 4px solid #7B1C28;
}
.letter-output::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #7B1C28, #A9812F);
    border-radius: 0 0 10px 10px;
}

/* Action bar */
.action-bar {
    background: white;
    border-radius: 8px;
    padding: 14px 18px;
    display: flex; gap: 10px;
    box-shadow: 0 2px 8px rgba(0,0,0,.06);
    margin-bottom: 16px;
    flex-wrap: wrap;
}

/* Secret key badge */
.secret-badge {
    background: #e8f5e9;
    border: 1px solid #a5d6a7;
    border-radius: 6px;
    padding: 8px 14px;
    color: #2e7d32;
    font-size: .83rem;
    display: flex; align-items: center; gap: 6px;
    margin-bottom: 12px;
}

/* Edit textarea */
[data-testid="stTextArea"].letter-editor textarea {
    font-family: 'Noto Serif Malayalam', serif !important;
    font-size: 1rem !important;
    line-height: 2.1 !important;
    border: 2px solid #7B1C28 !important;
    border-radius: 8px !important;
    padding: 40px 48px !important;
}

/* Divider */
hr { border-color: #e8e0d4 !important; margin: 8px 0 !important; }

/* Error */
[data-testid="stAlert"] {
    border-radius: 8px !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
}

@media print {
    .app-header, .form-card, [data-testid="stButton"],
    .action-bar, [data-testid="stAlert"] { display: none !important; }
    .letter-output {
        box-shadow: none !important;
        border: none !important;
        padding: 0 !important;
    }
    .letter-output::after { display: none !important; }
}
</style>
""", unsafe_allow_html=True)

# ── Constants ────────────────────────────────────────────────────────────

APP_TYPES_LIST = [
    "app_general","app_income","app_nativity","app_residence","app_caste",
    "app_noc","app_building","app_trade","app_pension","app_land",
    "app_complaint","app_leave","app_scholarship","app_water","app_road"
]

DOC_GROUPS = {
    "📨 കത്തുകൾ": {
        "letter":     "ഔദ്യോഗിക കത്ത്",
        "do_letter":  "അർദ്ധ-ഔദ്യോഗിക (D.O.) കത്ത്",
        "forwarding": "അയക്കൽ കത്ത്",
        "reminder":   "ഓർമ്മപ്പെടുത്തൽ",
        "invitation": "ക്ഷണക്കത്ത്",
    },
    "📜 ഉത്തരവ് / അറിയിപ്പ്": {
        "order":        "ഉത്തരവ്",
        "sanction":     "അനുമതി ഉത്തരവ്",
        "circular":     "സർക്കുലർ",
        "public_notice":"പൊതു അറിയിപ്പ്",
        "show_cause":   "കാരണം കാണിക്കൽ നോട്ടീസ്",
        "rti_reply":    "RTI മറുപടി",
    },
    "📝 അപേക്ഷകൾ (Public → Office)": {
        "app_general":    "പൊതു അപേക്ഷ",
        "app_income":     "വരുമാന Certificate അപേക്ഷ",
        "app_nativity":   "ജനന/നാട്ടുകാർ Certificate",
        "app_residence":  "താമസ Certificate",
        "app_caste":      "ജാതി Certificate",
        "app_noc":        "NOC അപേക്ഷ",
        "app_building":   "കെട്ടിട അനുമതി",
        "app_trade":      "വ്യാപാര ലൈസൻസ്",
        "app_pension":    "പെൻഷൻ/ആനുകൂല്യം",
        "app_land":       "ഭൂമി/Mutation",
        "app_complaint":  "പരാതി / Grievance",
        "app_leave":      "അവധി അപേക്ഷ",
        "app_scholarship":"സ്കോളർഷിപ്പ്",
        "app_water":      "കുടിവെള്ള/Drainage",
        "app_road":       "റോഡ്/ഇൻഫ്രാ ആവശ്യം",
    }
}

ALL_TYPES = {k: v for g in DOC_GROUPS.values() for k, v in g.items()}

FORMAT_GUIDES = {
    "letter":      "ഓഫീസ് ഹെഡർ → നം./തീയതി → സ്വീകർത്താവ് → വിഷയം → സൂചന (ഉണ്ടെങ്കിൽ) → ഖണ്ഡികകൾ → 'വിശ്വസ്തതയോടെ' → ഒപ്പ്/പദവി",
    "do_letter":   "'പ്രിയപ്പെട്ട ശ്രീ./ശ്രീമതി [പേര്],' – personal yet professional ഭാഷ → 'സ്നേഹപൂർവ്വം'",
    "forwarding":  "ഹ്രസ്വം. 'മേൽ സൂചിപ്പിച്ച രേഖ ഇതോടൊപ്പം അയക്കുന്നു, ആവശ്യ നടപടി സ്വീകരിക്കണം'",
    "reminder":    "സൂചനയിൽ മുൻ കത്ത്. മാന്യഭാഷ. 'ഇതുവരെ മറുപടി ലഭിച്ചിട്ടില്ല, ഉടൻ നടപടി ആവശ്യം'",
    "invitation":  "ചടങ്ങ്/തീയതി/സമയം/സ്ഥലം. ഊഷ്മള ഭാഷ.",
    "order":       "'പരാമർശം' → വസ്തുത → 'ഇതിനാൽ ഉത്തരവാകുന്നു' → 'പകർപ്പ്:'",
    "sanction":    "ആവശ്യം/ചട്ടം/തുക/നിബന്ധന → 'സാങ്ഷൻ ചെയ്ത് ഉത്തരവാകുന്നു'",
    "circular":    "നിർദ്ദേശം/ആർ ബാധകം/സമയപരിധി",
    "public_notice":"ആർ ബാധകം/കാര്യം/അവസാന തീയതി – ലളിത ഭാഷ",
    "show_cause":  "ആരോപണം/ചട്ടം/'X ദിവസത്തിനകം മറുപടി'/തുടർനടപടി warning",
    "rti_reply":   "RTI Act 2005: SPIO ഹെഡർ → ഓരോ ചോദ്യത്തിനും 'ചോദ്യം N: / ഉത്തരം:' → Section 19(1) appeal para → SPIO ഒപ്പ്",
    "app_general": "'മഹോദയ/മഹോദയേ,' → അപേക്ഷകൻ intro → ആവശ്യം/കാരണം → request → 'അപേക്ഷകൻ' ഒപ്പ്/വിലാസം/തീയതി",
    "app_income":  "വരുമാന source/തുക, ഉദ്ദേശ്യം, village officer verify ആവശ്യം, സത്യസന്ധ declaration",
    "app_nativity":"ജനനം/residence confirm, ഉദ്ദേശ്യം, ജനനതീയതി/ജന്മഗ്രാമം",
    "app_residence":"X വർഷം/തീയതി മുതൽ residence, ഉദ്ദേശ്യം, ID proof reference",
    "app_caste":   "ജാതി/community/list (SC/ST/OBC), ഉദ്ദേശ്യം, Tahsildar verify ആവശ്യം",
    "app_noc":     "ഏത് ആവശ്യം/activity/location, objection ഇല്ലെന്ന് confirm ആവശ്യം",
    "app_building":"Plot/Survey No., ഉദ്ദേശ്യം, floor area, Building Rules compliance, permit request",
    "app_trade":   "ബിസിനസ് പേര്/trade/Ward, owner, NOC ready, license/renewal",
    "app_pension": "ഏത് scheme, അർഹത, proof, bank account, direct transfer request",
    "app_land":    "Survey No., owner, ആവശ്യ change, docs, Tahsildar approval",
    "app_complaint":"ആര്/സംഭവം/തീയതി, prior complaint, ആഗ്രഹിക്കുന്ന remedy – objective",
    "app_leave":   "Designation/office, leave type/dates, കാരണം, alternate arrangement",
    "app_scholarship":"Scheme, qualification/marks, income, community, bank account",
    "app_water":   "Connection type/address/Ward, present source, pipeline, fee paid",
    "app_road":    "Location/Ward, problem, affected count, estimate/survey, priority",
}

COMMON_RULES = """
നിർദ്ദേശങ്ങൾ:
- ശരിയായ ഭരണമലയാളം ഉപയോഗിക്കുക
- Rough notes → ഔദ്യോഗിക ഖണ്ഡികകൾ ആക്കുക
- Output-ൽ letter/application മാത്രം; ** ## -- തുടങ്ങിയ Markdown ഇടരുത്
- ശരിയായ line breaks ഉപയോഗിക്കുക"""

AI_STEPS = [
    ("🔍", "ആവശ്യം മനസ്സിലാക്കുന്നു..."),
    ("📐", "ഔദ്യോഗിക format തിരഞ്ഞെടുക്കുന്നു..."),
    ("✍️", "ഭരണമലയാളത്തിൽ എഴുതുന്നു..."),
    ("✅", "തയ്യാറാക്കി!"),
]

# ── Session State ────────────────────────────────────────────────────────
for k, v in {
    'output': '', 'edit_mode': False,
    'office_name': '', 'office_addr': '',
    'generating': False,
    'doc_label': '',
    'docx_cache': None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── API Key: secrets → fallback input ───────────────────────────────────
secret_key = st.secrets.get("GEMINI_API_KEY", "")

# ── Functions ────────────────────────────────────────────────────────────
def build_prompt(d):
    is_app = d['doc_type'] in APP_TYPES_LIST
    app_block = ""
    if is_app:
        app_block = f"""
അപേക്ഷകൻ:
  പേര്    : {d.get('app_name') or '—'}
  വയസ്സ്  : {d.get('app_age') or '—'}
  വിലാസം : {d.get('app_addr') or '—'}
  ഫോൺ    : {d.get('app_phone') or '—'}
  ID      : {d.get('app_id') or '—'}"""

    ctx = (
        "Kerala government office-ൽ ഒരു പൗരൻ നൽകുന്ന official application തയ്യാറാക്കുക."
        if is_app else
        f"Kerala സർക്കാർ/തദ്ദേശ ഓഫീസ് ഫയൽ എഴുത്തിൽ expert ആണ്. ഒരു {ALL_TYPES[d['doc_type']]} തയ്യാറാക്കുക."
    )
    return f"""{ctx}

Format: {FORMAT_GUIDES[d['doc_type']]}
{COMMON_RULES}

വിവരങ്ങൾ:
  ഓഫീസ്   : {d.get('office_name') or '—'}
  വിലാസം  : {d.get('office_addr') or '—'}
  ഫയൽ നം. : {d.get('file_no') or '—'}
  തീയതി   : {d.get('date_str') or '—'}
  ആർക്ക്   : {d.get('to_whom') or '—'}
  വിഷയം   : {d.get('subject') or '—'}
  സൂചന    : {d.get('reference') or 'ഇല്ല'}{app_block}
  Details  :
{d.get('points') or '—'}
  ഒപ്പ്    : {d.get('sign_name') or ''} {('(' + d['sign_desig'] + ')') if d.get('sign_desig') else ''}"""


def call_gemini(api_key, model_name, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    resp = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(temperature=0.35)
    )
    return resp.text.strip()


def make_docx(text, label):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3);  sec.right_margin = Cm(2.5)
    # label row
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(f"[ {label} ]")
    tr.font.size = Pt(8); tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x7B, 0x1C, 0x28)
    tr.font.name = "Noto Serif Malayalam"
    pPr = tp._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'4')
    bot.set(qn('w:space'),'4');   bot.set(qn('w:color'),'7B1C28')
    pBdr.append(bot); pPr.append(pBdr)
    doc.add_paragraph()
    for line in text.split('\n'):
        p = doc.add_paragraph()
        r = p.add_run(line)
        r.font.size = Pt(11); r.font.name = "Noto Serif Malayalam"
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:line'),'360'); sp.set(qn('w:lineRule'),'auto')
        p._p.get_or_add_pPr().append(sp)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()


# ── Header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="app-header-icon">📋</div>
  <div>
    <h1>ഔദ്യോഗിക അപേക്ഷ / കത്ത് നിർമ്മാതാവ്</h1>
    <p>Kerala Govt · ഭരണമലയാളം · AI-powered · seconds-ൽ ready</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── API Key section ───────────────────────────────────────────────────────
if secret_key:
    active_key = secret_key
    st.markdown("""
    <div class="secret-badge">
      🔒 <span>API Key: <b>secrets.toml</b>-ൽ നിന്ന് load ചെയ്തു — ready!</span>
    </div>""", unsafe_allow_html=True)
else:
    active_key = ""
    with st.expander("🔑 Gemini API Key നൽകുക", expanded=True):
        col_k, col_m = st.columns([2, 1])
        manual_key = col_k.text_input(
            "API Key",
            type="password",
            placeholder="AIza...",
            label_visibility="collapsed"
        )
        model_sel = col_m.selectbox(
            "Model",
            ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"],
            label_visibility="collapsed"
        )
        if manual_key:
            active_key = manual_key
            st.success("✅ Key ready")
        else:
            st.caption(
                "💡 Permanent ആക്കാൻ: `.streamlit/secrets.toml` ഉണ്ടാക്കി "
                "`GEMINI_API_KEY = \"AIza...\"` add ചെയ്യുക"
            )

model_name = st.session_state.get('model_sel', 'gemini-2.5-flash')

# ── Form Card ─────────────────────────────────────────────────────────────
st.markdown('<div class="form-card">', unsafe_allow_html=True)

# Doc Type
st.markdown('<div class="form-section-title">📋 രേഖയുടെ തരം</div>', unsafe_allow_html=True)
group_sel = st.selectbox("Category", list(DOC_GROUPS.keys()), label_visibility="collapsed")
type_map  = DOC_GROUPS[group_sel]
doc_type  = st.selectbox(
    "Type", list(type_map.keys()),
    format_func=lambda x: type_map[x],
    label_visibility="collapsed"
)
is_app = doc_type in APP_TYPES_LIST
st.markdown("<hr>", unsafe_allow_html=True)

# Applicant section (apps only)
app_name = app_age = app_addr = app_phone = app_id = ""
if is_app:
    st.markdown('<div class="form-section-title">👤 അപേക്ഷകൻ</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    app_name  = c1.text_input("പേര് *", placeholder="രാജേഷ് കുമാർ")
    app_age   = c2.text_input("വയസ്സ്", placeholder="42")
    app_addr  = st.text_area("വിലാസം", placeholder="'ശ്രീനിലയം', ആനക്കയം P.O., നിലമ്പൂർ - 679329", height=70)
    c3, c4    = st.columns(2)
    app_phone = c3.text_input("ഫോൺ", placeholder="9876543210")
    app_id    = c4.text_input("Voter ID / Aadhaar", placeholder="ABC1234567")
    st.markdown("<hr>", unsafe_allow_html=True)

# From / Office
st.markdown(
    f'<div class="form-section-title">🏢 {"Authority / ഓഫീസ്" if is_app else "From — ഓഫീസ്"}</div>',
    unsafe_allow_html=True
)
c1, c2 = st.columns(2)
office_name = c1.text_input(
    "ഓഫീസ് പേര്",
    value=st.session_state.office_name,
    placeholder="നിലമ്പൂർ ഗ്രാമ പഞ്ചായത്ത്"
)
office_addr = c2.text_input(
    "വിലാസം",
    value=st.session_state.office_addr,
    placeholder="നിലമ്പൂർ, മലപ്പുറം - 679329"
)
if office_name: st.session_state.office_name = office_name
if office_addr: st.session_state.office_addr = office_addr

c3, c4 = st.columns(2)
file_no  = c3.text_input("ഫയൽ നം. / Ref", placeholder="B2-1234/2026")
date_val = c4.date_input("തീയതി", value=datetime.date.today())
st.markdown("<hr>", unsafe_allow_html=True)

# To
if doc_type not in ['note', 'order', 'sanction', 'memo']:
    to_label = {
        "rti_reply": "To — അപേക്ഷകൻ (പേര്/വിലാസം)",
    }.get(doc_type, "To — ആർക്ക് / Authority")
    st.markdown(f'<div class="form-section-title">📬 {to_label}</div>', unsafe_allow_html=True)
    to_whom = st.text_input(
        to_label,
        placeholder="സെക്രട്ടറി, നിലമ്പൂർ ഗ്രാമ പഞ്ചായത്ത്" if is_app else "അസി. എൻജിനീയർ, PWD",
        label_visibility="collapsed"
    )
    st.markdown("<hr>", unsafe_allow_html=True)
else:
    to_whom = ""

# Subject / Details
st.markdown('<div class="form-section-title">📝 Subject & Details</div>', unsafe_allow_html=True)
subject   = st.text_input("വിഷയം (Subject) *", placeholder="ഉദാ: വാർഡ് 12-ലെ റോഡ് അറ്റകുറ്റ പ്പണി സംബന്ധിച്ച്")
reference = st.text_input("സൂചന / Reference (ഐച്ഛികം)", placeholder="ഉദാ: ശ്രീ. XXX-ന്റെ കത്ത് dt. 01-06-2026")
pts_label = (
    "RTI ചോദ്യങ്ങൾ (നമ്പറിട്ട്) *" if doc_type == "rti_reply"
    else "ആവശ്യം / Details *" if is_app
    else "പ്രധാന കാര്യങ്ങൾ — rough notes *"
)
pts_ph = (
    "1. ചോദ്യം...\n2. ചോദ്യം..." if doc_type == "rti_reply"
    else "- ആവശ്യത്തിന്റെ കാരണം\n- Survey No., Ward, dates...\n- Attach ചെയ്ത documents" if is_app
    else "- പ്രശ്നം / ആവശ്യം\n- relevant facts\n- ആഗ്രഹിക്കുന്ന നടപടി"
)
points = st.text_area(pts_label, placeholder=pts_ph, height=130)
st.markdown("<hr>", unsafe_allow_html=True)

# Sign
sign_label = "✍️ അപേക്ഷകൻ" if is_app else "✍️ ഒപ്പ്"
st.markdown(f'<div class="form-section-title">{sign_label}</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
sign_name  = c1.text_input("പേര്", placeholder="കെ. രാജൻ")
sign_desig = c2.text_input(
    "പദവി" if not is_app else "തൊഴിൽ (ഐച്ഛികം)",
    placeholder="സെക്രട്ടറി" if not is_app else "കർഷകൻ"
)

st.markdown('</div>', unsafe_allow_html=True)  # /form-card

# Model select (shown only when no secret key)
if secret_key:
    model_name = st.selectbox(
        "Model",
        ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"],
        label_visibility="collapsed"
    )

# ── Generate Button ───────────────────────────────────────────────────────
btn_text = "⚡ അപേക്ഷ തയ്യാറാക്കുക" if is_app else "⚡ കത്ത് / രേഖ തയ്യാറാക്കുക"
gen_btn  = st.button(btn_text, type="primary", use_container_width=True)

if gen_btn:
    # Validation
    err = None
    if not active_key:
        err = "⚠️ API Key നൽകുക (മുകളിൽ)"
    elif not subject or not points:
        err = "⚠️ വിഷയവും Details-ഉം നിർബന്ധം"
    elif is_app and not app_name:
        err = "⚠️ അപേക്ഷകന്റെ പേര് നൽകുക"
    if err:
        st.error(err)
    else:
        data = dict(
            doc_type=doc_type, office_name=office_name, office_addr=office_addr,
            file_no=file_no, date_str=date_val.strftime('%d/%m/%Y'),
            to_whom=to_whom, subject=subject, reference=reference, points=points,
            sign_name=sign_name, sign_desig=sign_desig,
            app_name=app_name, app_age=app_age, app_addr=app_addr,
            app_phone=app_phone, app_id=app_id,
        )

        # ── AI Loading Animation ──────────────────────────────────────────
        loader_ph = st.empty()

        def show_loader(step_idx):
            steps_html = ""
            for i, (icon, txt) in enumerate(AI_STEPS):
                if i < step_idx:
                    cls = "done"; ico = "✅"
                elif i == step_idx:
                    cls = "active"; ico = icon
                else:
                    cls = ""; ico = icon
                steps_html += f'<div class="ai-step {cls}"><span class="step-icon">{ico}</span>{txt}</div>'
            loader_ph.markdown(f"""
<div class="ai-loader">
  <div class="ai-loader-title">AI തയ്യാറാക്കുന്നു...</div>
  <div class="ai-loader-sub">{ALL_TYPES.get(doc_type, '')} · ഭരണമലയാളം</div>
  <div class="typing-dots"><span></span><span></span><span></span></div>
  <div class="ai-steps">{steps_html}</div>
</div>
""", unsafe_allow_html=True)

        try:
            show_loader(0); time.sleep(0.6)
            show_loader(1); time.sleep(0.5)
            prompt = build_prompt(data)
            show_loader(2)
            result = call_gemini(active_key, model_name, prompt)
            show_loader(3); time.sleep(0.4)
            loader_ph.empty()
            st.session_state.output    = result
            st.session_state.edit_mode = False
            st.session_state.doc_label = ALL_TYPES.get(doc_type, '')
            st.session_state.docx_cache = make_docx(result, ALL_TYPES.get(doc_type, ''))
            st.rerun()
        except Exception as e:
            loader_ph.empty()
            st.error(f"❌ Error: {e}")

# ── Output ────────────────────────────────────────────────────────────────
if st.session_state.output:
    doc_label = st.session_state.doc_label or ALL_TYPES.get(doc_type, "")

    # Rebuild docx if edit mode changed the text
    if st.session_state.edit_mode:
        _docx_bytes = make_docx(st.session_state.output, doc_label)
    else:
        _docx_bytes = st.session_state.docx_cache or make_docx(st.session_state.output, doc_label)

    # Action bar
    c1, c2, c3, c4 = st.columns(4)
    if c1.button("✏️ Edit" if not st.session_state.edit_mode else "👁️ Preview", use_container_width=True):
        st.session_state.edit_mode = not st.session_state.edit_mode
        st.rerun()

    c2.download_button(
        "⬇️ .docx",
        data=_docx_bytes,
        file_name=f"letter-{datetime.date.today()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    if c3.button("🖨️ Print", use_container_width=True):
        st.session_state['do_print'] = True
    if c4.button("🗑️ Clear", use_container_width=True):
        st.session_state.output = ''
        st.session_state.edit_mode = False
        st.rerun()

    # Edit mode
    if st.session_state.edit_mode:
        st.caption("✏️ നേരിട്ട് edit ചെയ്യാം")
        edited = st.text_area(
            "edit",
            value=st.session_state.output,
            height=600,
            label_visibility="collapsed"
        )
        st.session_state.output = edited
    else:
        # Print trigger
        print_js = ""
        if st.session_state.get('do_print'):
            print_js = "<script>window.print();</script>"
            st.session_state['do_print'] = False

        safe = (st.session_state.output
                .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))

        st.components.v1.html(f"""
<!DOCTYPE html><html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:#f0ebe0;padding:12px;font-size:14px}}
.sheet{{
  background:white;border-radius:8px;
  padding:48px 52px;
  font-family:'Noto Serif Malayalam',serif;
  font-size:14px;line-height:2.1;color:#111;
  white-space:pre-wrap;
  box-shadow:0 4px 20px rgba(0,0,0,.1);
  position:relative;
}}
.sheet::before{{
  content:'';position:absolute;top:0;left:0;right:0;height:4px;
  background:linear-gradient(90deg,#7B1C28,#A9812F);
  border-radius:8px 8px 0 0;
}}
@media print{{
  body{{background:white;padding:0}}
  .sheet{{box-shadow:none;border:none;border-radius:0}}
  .sheet::before{{display:none}}
}}
</style></head>
<body><div class="sheet">{safe}</div>{print_js}</body></html>
""", height=680, scrolling=True)

        wc = len(st.session_state.output.split())
        st.caption(f"📊 {wc} words · {len(st.session_state.output)} chars")

# ── Footer ────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("🔒 API Key session-ൽ മാത്രം · Kerala Govt Document Tool")
