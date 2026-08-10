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
[data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] > div > div {
    border-radius: 6px !important;
    border-color: #ddd !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .92rem !important;
}
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

# ── Session State ────────────────────────────────────────────────────────
for k, v in {
    'output': '', 'edit_mode': False,
    'generating': False, 'doc_label': '',
    'docx_cache': None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── API Key: secrets ─────────────────────────────────────────────────────
secret_key = st.secrets.get("GEMINI_API_KEY", "")

# ── Functions ────────────────────────────────────────────────────────────
def build_prompt(doc_type_key, user_text):
    doc_label = ALL_TYPES[doc_type_key]
    format_guide = FORMAT_GUIDES[doc_type_key]
    
    return f"""നീ Kerala സർക്കാർ/തദ്ദേശ ഓഫീസ് ഫയൽ എഴുത്തിൽ വിദഗ്ദ്ധനായ ഒരു Assistant ആണ്.
താഴെ നൽകിയിരിക്കുന്ന വിവരങ്ങൾ ഉപയോഗിച്ച് ഔദ്യോഗിക ഭരണമലയാള ശൈലിയിൽ ഒരു '{doc_label}' തയ്യാറാക്കുക.

Format നിർദ്ദേശം:
{format_guide}
- ലളിതവും സ്പഷ്ടവുമായ ഭരണമലയാളം ഉപയോഗിക്കുക.
- ഔട്ട്പുട്ടിൽ തയ്യാറാക്കിയ കത്ത്/അപേക്ഷ മാത്രമേ ഉണ്ടാകാവൂ. ** ## -- തുടങ്ങിയ Markdown ചിഹ്നങ്ങൾ ഉപയോഗിക്കരുത്.
- ശരിയായ വരി ഇടവേളകൾ (line breaks) ഉപയോഗിക്കുക.

ഉപയോക്താവ് നൽകിയ വിവരങ്ങൾ (ഇതിൽ നിന്ന് അയക്കുന്ന ആൾ, സ്വീകർത്താവ്, വിഷയം, മറ്റ് കാര്യങ്ങൾ എന്നിവ വേർതിരിച്ചെടുക്കുക):
{user_text}"""


def call_gemini(api_key, model_name, prompt):
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.35
        )
    )
    return response.text.strip()


def make_docx(text, label):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3);  sec.right_margin = Cm(2.5)
    
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
    <h1>Single Window AI Creator</h1>
    <p>ഭരണഭാഷ &nbsp;·&nbsp; മാതൃഭാഷ</p>
  </div>
</div>
""", unsafe_allow_html=True)

if secret_key:
    st.markdown("""
    <div class="secret-badge">
      🔒 <span>API Key: <b>secrets.toml</b>-ൽ നിന്ന് load ചെയ്തു — ready!</span>
    </div>""", unsafe_allow_html=True)
else:
    st.error("⚠️ Streamlit secrets.toml-ൽ API Key നൽകിയിട്ടില്ല.")

# ── Form (Single Window) ─────────────────────────────────────────────────
st.markdown('<div class="form-card">', unsafe_allow_html=True)

c1, c2 = st.columns([1, 1])
group_sel = c1.selectbox("രേഖയുടെ വിഭാഗം", list(DOC_GROUPS.keys()))
type_map  = DOC_GROUPS[group_sel]
doc_type  = c2.selectbox("രേഖയുടെ തരം", list(type_map.keys()), format_func=lambda x: type_map[x])

st.markdown('<div class="form-section-title">📝 വിവരങ്ങൾ നൽകുക</div>', unsafe_allow_html=True)
user_input = st.text_area(
    "എല്ലാ വിവരങ്ങളും ഇവിടെ നൽകുക (Single Window)",
    placeholder="ഉദാ: എന്റെ പേര് രാജേഷ്, വാർഡ് 12 ലെ കുടിവെള്ള കണക്ഷൻ ലഭിക്കാൻ പഞ്ചായത്ത് സെക്രട്ടറിക്ക് ഒരു അപേക്ഷ തയ്യാറാക്കണം. എന്റെ ഫോൺ നമ്പർ: 9876543210...",
    height=200,
    label_visibility="collapsed"
)

# മോഡലുകൾ നിലവിൽ ലഭ്യമായവ മാത്രം ഉൾപ്പെടുത്തി തിരുത്തിയിരിക്കുന്നു
model_name = st.selectbox(
    "AI Model",
    ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash", "gemini-1.5-pro"],
    index=0
)

st.markdown('</div>', unsafe_allow_html=True)

# ── Generate Button ───────────────────────────────────────────────────────
gen_btn = st.button("⚡ രേഖ തയ്യാറാക്കുക", type="primary", use_container_width=True)

if gen_btn:
    if not secret_key:
        st.error("⚠️ API Key ലഭ്യമല്ല.")
    elif not user_input.strip():
        st.error("⚠️ ദയവായി വിവരങ്ങൾ ടെക്സ്റ്റ് ബോക്സിൽ നൽകുക.")
    else:
        with st.spinner("AI രേഖ തയ്യാറാക്കുന്നു..."):
            try:
                prompt = build_prompt(doc_type, user_input)
                result = call_gemini(secret_key, model_name, prompt)
                
                st.session_state.output = result
                st.session_state.edit_mode = False
                st.session_state.doc_label = ALL_TYPES.get(doc_type, '')
                st.session_state.docx_cache = make_docx(result, ALL_TYPES.get(doc_type, ''))
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ── Output ────────────────────────────────────────────────────────────────
if st.session_state.output:
    doc_label = st.session_state.doc_label or ALL_TYPES.get(doc_type, "")

    if st.session_state.edit_mode:
        _docx_bytes = make_docx(st.session_state.output, doc_label)
    else:
        _docx_bytes = st.session_state.docx_cache or make_docx(st.session_state.output, doc_label)

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

    if st.session_state.edit_mode:
        st.caption("✏️ നേരിട്ട് edit ചെയ്യാം")
        edited = st.text_area(
            "edit",
            value=st.session_state.output,
            height=500,
            label_visibility="collapsed"
        )
        st.session_state.output = edited
    else:
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
""", height=650, scrolling=True)

        wc = len(st.session_state.output.split())
        st.caption(f"📊 {wc} words · {len(st.session_state.output)} chars")

st.markdown("---")
st.caption("🔒 API Key session-ൽ മാത്രം · Single Window Document Creator")
