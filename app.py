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
    page_title="Smart Document Creator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&family=Noto+Sans+Malayalam:wght@400;500;600&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #f0ebe0 !important;
    font-family: 'Noto Sans Malayalam', sans-serif;
}

/* Hide streamlit default chrome */
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stToolbar"] { display: none; }

.block-container {
    max-width: 800px !important;
    padding: 2rem 1.5rem 4rem !important;
}

/* App header */
.app-header {
    background: linear-gradient(135deg, #1A4D2E 0%, #0F2C1A 100%);
    border-radius: 10px;
    padding: 22px 28px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 4px 20px rgba(26,77,46,.25);
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
    border-top: 3px solid #1A4D2E;
}
.form-section-title {
    font-family: 'Noto Serif Malayalam', serif;
    font-size: .95rem;
    font-weight: 700;
    color: #1A4D2E;
    margin-bottom: 14px;
    padding-bottom: 8px;
    border-bottom: 1px solid #f0e8e8;
    display: flex; align-items: center; gap: 8px;
}

/* Streamlit input overrides */
[data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] > div > div,
[data-testid="stTextInput"] input {
    border-radius: 6px !important;
    border-color: #ddd !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .92rem !important;
}
[data-testid="stTextArea"] textarea:focus,
[data-testid="stTextInput"] input:focus {
    border-color: #1A4D2E !important;
    box-shadow: 0 0 0 2px rgba(26,77,46,.12) !important;
}

/* Generate button */
[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #1A4D2E, #2A7347) !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    height: 52px !important;
    box-shadow: 0 4px 14px rgba(26,77,46,.3) !important;
}

/* Sidebar Customization */
[data-testid="stSidebar"] {
    background-color: #11301c;
}
[data-testid="stSidebar"] * {
    color: white !important;
}
.sidebar-info {
    font-size: 0.85rem;
    line-height: 1.6;
    color: #d1e8d5 !important;
    background: rgba(0,0,0,0.2);
    padding: 15px;
    border-radius: 8px;
    margin-top: 20px;
}
.sidebar-info a {
    color: #81c784 !important;
}

@media print {
    .app-header, .form-card, [data-testid="stButton"],
    .action-bar, [data-testid="stAlert"], [data-testid="stSidebar"] { display: none !important; }
}
</style>
""", unsafe_allow_html=True)

# ── Constants ────────────────────────────────────────────────────────────
DOC_GROUPS = {
    "📨 കത്തുകൾ (Letters)": {
        "letter":     "ഔദ്യോഗിക കത്ത് (Official Letter)",
        "do_letter":  "അർദ്ധ-ഔദ്യോഗിക കത്ത് (D.O. Letter)",
        "forwarding": "അയക്കൽ കത്ത് (Forwarding)",
        "reminder":   "ഓർമ്മപ്പെടുത്തൽ (Reminder)",
        "invitation": "ക്ഷണക്കത്ത് (Invitation)",
    },
    "📜 ഉത്തരവ് / അറിയിപ്പ് (Orders)": {
        "order":        "ഉത്തരവ് (Order)",
        "sanction":     "അനുമതി ഉത്തരവ് (Sanction)",
        "circular":     "സർക്കുലർ (Circular)",
        "public_notice":"പൊതു അറിയിപ്പ് (Notice)",
        "show_cause":   "കാരണം കാണിക്കൽ നോട്ടീസ്",
        "rti_reply":    "വിവരാവകാശ രേഖ (RTI)",
    },
    "📝 അപേക്ഷകൾ (Applications)": {
        "app_general":    "പൊതു അപേക്ഷ (General Application)",
        "app_income":     "വരുമാന Certificate അപേക്ഷ",
        "app_nativity":   "ജനന/നാട്ടുകാർ Certificate",
        "app_residence":  "താമസ Certificate",
        "app_caste":      "ജാതി Certificate",
        "app_noc":        "NOC അപേക്ഷ",
        "app_building":   "കെട്ടിട അനുമതി",
        "app_trade":      "വ്യാപാര ലൈസൻസ്",
        "app_pension":    "പെൻഷൻ/ആനുകൂല്യം",
        "app_leave":      "അവധി അപേക്ഷ (Leave Letter)",
        "app_complaint":  "പരാതി (Complaint)",
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
    "rti_reply":   "RTI Act 2005 പ്രകാരമുള്ള ശരിയായ ഫോർമാറ്റ്",
    "app_general": "'മഹോദയ/മഹോദയേ,' → അപേക്ഷകൻ intro → ആവശ്യം/കാരണം → request → 'അപേക്ഷകൻ' ഒപ്പ്/വിലാസം/തീയതി",
    "app_income":  "വരുമാന source/തുക, ഉദ്ദേശ്യം, village officer verify ആവശ്യം, സത്യസന്ധ declaration",
    "app_nativity":"ജനനം/residence confirm, ഉദ്ദേശ്യം, ജനനതീയതി/ജന്മഗ്രാമം",
    "app_residence":"X വർഷം/തീയതി മുതൽ residence, ഉദ്ദേശ്യം, ID proof reference",
    "app_caste":   "ജാതി/community/list (SC/ST/OBC), ഉദ്ദേശ്യം, Tahsildar verify ആവശ്യം",
    "app_noc":     "ഏത് ആവശ്യം/activity/location, objection ഇല്ലെന്ന് confirm ആവശ്യം",
    "app_building":"Plot/Survey No., ഉദ്ദേശ്യം, floor area, Building Rules compliance, permit request",
    "app_trade":   "ബിസിനസ് പേര്/trade/Ward, owner, NOC ready, license/renewal",
    "app_pension": "ഏത് scheme, അർഹത, proof, bank account, direct transfer request",
    "app_complaint":"ആര്/സംഭവം/തീയതി, prior complaint, ആഗ്രഹിക്കുന്ന remedy – objective",
    "app_leave":   "Designation/office, leave type/dates, കാരണം, alternate arrangement",
}

# ── Session State ────────────────────────────────────────────────────────
for k, v in {
    'output': '', 'edit_mode': False,
    'doc_label': '', 'docx_cache': None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Sidebar: API Key Setup ───────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔑 API Key Setup")
    
    # Try to load from secrets if available
    secret_key = st.secrets.get("GEMINI_API_KEY", "")
    
    user_api_key = st.text_input(
        "Google Gemini API Key", 
        type="password", 
        value=secret_key,
        placeholder="AIzaSy..."
    )
    
    active_api_key = user_api_key or secret_key

    if active_api_key:
        st.success("✅ API Key Ready!")
    
    st.markdown("""
    <div class="sidebar-info">
        <h3 style="margin-bottom: 10px; color: #fff;">❓ API Key എങ്ങനെ ലഭിക്കും?</h3>
        <ol style="padding-left: 15px; margin-bottom: 0;">
            <li><a href="https://aistudio.google.com/app/apikey" target="_blank">Google AI Studio</a> സന്ദർശിക്കുക.</li>
            <li>Google Account വഴി ലോഗിൻ ചെയ്യുക.</li>
            <li>"Create API Key" ക്ലിക്ക് ചെയ്യുക.</li>
            <li>ലഭിക്കുന്ന Key കോപ്പി ചെയ്ത് മുകളിലെ ബോക്സിൽ നൽകുക.</li>
        </ol>
    </div>
    
    <div class="sidebar-info">
        <h3 style="margin-bottom: 10px; color: #fff;">📄 ഈ സൈറ്റിലെ സേവനങ്ങൾ</h3>
        <ul style="padding-left: 15px; margin-bottom: 0;">
            <li>സർക്കാർ ഓഫീസുകളിലേക്കുള്ള അപേക്ഷകൾ</li>
            <li>വിവരാവകാശ രേഖകൾ (RTI)</li>
            <li>ഔദ്യോഗിക കത്തുകൾ</li>
            <li>പരാതികൾ, അവധി അപേക്ഷകൾ</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

# ── Functions ────────────────────────────────────────────────────────────
def build_prompt(doc_type_key, from_addr, to_addr, details, language):
    doc_label = ALL_TYPES[doc_type_key]
    format_guide = FORMAT_GUIDES.get(doc_type_key, "")
    
    if language == "English":
        system_role = "You are an expert AI Assistant specializing in drafting official government and professional documents."
        lang_instruction = "professional and official English"
        strict_rule = "Ensure the document is strictly in English."
    else:
        system_role = "നീ Kerala സർക്കാർ/തദ്ദേശ ഓഫീസ് ഫയൽ എഴുത്തിൽ വിദഗ്ദ്ധനായ ഒരു Assistant ആണ്."
        lang_instruction = "ഔദ്യോഗിക ഭരണമലയാള ശൈലിയിൽ (Official Malayalam)"
        strict_rule = "രേഖ പൂർണ്ണമായും മലയാളത്തിൽ ആയിരിക്കണം."

    return f"""{system_role}
Based on the provided information, draft a complete '{doc_label}' in {lang_instruction}.

Format Guide:
{format_guide}

Strict Instructions:
- {strict_rule}
- The output must strictly be the final document/letter ONLY. Do not include any introductory or concluding conversational text.
- Do not use Markdown symbols like **, ##, etc.
- Use proper line breaks (formatting) suitable for an official document.

Information Provided:
From Address: {from_addr}
To Address: {to_addr}
Subject & Details:
{details}"""


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
    tr.font.color.rgb = RGBColor(0x1A, 0x4D, 0x2E)
    tr.font.name = "Noto Serif Malayalam"
    pPr = tp._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'4')
    bot.set(qn('w:space'),'4');   bot.set(qn('w:color'),'1A4D2E')
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


# ── Main UI Header ───────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="app-header-icon">📋</div>
  <div>
    <h1>Smart Document Creator</h1>
    <p>English & Malayalam Support &nbsp;·&nbsp; AI Powered</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Form Input Section ───────────────────────────────────────────────────
st.markdown('<div class="form-card">', unsafe_allow_html=True)

# 1. Document Type & Language Selection
st.markdown('<div class="form-section-title">📋 രേഖയുടെ തരം & ഭാഷ (Select Document & Language)</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([1.5, 1.5, 1])
with c1:
    group_sel = st.selectbox("വിഭാഗം (Category)", list(DOC_GROUPS.keys()), label_visibility="collapsed")
with c2:
    type_map  = DOC_GROUPS[group_sel]
    doc_type  = st.selectbox("രേഖയുടെ തരം (Type)", list(type_map.keys()), format_func=lambda x: type_map[x], label_visibility="collapsed")
with c3:
    doc_language = st.selectbox("ഭാഷ (Language)", ["Malayalam", "English"], label_visibility="collapsed")

# 2. Details Input Windows
st.markdown('<div class="form-section-title" style="margin-top: 20px;">📝 വിവരങ്ങൾ നൽകുക (Enter Details)</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    from_addr = st.text_area("അയക്കുന്ന ആൾ (From Address)", placeholder="പേര്\nവിലാസം\nഫോൺ നമ്പർ", height=100)
with col2:
    to_addr = st.text_area("സ്വീകർത്താവ് (To Address)", placeholder="പദവി\nഓഫീസിന്റെ പേര്\nസ്ഥലം", height=100)

# Dynamic Placeholders based on selected doc type
if group_sel == "📝 അപേക്ഷകൾ (Applications)":
    ph_text = "ഉദാ: \nവിഷയം: കുടിവെള്ള കണക്ഷൻ ലഭിക്കുന്നത് സംബന്ധിച്ച്.\n\nവിവരങ്ങൾ: എന്റെ വീടിന്റെ പണി പൂർത്തിയായി. പുതിയ കുടിവെള്ള കണക്ഷൻ ലഭിക്കാൻ ആവശ്യമായ രേഖകൾ ഇതോടൊപ്പം സമർപ്പിക്കുന്നു. എത്രയും വേഗം നടപടി സ്വീകരിക്കണം..."
elif group_sel == "📨 കത്തുകൾ (Letters)":
    if doc_type == "do_letter":
        ph_text = "ഉദാ: \nപ്രിയപ്പെട്ട സുനിൽ,\nപഞ്ചായത്തിലെ മാലിന്യ സംസ്കരണ പദ്ധതിയുമായി ബന്ധപ്പെട്ട് അടുത്തയാഴ്ച നടക്കുന്ന യോഗത്തിൽ പങ്കെടുക്കണമെന്ന് അഭ്യർത്ഥിക്കുന്നു..."
    else:
        ph_text = "ഉദാ: \nവിഷയം: പുതിയ കമ്പ്യൂട്ടറുകൾ അനുവദിക്കുന്നത് സംബന്ധിച്ച്.\n\nവിവരങ്ങൾ: ഓഫീസിലെ പഴയ 2 കമ്പ്യൂട്ടറുകൾ കേടായതിനാൽ ഫയൽ നീക്കം തടസ്സപ്പെടുന്നുണ്ട്. അതിനാൽ പകരം 2 പുതിയ കമ്പ്യൂട്ടറുകൾ അനുവദിക്കണമെന്ന് അഭ്യർത്ഥിക്കുന്നു..."
elif doc_type == "rti_reply":
    ph_text = "ഉദാ: \n1. പഞ്ചായത്തിൽ കഴിഞ്ഞ സാമ്പത്തിക വർഷം റോഡ് പണിക്കായി എത്ര രൂപ ഫണ്ട് അനുവദിച്ചു?\n2. ഇതിൽ എത്ര രൂപ ചെലവാക്കി?\n3. ബാക്കി തുകയുടെ വിശദാംശങ്ങൾ നൽകുക."
else:
    ph_text = "കത്തിലോ ഉത്തരവിലോ ഉൾപ്പെടുത്തേണ്ട വിഷയവും മറ്റു പ്രധാന വിവരങ്ങളും ഇവിടെ ടൈപ്പ് ചെയ്യുക..."

details = st.text_area(
    "വിഷയവും മറ്റ് വിവരങ്ങളും (Subject & Details)",
    placeholder=ph_text,
    height=150
)

# AI Model Selection
model_name = st.selectbox(
    "AI Model",
    ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.5-pro"],
    index=0
)

st.markdown('</div>', unsafe_allow_html=True)

# ── Generate Button ───────────────────────────────────────────────────────
gen_btn = st.button(f"⚡ രേഖ തയ്യാറാക്കുക ({doc_language})", type="primary", use_container_width=True)

if gen_btn:
    if not active_api_key:
        st.error("⚠️ ദയവായി ഇടത് വശത്ത് (Sidebar) API Key നൽകുക.")
    elif not details.strip():
        st.error("⚠️ ദയവായി വിഷയവും വിവരങ്ങളും നൽകുക.")
    else:
        with st.spinner(f"AI ({ALL_TYPES[doc_type]}) തയ്യാറാക്കുന്നു..."):
            try:
                prompt = build_prompt(doc_type, from_addr, to_addr, details, doc_language)
                result = call_gemini(active_api_key, model_name, prompt)
                
                final_label = ALL_TYPES.get(doc_type, '') if doc_language == "Malayalam" else "Official Document"

                st.session_state.output = result
                st.session_state.edit_mode = False
                st.session_state.doc_label = final_label
                st.session_state.docx_cache = make_docx(result, final_label)
                st.rerun()
            except Exception as e:
                error_msg = str(e)
                if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                    st.error("⚠️ API ലിമിറ്റ് കഴിഞ്ഞിരിക്കുന്നു. ദയവായി കുറച്ചുസമയം കാത്തിരിക്കുക അല്ലെങ്കിൽ 'Flash' മോഡലുകൾ ഉപയോഗിക്കുക.")
                else:
                    st.error(f"❌ Error: {e}")

# ── Output Section ────────────────────────────────────────────────────────
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
        file_name=f"document-{datetime.date.today()}.docx",
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
  background:linear-gradient(90deg,#1A4D2E,#2A7347);
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
