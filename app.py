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
    page_icon="🤖",
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

#MainMenu, footer, header { visibility: hidden; }
[data-testid="stToolbar"] { display: none; }

.block-container {
    max-width: 780px !important;
    padding: 2rem 1.5rem 4rem !important;
}

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
}
</style>
""", unsafe_allow_html=True)

# ── Session State ────────────────────────────────────────────────────────
for k, v in {
    'output': '', 'edit_mode': False,
    'docx_cache': None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── API Key: secrets ─────────────────────────────────────────────────────
secret_key = st.secrets.get("GEMINI_API_KEY", "")

# ── Master Prompt Function ───────────────────────────────────────────────
def build_smart_prompt(preshitan, sweekarthavu, details):
    return f"""നീ കേരള സർക്കാരിന്റെ ഔദ്യോഗിക ഫയലുകൾ, അപേക്ഷകൾ, വിവരാവകാശ രേഖകൾ (RTI) എന്നിവ തയ്യാറാക്കുന്നതിൽ അഗാധമായ അറിവുള്ള ഒരു 'Senior Section Officer' ആണ്. 
താഴെ നൽകിയിരിക്കുന്ന വിവരങ്ങൾ വായിച്ച് മനസ്സിലാക്കി, ഉപയോക്താവിന് വേണ്ടത് ഒരു 'അപേക്ഷ' (Application) ആണോ, 'വിവരാവകാശ രേഖ' (RTI) ആണോ, അതോ ഒരു 'ഔദ്യോഗിക കത്ത്' (Official Letter) ആണോ എന്ന് നീ സ്വയം തിരിച്ചറിയുക.

ഏറ്റവും മികച്ച, നിയമപരമായി ശരിയായ ഔദ്യോഗിക ഭരണമലയാള (Official Malayalam) ശൈലിയിൽ ഈ രേഖ തയ്യാറാക്കുക.

ശ്രദ്ധിക്കേണ്ട കർശന നിയമങ്ങൾ:
1. ഘടന:
   - പ്രേഷിതൻ (From) മുകളിലും, സ്വീകർത്താവ് (To) അതിനു താഴെയും കൃത്യമായി വരണം.
   - സംബോധന: 'സർ,' അല്ലെങ്കിൽ 'മാന്യരേ,' എന്ന് ഉപയോഗിക്കുക.
   - വിഷയം (Subject): കത്തിന്റെ ഉള്ളടക്കം ഒറ്റ വരിയിൽ കൃത്യമായി പറയുക.
   - ഉള്ളടക്കം (Body): കാര്യം വളച്ചുകെട്ടില്ലാതെ, വ്യക്തമായി പാരഗ്രാഫുകളായി എഴുതുക. 'അപേക്ഷിക്കുന്നു', 'ശ്രദ്ധയിൽപ്പെടുത്തുന്നു', 'കനിവുണ്ടാകണമെന്ന് അഭ്യർത്ഥിക്കുന്നു' തുടങ്ങിയ മാന്യമായ ഔദ്യോഗിക പദങ്ങൾ ഉപയോഗിക്കുക.
2. വിവരാവകാശ അപേക്ഷയാണെങ്കിൽ (RTI):
   - തലക്കെട്ട്: "വിവരാവകാശ നിയമം 2005 പ്രകാരമുള്ള അപേക്ഷ" എന്ന് നൽകുക.
   - സ്വീകർത്താവ്: 'സ്റ്റേറ്റ് പബ്ലിക് ഇൻഫർമേഷൻ ഓഫീസർ' (SPIO) എന്ന് ഉപയോഗിക്കുക.
   - ഫീസ്: "നിയമപ്രകാരമുള്ള അപേക്ഷാ ഫീസായ 10 രൂപയുടെ കോർട്ട് ഫീ സ്റ്റാമ്പ് ഇതിനോടൊപ്പം പതിച്ചിട്ടുണ്ട്" എന്ന് ചേർക്കുക. ചോദ്യങ്ങൾ 1, 2, 3 എന്ന് നമ്പറിട്ട് നൽകുക.
3. ഉപസംഹാരം: 'വിശ്വസ്തതയോടെ,' എന്ന് നൽകി ഒപ്പിടാനുള്ള സ്ഥലം നൽകുക. ഔട്ട്പുട്ടിൽ തയ്യാറാക്കിയ പൂർണ്ണമായ രേഖ മാത്രമേ ഉണ്ടാകാവൂ. Markdown ചിഹ്നങ്ങളോ അനാവശ്യ വിവരണങ്ങളോ പാടില്ല.

ഉപയോക്താവ് നൽകിയ വിവരങ്ങൾ:
--------------------------
പ്രേഷിതൻ (Preshitan): {preshitan}
സ്വീകർത്താവ് (Sweekarthavu): {sweekarthavu}
വിഷയവും മറ്റ് വിവരങ്ങളും: {details}
--------------------------
"""

def call_gemini(api_key, model_name, prompt):
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.3
        )
    )
    return response.text.strip()

def make_docx(text):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3);  sec.right_margin = Cm(2.5)
    
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("[ ഔദ്യോഗിക രേഖ ]")
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


# ── Header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="app-header-icon">🤖</div>
  <div>
    <h1>Smart Document Creator</h1>
    <p>പ്രേഷിതൻ & സ്വീകർത്താവ് &nbsp;·&nbsp; അപേക്ഷ, RTI, ഔദ്യോഗിക കത്ത്</p>
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

# ── Form ────────────────────────────────────────────────_________________
st.markdown('<div class="form-card">', unsafe_allow_html=True)
st.markdown('<div class="form-section-title">📝 വിവരങ്ങൾ നൽകുക</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    preshitan = st.text_area("പ്രേഷിതൻ (Preshitan)", placeholder="പേര്\nവിലാസം\nഫോൺ നമ്പർ", height=100)
with col2:
    sweekarthavu = st.text_area("സ്വീകർത്താവ് (Sweekarthavu)", placeholder="പദവി\nഓഫീസിന്റെ പേര്\nസ്ഥലം", height=100)

ph_text = """ഉദാഹരണങ്ങൾ:
1. (RTI) പഞ്ചായത്തിൽ കഴിഞ്ഞ വർഷം റോഡ് പണിക്കായി അനുവദിച്ച ഫണ്ട് എത്രയെന്ന് അറിയാൻ...
2. (അപേക്ഷ) വീടിന്റെ പണി കഴിഞ്ഞു, പുതിയ കുടിവെള്ള കണക്ഷൻ ലഭിക്കാൻ...
3. (കത്ത്) ഓഫീസിൽ ഫയൽ നീക്കം എളുപ്പമാക്കാൻ പുതിയ 2 കമ്പ്യൂട്ടറുകൾ അനുവദിക്കുന്നത് സംബന്ധിച്ച്..."""

details = st.text_area(
    "വിഷയവും മറ്റ് വിവരങ്ങളും (Subject & Details)",
    placeholder=ph_text,
    height=160
)

# AI Model Selection
model_name = st.selectbox(
    "AI Model",
    ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-1.5-flash"],
    index=0
)

st.markdown('</div>', unsafe_allow_html=True)

# ── Generate Button ───────────────────────────────────────────────────────
gen_btn = st.button("⚡ മികച്ച ഔദ്യോഗിക രേഖ തയ്യാറാക്കുക", type="primary", use_container_width=True)

if gen_btn:
    if not secret_key:
        st.error("⚠️ API Key ലഭ്യമല്ല.")
    elif not details.strip():
        st.error("⚠️ ദയവായി വിഷയവും വിവരങ്ങളും നൽകുക.")
    else:
        with st.spinner("AI രേഖ തയ്യാറാക്കുന്നു..."):
            try:
                prompt = build_smart_prompt(preshitan, sweekarthavu, details)
                result = call_gemini(secret_key, model_name, prompt)
                
                st.session_state.output = result
                st.session_state.edit_mode = False
                st.session_state.docx_cache = make_docx(result)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ── Output ────────────────────────────────────────────────────────────────
if st.session_state.output:
    if st.session_state.edit_mode:
        _docx_bytes = make_docx(st.session_state.output)
    else:
        _docx_bytes = st.session_state.docx_cache or make_docx(st.session_state.output)

    c1, c2, c3, c4 = st.columns(4)
    if c1.button("✏️ Edit" if not st.session_state.edit_mode else "👁️ Preview", use_container_width=True):
        st.session_state.edit_mode = not st.session_state.edit_mode
        st.rerun()

    c2.download_button(
        "⬇️ .docx",
        data=_docx_bytes,
        file_name=f"official-document-{datetime.date.today()}.docx",
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
        st.caption(f"📊 {wc} words · {len(st.session_state.output)} characters")

st.markdown("---")
st.caption("🔒 Advanced AI Auto-Detect Mode · Kerala Govt Formats Trained")
