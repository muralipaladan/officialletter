import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, datetime, time

st.set_page_config(page_title="Letter Creator", page_icon="✍️",
                   layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&family=Noto+Sans+Malayalam:wght@400;500;600&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body, [data-testid="stAppViewContainer"] {
    background: #1C1C1E !important;
    color: #F0EBE0;
    font-family: 'Noto Sans Malayalam', sans-serif;
}
#MainMenu, footer, header, [data-testid="stToolbar"],
[data-testid="collapsedControl"] { display: none !important; }
.block-container { max-width: 720px !important; padding: 2rem 1.2rem 5rem !important; }

/* ── Top bar ── */
.topbar {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 28px;
}
.topbar-title { font-family: 'Noto Serif Malayalam', serif; font-size: 1.15rem; font-weight: 700; color: #F0EBE0; }
.topbar-sub   { font-size: .75rem; color: #888; margin-top: 2px; }
.model-pill {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    border-radius: 20px; padding: 5px 14px;
    font-size: .75rem; color: #A0A0A0;
    display: flex; align-items: center; gap: 6px;
}
.dot { width:7px; height:7px; border-radius:50%; background:#30D158; display:inline-block; }

/* ── Input card ── */
.input-card {
    background: #2C2C2E;
    border-radius: 16px;
    padding: 20px 22px;
    border: 1px solid #3A3A3C;
    margin-bottom: 16px;
    position: relative;
}
.input-label {
    font-size: .72rem; font-weight: 600; color: #888;
    letter-spacing: .08em; text-transform: uppercase;
    margin-bottom: 10px;
}
.hint-chips {
    display: flex; flex-wrap: wrap; gap: 7px; margin-top: 14px;
}
.hint-chip {
    background: #3A3A3C; border-radius: 20px;
    padding: 5px 13px; font-size: .76rem; color: #B0B0B0;
    cursor: pointer; border: 1px solid #48484A;
    white-space: nowrap;
}

/* Streamlit textarea override */
[data-testid="stTextArea"] textarea {
    background: transparent !important;
    border: none !important;
    border-radius: 0 !important;
    color: #F0EBE0 !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: 1rem !important;
    line-height: 1.75 !important;
    resize: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    caret-color: #0A84FF;
}
[data-testid="stTextArea"] textarea:focus {
    box-shadow: none !important;
    border: none !important;
}
[data-testid="stTextArea"] { border: none !important; }
[data-testid="stTextArea"] > div { border: none !important; background: transparent !important; }

/* Streamlit text_input override */
[data-testid="stTextInput"] input {
    background: #3A3A3C !important;
    border: 1px solid #48484A !important;
    border-radius: 8px !important;
    color: #F0EBE0 !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .9rem !important;
}
[data-testid="stTextInput"] input:focus {
    border-color: #0A84FF !important;
    box-shadow: 0 0 0 2px rgba(10,132,255,.2) !important;
}

/* selectbox */
[data-testid="stSelectbox"] > div > div {
    background: #3A3A3C !important;
    border: 1px solid #48484A !important;
    border-radius: 8px !important;
    color: #F0EBE0 !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .88rem !important;
}

/* Labels */
label, .stTextInput label, .stTextArea label,
.stSelectbox label { color: #888 !important; font-size: .78rem !important; }

/* ── Generate button ── */
[data-testid="stButton"] > button[kind="primary"] {
    background: #0A84FF !important;
    border: none !important;
    border-radius: 12px !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .95rem !important;
    font-weight: 600 !important;
    height: 50px !important;
    width: 100% !important;
    color: white !important;
    letter-spacing: .02em;
    box-shadow: 0 4px 20px rgba(10,132,255,.35) !important;
    transition: all .2s !important;
}
[data-testid="stButton"] > button[kind="primary"]:hover {
    background: #0070E0 !important;
    box-shadow: 0 6px 28px rgba(10,132,255,.5) !important;
    transform: translateY(-1px) !important;
}
[data-testid="stButton"] > button[kind="secondary"] {
    background: #3A3A3C !important;
    border: 1px solid #48484A !important;
    border-radius: 8px !important;
    color: #C0C0C0 !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .82rem !important;
}

/* ── Loader ── */
.ai-loader {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    border-radius: 16px; padding: 36px 24px;
    text-align: center; margin-bottom: 16px;
}
.loader-title { font-family:'Noto Serif Malayalam',serif; color:#F0EBE0; font-size:1rem; font-weight:600; margin-bottom:4px; }
.loader-sub   { color:#666; font-size:.78rem; margin-bottom:24px; }
.dots { display:inline-flex; gap:8px; margin-bottom:22px; }
.dots span {
    width:9px; height:9px; border-radius:50%; background:#0A84FF;
    display:inline-block; animation:bop 1.2s infinite ease-in-out;
}
.dots span:nth-child(2) { animation-delay:.2s; background:#30A0FF; }
.dots span:nth-child(3) { animation-delay:.4s; background:#60C0FF; }
@keyframes bop {
    0%,60%,100% { transform:translateY(0); opacity:.5; }
    30% { transform:translateY(-9px); opacity:1; }
}
.steps { display:flex; flex-direction:column; gap:8px; max-width:320px; margin:0 auto; text-align:left; }
.step {
    display:flex; align-items:center; gap:10px;
    padding:8px 14px; border-radius:8px;
    font-size:.83rem; color:#666;
    background:#232323; border:1px solid #333;
}
.step.active { color:#0A84FF; background:#0A1525; border-color:#0A84FF66; font-weight:600; }
.step.done   { color:#30D158; background:#0D1F12; border-color:#30D15866; }

/* ── Output sheet ── */
.out-header {
    display:flex; align-items:center; justify-content:space-between;
    margin-bottom: 12px;
}
.out-label {
    font-size:.72rem; font-weight:600; color:#888;
    letter-spacing:.08em; text-transform:uppercase;
}
.out-type-pill {
    background:#3A3A3C; border-radius:20px;
    padding:3px 12px; font-size:.72rem; color:#B0B0B0;
    border:1px solid #48484A;
}
.letter-sheet {
    background: #F8F4EC;
    border-radius: 12px;
    padding: 48px 52px;
    font-family: 'Noto Serif Malayalam', serif;
    font-size: 14px;
    line-height: 2.1;
    color: #111;
    white-space: pre-wrap;
    box-shadow: 0 8px 40px rgba(0,0,0,.5);
    position: relative;
    margin-bottom: 14px;
    border-top: 4px solid #0A84FF;
}

/* ── Action row ── */
.act-row { display:flex; gap:10px; margin-bottom:20px; flex-wrap:wrap; }

/* download btn */
[data-testid="stDownloadButton"] button {
    background: #30D158 !important;
    border: none !important;
    border-radius: 8px !important;
    color: #000 !important;
    font-family: 'Noto Sans Malayalam', sans-serif !important;
    font-size: .82rem !important;
    font-weight: 600 !important;
    padding: 0 16px !important;
}

/* secret badge */
.sec-badge {
    background:#0D1F12; border:1px solid #30D15866;
    border-radius:8px; padding:8px 14px;
    color:#30D158; font-size:.78rem;
    display:flex; align-items:center; gap:7px;
    margin-bottom:14px;
}

/* key expander */
[data-testid="stExpander"] {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    border-radius: 12px !important;
    margin-bottom: 16px !important;
}
[data-testid="stExpander"] summary { color: #B0B0B0 !important; font-size:.88rem !important; }

hr { border-color: #3A3A3C !important; }

@media print {
    html,body,[data-testid="stAppViewContainer"] { background:white !important; }
    .topbar,.input-card,[data-testid="stButton"],
    [data-testid="stDownloadButton"],.act-row,.sec-badge,
    [data-testid="stExpander"],.out-header { display:none !important; }
    .letter-sheet {
        box-shadow:none !important; border:none !important;
        padding:0 !important; border-radius:0 !important;
        background:white !important; color:#000 !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ── SYSTEM PROMPT ────────────────────────────────────────────────────────
SYSTEM_PROMPT = """നീ ഒരു expert Kerala government document writer ആണ്.
User type ചെയ്യുന്നത് — Malayalam, Manglish, English ഏതും ആകാം — നോക്കി:
1. ഏത് തരം രേഖ വേണം എന്ന് identify ചെയ്യുക (letter/application/order/circular/RTI/notice...)
2. From, To, Subject, Reference, ഉള്ളടക്കം, ഒപ്പ് — ഉള്ള വിവരങ്ങൾ use ചെയ്യുക
3. ഇല്ലാത്ത fields ന്യായമായ default ഇടുക (eg. തീയതി ഇന്നത്തേത്, ഒപ്പ് blank)
4. ശരിയായ ഭരണമലയാളം ശൈലിയിൽ complete official document തയ്യാറാക്കുക

Output rules:
- Document text ONLY — explanation, preamble, markdown symbols (** ## --) ഒന്നും വേണ്ട
- ശരിയായ line breaks ഉപയോഗിക്കുക
- ഏത് ഭാഷയിൽ input ആണെങ്കിലും output ഭരണമലയാളം ആയിരിക്കണം"""

MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]

AI_STEPS = [
    ("🔍", "ആവശ്യം മനസ്സിലാക്കുന്നു..."),
    ("📐", "രേഖയുടെ format തയ്യാറാക്കുന്നു..."),
    ("✍️", "ഭരണമലയാളത്തിൽ എഴുതുന്നു..."),
    ("✅", "തയ്യാറായി!"),
]

HINTS = [
    "income certificate application",
    "road repair letter to panchayat",
    "RTI application on ward works",
    "NOC for building construction",
    "pension application",
    "trade licence renewal",
    "residence certificate",
    "leave application",
]

# ── Session state ────────────────────────────────────────────────────────
for k, v in {
    "output": "", "doc_type_detected": "",
    "docx_cache": None, "edit_mode": False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── API key ──────────────────────────────────────────────────────────────
secret_key = st.secrets.get("GEMINI_API_KEY", "")

# ── Helpers ──────────────────────────────────────────────────────────────
def call_gemini(api_key, model_name, user_text):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name,
        system_instruction=SYSTEM_PROMPT
    )
    resp = model.generate_content(
        user_text,
        generation_config=genai.types.GenerationConfig(temperature=0.35)
    )
    return resp.text.strip()

def make_docx(text, label="Official Document"):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3);  sec.right_margin = Cm(2.5)
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(f"[ {label} ]")
    tr.font.size = Pt(8); tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x0A, 0x84, 0xFF)
    tr.font.name = "Noto Serif Malayalam"
    pPr = tp._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '4')
    bot.set(qn('w:space'), '4');   bot.set(qn('w:color'), '0A84FF')
    pBdr.append(bot); pPr.append(pBdr)
    doc.add_paragraph()
    for line in text.split('\n'):
        p = doc.add_paragraph()
        r = p.add_run(line)
        r.font.size = Pt(11); r.font.name = "Noto Serif Malayalam"
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')
        p._p.get_or_add_pPr().append(sp)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# ── Top bar ──────────────────────────────────────────────────────────────
model_name = MODELS[0]
if secret_key:
    st.markdown("""
    <div class="sec-badge">
      🔒 <span>API Key ready — secrets.toml</span>
    </div>""", unsafe_allow_html=True)
else:
    with st.expander("🔑 Gemini API Key", expanded=not bool(secret_key)):
        c1, c2 = st.columns([3, 2])
        manual_key = c1.text_input("Key", type="password",
            placeholder="AIza...", label_visibility="collapsed")
        model_sel  = c2.selectbox("Model", MODELS, label_visibility="collapsed")
        secret_key  = manual_key
        model_name  = model_sel
        if manual_key:
            st.success("✅ Ready")
        else:
            st.caption("💡 `.streamlit/secrets.toml` → `GEMINI_API_KEY = \"AIza...\"`")

if secret_key:
    model_name = st.selectbox("Model", MODELS, label_visibility="collapsed")

active_key = secret_key

# model pill
st.markdown(f"""
<div class="topbar">
  <div>
    <div class="topbar-title">Application / Office Letter Creator</div>
    <div class="topbar-sub">ഭരണഭാഷ &nbsp;·&nbsp; മാതൃഭാഷ</div>
  </div>
  <div class="model-pill"><span class="dot"></span>{model_name}</div>
</div>
""", unsafe_allow_html=True)

# ── Input card ───────────────────────────────────────────────────────────
st.markdown('<div class="input-card">', unsafe_allow_html=True)
st.markdown('<div class="input-label">ആവശ്യം ഇവിടെ type ചെയ്യൂ — Malayalam · Manglish · English</div>',
            unsafe_allow_html=True)

user_input = st.text_area(
    "input",
    height=200,
    placeholder=(
        "ഉദാ: നിലമ്പൂർ ഗ്രാമ പഞ്ചായത്ത് സെക്രട്ടറിക്ക് ഒരു income certificate application വേണം. "
        "എന്റെ പേര് രാജേഷ് കുമാർ, വിലാസം ആനക്കയം. scholarship-ന് വേണ്ടിയാണ്.\n\n"
        "or: Road repair letter to panchayat, ward 12 pothole, urgent\n\n"
        "or: RTI application — last year ward works estimate and contractor details"
    ),
    label_visibility="collapsed"
)

# Hint chips (clickable via query params workaround using button)
st.markdown('<div class="hint-chips">', unsafe_allow_html=True)
hint_cols = st.columns(4)
for i, hint in enumerate(HINTS[:4]):
    if hint_cols[i].button(hint, key=f"hint_{i}", use_container_width=True):
        st.session_state["_hint"] = hint
        st.rerun()
hint_cols2 = st.columns(4)
for i, hint in enumerate(HINTS[4:]):
    if hint_cols2[i].button(hint, key=f"hint2_{i}", use_container_width=True):
        st.session_state["_hint"] = hint
        st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# Apply hint
if "_hint" in st.session_state:
    user_input = st.session_state.pop("_hint")

st.markdown('</div>', unsafe_allow_html=True)  # /input-card

# ── Generate button ───────────────────────────────────────────────────────
gen = st.button("✦ Letter / Application തയ്യാറാക്കുക", type="primary", use_container_width=True)

if gen:
    if not active_key:
        st.error("⚠️ API Key നൽകുക")
    elif not user_input.strip():
        st.error("⚠️ ആവശ്യം type ചെയ്യൂ")
    else:
        loader = st.empty()
        def show(idx):
            rows = ""
            for i, (ic, tx) in enumerate(AI_STEPS):
                cls = "done" if i < idx else ("active" if i == idx else "")
                ico = "✅" if i < idx else ic
                rows += f'<div class="step {cls}"><span>{ico}</span>{tx}</div>'
            loader.markdown(f"""
<div class="ai-loader">
  <div class="loader-title">AI തയ്യാറാക്കുന്നു...</div>
  <div class="loader-sub">ഭരണമലയാളം · Official Format</div>
  <div class="dots"><span></span><span></span><span></span></div>
  <div class="steps">{rows}</div>
</div>""", unsafe_allow_html=True)

        try:
            show(0); time.sleep(0.5)
            show(1); time.sleep(0.4)
            show(2)
            result = call_gemini(active_key, model_name, user_input)
            show(3); time.sleep(0.35)
            loader.empty()
            st.session_state.output     = result
            st.session_state.edit_mode  = False
            st.session_state.docx_cache = make_docx(result)
            st.rerun()
        except Exception as e:
            loader.empty()
            st.error(f"❌ {e}")

# ── Output ────────────────────────────────────────────────────────────────
if st.session_state.output:
    st.markdown("---")

    # Action row
    c1, c2, c3, c4 = st.columns(4)
    if c1.button("✏️ Edit" if not st.session_state.edit_mode else "👁️ Preview",
                 use_container_width=True):
        st.session_state.edit_mode = not st.session_state.edit_mode
        st.rerun()

    # Rebuild docx if in edit mode (text may have changed)
    _docx = (make_docx(st.session_state.output)
             if st.session_state.edit_mode
             else st.session_state.docx_cache or make_docx(st.session_state.output))

    c2.download_button("⬇️ .docx", data=_docx,
        file_name=f"letter-{datetime.date.today()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)

    if c3.button("🖨️ Print", use_container_width=True):
        st.session_state["_print"] = True
    if c4.button("🗑️ Clear", use_container_width=True):
        st.session_state.output = ""; st.session_state.docx_cache = None
        st.session_state.edit_mode = False; st.rerun()

    if st.session_state.edit_mode:
        st.caption("✏️ നേരിട്ട് edit ചെയ്യാം")
        edited = st.text_area("edit_area", value=st.session_state.output,
                              height=600, label_visibility="collapsed")
        st.session_state.output = edited
    else:
        pjs = ""
        if st.session_state.get("_print"):
            pjs = "<script>window.print();</script>"
            st.session_state["_print"] = False

        safe = (st.session_state.output
                .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))

        st.components.v1.html(f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Malayalam:wght@400;600;700&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:#1C1C1E;padding:12px}}
.sheet{{
  background:#F8F4EC;border-radius:12px;
  padding:48px 52px;
  font-family:'Noto Serif Malayalam',serif;
  font-size:14px;line-height:2.1;color:#111;
  white-space:pre-wrap;
  box-shadow:0 8px 40px rgba(0,0,0,.6);
  border-top:4px solid #0A84FF;
}}
@media print{{
  body{{background:white;padding:0}}
  .sheet{{box-shadow:none;border:none;border-radius:0;background:white;border-top:none}}
}}
</style></head>
<body><div class="sheet">{safe}</div>{pjs}</body></html>
""", height=660, scrolling=True)

        wc = len(st.session_state.output.split())
        st.caption(f"📊 {wc} words · {len(st.session_state.output)} chars")

st.markdown("---")
st.caption("🔒 API Key session-ൽ മാത്രം · Kerala Govt Document Tool")
