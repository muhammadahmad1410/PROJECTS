import streamlit as st
from sentence_transformers import SentenceTransformer, util
from pdfminer.high_level import extract_text
from docx import Document
import re
import io
import html
import pandas as pd
from datetime import datetime

USER_NAME = "Muhammad Ahmad"   # shown in the top-right profile chip

st.set_page_config(
    page_title="AI Resume Matching System",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# SESSION STATE
# ============================================================
for key, default in {
    "analysis": None,
    "history": [],
    "page": "Analyzer",
    "theme": "Light",
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
# COLORS / THEME
# ============================================================
def colors():
    if st.session_state.theme == "Dark":
        return {
            "BG": "#0b1220", "SIDEBAR": "#0d1830", "CARD": "#111c30",
            "CARD2": "#16243b", "TEXT": "#f7f9ff", "MUTED": "#9ca9bd",
            "BORDER": "#263754", "PRIMARY": "#2f7df6", "PRIMARY2": "#5a8cff",
            "GREEN": "#19b66a", "RED": "#e85a63", "YELLOW": "#f0a23a",
            "PURPLE": "#9b6bff", "TRACK": "#1d3a2f", "HERO": "#10281f",
            "SG": "#12382b", "SR": "#3a1c24", "SB": "#14294a", "SY": "#3b2d14",
        }
    return {
        "BG": "#f5f8fc", "SIDEBAR": "#0b1730", "CARD": "#ffffff",
        "CARD2": "#f8fbff", "TEXT": "#14213d", "MUTED": "#64748b",
        "BORDER": "#dfe7f2", "PRIMARY": "#2675ee", "PRIMARY2": "#4c8dff",
        "GREEN": "#17b66a", "RED": "#e85a63", "YELLOW": "#f2a33b",
        "PURPLE": "#9b6bff", "TRACK": "#d7efe3", "HERO": "#f1fbf6",
        "SG": "#e8f8f0", "SR": "#ffebed", "SB": "#eaf3ff", "SY": "#fff4df",
    }
C = colors()
SOFT = {"GREEN": "SG", "YELLOW": "SY", "RED": "SR"}

css = '''
<style>
.stApp { background: __BG__; color: __TEXT__; }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stToolbar"] { visibility:hidden; }
.block-container { padding-top:1rem; padding-left:1.8rem; padding-right:1.8rem; max-width:100%; }
h1,h2,h3,h4,p,span,label,li { color:__TEXT__; }

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] { background: __SIDEBAR__; border-right:0; min-width:296px; max-width:296px; width:296px !important; }
section[data-testid="stSidebar"] > div { padding-top:1.1rem; }
section[data-testid="stSidebar"] * { color:#eef4ff !important; }
section[data-testid="stSidebar"] .stButton > button {
    background:transparent; color:#d9e5f7 !important; border:1px solid transparent;
    border-radius:10px; height:48px; font-weight:600; margin-bottom:4px;
    justify-content:flex-start; padding-left:16px; box-shadow:none;
}
section[data-testid="stSidebar"] .stButton > button p { text-align:left; font-size:15px; }
section[data-testid="stSidebar"] .stButton > button:hover { background:#14284d; border-color:#1f3a68; }
section[data-testid="stSidebar"] .stButton > button[kind="primary"],
section[data-testid="stSidebar"] button[data-testid="stBaseButton-primary"] {
    background:#1f4fb8 !important; border:1px solid #2a63da !important; color:#fff !important;
}
.brand { display:flex; gap:12px; align-items:center; padding:6px 2px 22px; }
.brand-title { font-size:20px;font-weight:700;line-height:1.2; }
.brand-subtitle { font-size:12px;color:#a9bad3 !important;margin-top:4px; }
.side-foot { color:#a9bad3 !important;font-size:12px;line-height:1.6;padding:14px 4px; }

/* ---------- top bar ---------- */
.user-chip { display:flex; align-items:center; justify-content:flex-end; gap:10px; height:42px; }
.avatar { width:36px;height:36px;border-radius:50%;background:#1e3a8a;color:#fff !important;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700; }
.user-name { font-size:13px;font-weight:500; }

/* ---------- page header ---------- */
.ph { display:flex; gap:16px; align-items:center; margin-bottom:16px; }
.ph-icon { width:58px;height:58px;border-radius:16px;background:__SB__;display:flex;align-items:center;justify-content:center; }
.page-title { font-size:30px;font-weight:800;letter-spacing:-.5px;line-height:1.15; }
.page-subtitle { color:__MUTED__;font-size:14px;margin-top:3px; }

/* ---------- bordered containers = cards ---------- */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background:__CARD__; border:1px solid __BORDER__ !important; border-radius:16px;
    box-shadow:0 8px 24px rgba(16,42,76,.05);
}
.sec-title { font-size:15px;font-weight:700;margin-bottom:8px; }
.act-note { color:__MUTED__;font-size:12.5px;line-height:1.6;margin-top:18px; }
.res-title { font-size:16px;font-weight:700;margin-bottom:10px; }

/* ---------- uploader ---------- */
[data-testid="stFileUploaderDropzone"] { background:__CARD2__; border:1.5px dashed #9bbcf0; border-radius:12px; }
[data-testid="stFileUploaderDropzone"] button { background:__PRIMARY__; color:#fff !important; border:none; border-radius:8px; font-weight:700; }
[data-testid="stFileUploaderFile"] { background:__CARD2__; border:1px solid __BORDER__; border-radius:12px; padding:4px 10px; margin-top:6px; }
[data-testid="stTextArea"] textarea { background:__CARD2__ !important;color:__TEXT__ !important;border:1px solid __BORDER__ !important;border-radius:12px !important; }

/* ---------- buttons ---------- */
.stButton > button { border-radius:10px;min-height:42px;font-weight:700;border:1px solid __BORDER__;background:__CARD__;color:__TEXT__; }
.stButton > button:hover { border-color:__PRIMARY__;color:__PRIMARY__; }
.stButton > button[kind="primary"], button[data-testid="stBaseButton-primary"] {
    background:linear-gradient(135deg,__PRIMARY__,__PRIMARY2__);color:white !important;border:none;min-height:50px;font-size:15px;
}

/* ---------- tabs ---------- */
button[data-baseweb="tab"] { font-weight:600; }
button[data-baseweb="tab"][aria-selected="true"] { color:__PRIMARY__ !important; }
div[data-baseweb="tab-highlight"] { background-color:__PRIMARY__ !important; }

/* ---------- result hero ---------- */
.res-hero { display:flex; gap:26px; align-items:center; padding:20px 24px; border-radius:14px; background:__HERO__; border:1px solid __BORDER__; margin-bottom:14px; }
.donut { width:150px;height:150px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0; }
.donut-in { width:116px;height:116px;border-radius:50%;background:__HERO__;display:flex;flex-direction:column;align-items:center;justify-content:center; }
.donut-num { font-size:34px;font-weight:800;line-height:1; }
.donut-num span { font-size:17px;font-weight:600;color:__MUTED__; }
.donut-lbl { font-size:12px;color:__MUTED__;margin-top:6px; }
.hero-body { flex:1; min-width:0; }
.cat-pill { display:inline-block;padding:7px 16px;border-radius:20px;font-size:15px;font-weight:700; }
.cat-pill.sm { padding:4px 12px;font-size:12px; }
.hero-msg { font-size:13px;color:__TEXT__;margin:12px 0 14px; }
.mini-stats { display:flex; gap:12px; }
.mini { flex:1; background:__CARD__; border:1px solid __BORDER__; border-radius:10px; padding:12px 14px; }
.mini-v { font-size:20px;font-weight:700; }
.mini-l { font-size:11.5px;color:__MUTED__;margin-top:2px; }
.cat-row { display:flex; gap:12px; align-items:center; margin:6px 0 8px; }
.cat-title { font-size:15px;font-weight:700; }
.legend { display:flex; gap:34px; font-size:11.5px; margin-bottom:14px; flex-wrap:wrap; }
.legend i { display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:7px;vertical-align:-1px; }
.legend small { color:__MUTED__; }

/* ---------- inner cards ---------- */
.sub-card { background:__CARD__; border:1px solid __BORDER__; border-radius:12px; padding:14px 16px; margin-bottom:12px; }
.sub-head { display:flex; justify-content:space-between; align-items:center; font-size:14px; font-weight:700; margin-bottom:10px; }
.sub-pill { font-size:10.5px;font-weight:600;padding:4px 10px;border-radius:8px;background:__SB__;color:__PRIMARY__ !important; }
.chips { display:flex; flex-wrap:wrap; gap:7px; }
.chip { display:inline-block;padding:5px 12px;border-radius:18px;font-size:11.5px;font-weight:500; }
.chip-green { background:__SG__;color:__GREEN__ !important; }
.chip-red { background:__SR__;color:__RED__ !important; }
.chip-blue { background:__SB__;color:__PRIMARY__ !important; }
.info-row { display:grid; grid-template-columns:24px 84px 1fr; gap:4px; font-size:12px; padding:5px 0; align-items:center; }
.info-row .k { color:__MUTED__; }
.info-row .v { font-weight:500; overflow-wrap:anywhere; }
.small-note { color:__MUTED__;font-size:11px;margin-top:3px; }

/* ---------- skills breakdown chart ---------- */
.bd-legend { display:flex; justify-content:flex-end; gap:16px; font-size:10.5px; margin:-2px 0 6px; }
.bd-legend i { display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:5px; }
.bd-wrap { position:relative; height:190px; margin:10px 4px 44px 40px; }
.bd-grid { position:absolute; left:0; right:0; border-top:1px solid __BORDER__; }
.bd-grid span { position:absolute; left:-40px; top:-7px; font-size:9.5px; color:__MUTED__; width:32px; text-align:right; }
.bd-bars { position:absolute; top:0; bottom:0; left:0; right:0; display:flex; justify-content:space-around; }
.bd-group { position:relative; height:100%; display:flex; align-items:flex-end; gap:3px; }
.bd-bar { width:20px; border-radius:4px 4px 0 0; position:relative; }
.bd-bar span { position:absolute; top:-14px; left:50%; transform:translateX(-50%); font-size:9px; font-weight:700; color:__TEXT__; }
.bd-name { position:absolute; top:100%; left:50%; transform:translateX(-50%); margin-top:6px; font-size:10px; text-align:center; width:78px; color:__TEXT__; }
.meter { margin-top:12px; }
.meter-h { display:flex; justify-content:space-between; font-size:13px; font-weight:700; margin-bottom:6px; }
.track { height:9px; border-radius:6px; background:__BORDER__; overflow:hidden; }
.fill { height:100%; border-radius:6px; }

/* ---------- detailed analysis ---------- */
.da-row { display:flex; align-items:center; gap:10px; margin:14px 0 10px; }
.da-ic { width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff !important;font-size:14px;font-weight:700; }
.da-t { font-size:14px; font-weight:700; flex:1; }
.da-t small { font-weight:500; color:__MUTED__; }
.da-c { font-size:13px; font-weight:600; }
.insight { padding:11px 12px;border-radius:11px;background:__CARD2__;border:1px solid __BORDER__;font-size:12.5px;margin-bottom:8px; }
.sum-title { display:flex; gap:10px; align-items:center; font-size:15px; font-weight:700; margin-bottom:8px; }
.sum-text { font-size:13px; line-height:1.65; color:__MUTED__; }
.rec-box { margin-top:14px; padding:14px 16px; border-radius:12px; background:__SB__; border:1px solid __BORDER__; }
.rec-title { font-size:14px; font-weight:700; margin-bottom:8px; }
.rec-item { font-size:12.5px; margin:6px 0; display:flex; gap:8px; }
.rec-item b { color:__GREEN__ !important; }

/* ---------- dashboard ---------- */
.hero { padding:24px 28px;border-radius:20px;background:linear-gradient(135deg,rgba(38,117,238,.12),rgba(91,156,255,.04));border:1px solid __BORDER__;margin-bottom:20px; }
.hero-title { font-size:29px;font-weight:850;letter-spacing:-.5px; }
.hero-subtitle { color:__MUTED__;font-size:14px;line-height:1.6;max-width:780px;margin-top:7px; }
.card { background:__CARD__;border:1px solid __BORDER__;border-radius:16px;padding:18px;margin-bottom:16px;box-shadow:0 8px 24px rgba(16,42,76,.05); }
.card-title { font-size:15px;font-weight:800;margin-bottom:4px; }
.card-subtitle { color:__MUTED__;font-size:12px;margin-bottom:6px; }
.stat-card { background:__CARD__;border:1px solid __BORDER__;border-radius:15px;padding:17px;min-height:118px;box-shadow:0 6px 20px rgba(16,42,76,.04); }
.stat-icon { width:36px;height:36px;border-radius:11px;background:__SB__;display:flex;align-items:center;justify-content:center;font-size:18px;margin-bottom:10px; }
.stat-label { color:__MUTED__;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.2px; }
.stat-value { font-size:27px;font-weight:850;margin-top:4px; }
.preview { background:__CARD2__;border:1px solid __BORDER__;border-radius:12px;padding:15px;height:260px;overflow:auto;white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px;line-height:1.6;color:__TEXT__; }
.footer { text-align:center;color:__MUTED__;padding:25px;font-size:11px; }
hr { border-color:__BORDER__; }
@media(max-width:900px){.page-title{font-size:25px}.res-hero{flex-direction:column;align-items:flex-start}}
</style>
'''
for k, v in C.items():
    css = css.replace(f"__{k}__", v)
st.markdown(css, unsafe_allow_html=True)

# ============================================================
# MODEL / EXTRACTION
# ============================================================
@st.cache_resource
def load_model():
    return SentenceTransformer("all-mpnet-base-v2")

def read_file(uploaded_file):
    """Read PDF / DOCX / TXT upload and return plain text (keeps line breaks)."""
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    if name.endswith(".pdf"):
        return extract_text(io.BytesIO(data))
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)
    if name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")
    return ""

def mask_contact(text):
    text = re.sub(r"[\w.+-]+@[\w-]+\.[\w.-]+", "[EMAIL REDACTED]", text)
    text = re.sub(r"(?<!\w)(?:\+?\d[\d\s().-]{7,}\d)(?!\w)", "[PHONE REDACTED]", text)
    return text

def clean_text(text):
    return re.sub(r"\s+", " ", text).strip()

def esc(x):
    return html.escape(str(x))

# ---------------- skills ----------------
SKILLS = ["python", "java", "javascript", "typescript", "c++", "c#", "sql", "mysql", "postgresql",
          "mongodb", "html", "css", "react", "angular", "node.js", "django", "flask", "fastapi",
          "machine learning", "deep learning", "artificial intelligence", "data science",
          "data analysis", "nlp", "natural language processing", "tensorflow", "pytorch", "pandas",
          "numpy", "scikit-learn", "matplotlib", "seaborn", "opencv", "jupyter notebook", "streamlit",
          "aws", "azure", "docker", "kubernetes", "git", "github", "linux", "ci/cd", "rest api",
          "excel", "power bi", "tableau", "communication", "leadership", "problem solving",
          "teamwork", "project management"]
SKILL_RES = {s: re.compile(r"(?<![a-z0-9+#.])" + re.escape(s) + r"(?![a-z0-9+#])") for s in SKILLS}
PRETTY = {"sql": "SQL", "nlp": "NLP", "aws": "AWS", "css": "CSS", "html": "HTML", "mysql": "MySQL",
          "postgresql": "PostgreSQL", "mongodb": "MongoDB", "javascript": "JavaScript",
          "typescript": "TypeScript", "node.js": "Node.js", "fastapi": "FastAPI",
          "tensorflow": "TensorFlow", "pytorch": "PyTorch", "numpy": "NumPy",
          "scikit-learn": "Scikit-learn", "github": "GitHub", "power bi": "Power BI",
          "opencv": "OpenCV", "ci/cd": "CI/CD", "rest api": "REST API"}

def pretty(skill):
    return PRETTY.get(skill, skill.title())

def extract_skills(text):
    low = text.lower()
    return sorted(s for s, p in SKILL_RES.items() if p.search(low))

def skill_count(text, skill):
    return len(SKILL_RES[skill].findall(text.lower()))

def section_score(text):
    low = text.lower()
    sections = ["experience", "education", "skills", "projects", "summary", "certification"]
    return round(sum(1 for s in sections if s in low) / len(sections) * 100, 1)

# ---------------- experience ----------------
def resume_years(text):
    vals = [float(v) for v in re.findall(r"(\d+(?:\.\d+)?)\s*\+?\s*(?:years?|yrs?)", text.lower())]
    return max(vals) if vals else 0.0

def required_years(text):
    low = text.lower()
    vals = [float(a) for a, _ in re.findall(r"(\d+)\s*(?:-|–|to)\s*(\d+)\s*\+?\s*(?:years?|yrs?)", low)]
    vals += [float(v) for v in re.findall(r"(\d+)\s*\+?\s*(?:years?|yrs?)", low)]
    return max(vals) if vals else 0.0

# ---------------- education ----------------
EDU_LEVELS = [
    (4, r"ph\.?d|doctorate"),
    (3, r"m\.?tech|m\.?sc|m\.e\.|mca|mba|master'?s|master of|masters"),
    (2, r"b\.?tech|b\.?sc|b\.e\.|bca|bba|b\.?com|bachelor"),
    (1, r"diploma"),
]
LEVEL_NAMES = {0: "Not found", 1: "Diploma", 2: "Bachelor's", 3: "Master's", 4: "PhD"}

def edu_levels(text):
    low = text.lower()
    found = set()
    for lvl, pat in EDU_LEVELS:
        if re.search(r"(?<![a-z])(?:" + pat + r")(?![a-z])", low):
            found.add(lvl)
    return found

def education_text(text):
    m = re.search(
        r"(?i)(?<![a-z])(b\.?\s?tech|b\.?\s?sc|bca|bba|b\.?com|bachelor[^\n,;|]{0,40}|"
        r"m\.?\s?tech|m\.?\s?sc|mca|mba|master[^\n,;|]{0,40}|ph\.?d)[^\n,;|]{0,45}", text)
    return clean_text(m.group(0))[:60] if m else "Not found"

# ---------------- candidate info ----------------
def guess_name(text):
    for line in text.splitlines()[:8]:
        l = line.strip()
        if 2 <= len(l.split()) <= 4 and re.fullmatch(r"[A-Za-z][A-Za-z .'-]+", l) \
                and not re.search(r"resume|curriculum|vitae|profile|summary|objective", l, re.I):
            return l.title() if l.isupper() else l
    return "Not found"

def guess_phone(text):
    for m in re.finditer(r"(?<!\w)\+?\d[\d\s().-]{8,}\d(?!\w)", text):
        digits = re.sub(r"\D", "", m.group(0))
        if 9 <= len(digits) <= 13:
            return clean_text(m.group(0))
    return "Not found"

def guess_email(text):
    m = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", text)
    return m.group(0) if m else "Not found"

def guess_location(text):
    head = text[:700]
    m = re.search(r"(?i)(?:location|address|city)\s*[:\-]\s*([^\n|]{3,60})", head)
    if m:
        return m.group(1).strip()
    m = re.search(r"\b([A-Z][a-z]+(?: [A-Z][a-z]+)?),\s*([A-Z][a-z]+(?: [A-Z][a-z]+)?|[A-Z]{2})\b", head)
    return m.group(0) if m else "Not found"

# ---------------- main analysis ----------------
def categorize(score):
    if score >= 70:
        return "Strong Match", "GREEN"
    if score >= 40:
        return "Moderate Match", "YELLOW"
    return "Weak Match", "RED"

def analyze_resume(resume_text, job_description):
    model = load_model()
    embeddings = model.encode([clean_text(resume_text), clean_text(job_description)], convert_to_tensor=True)
    semantic = max(0, min(100, util.cos_sim(embeddings[0], embeddings[1]).item() * 100))

    resume_skills = extract_skills(resume_text)
    job_skills = extract_skills(job_description)
    matched = sorted(set(resume_skills) & set(job_skills))
    missing = sorted(set(job_skills) - set(resume_skills))
    extra = sorted(set(resume_skills) - set(job_skills))
    keyword_score = (len(matched) / len(job_skills) * 100) if job_skills else min(100, semantic)

    # experience match
    have_y, need_y = resume_years(resume_text), required_years(job_description)
    if need_y > 0:
        exp_score = min(100.0, have_y / need_y * 100)
    else:
        exp_score = 100.0 if have_y > 0 else semantic

    # education match
    have_lvl = max(edu_levels(resume_text), default=0)
    need_set = edu_levels(job_description)
    need_lvl = min(need_set) if need_set else 0
    if need_lvl > 0:
        edu_score = min(100.0, have_lvl / need_lvl * 100)
    else:
        edu_score = 100.0 if have_lvl > 0 else 50.0

    formatting = section_score(resume_text)

    # Heuristic ATS score: semantic relevance + skills + experience + education + structure
    ats = round(semantic * 0.35 + keyword_score * 0.35 + exp_score * 0.15 + edu_score * 0.10 + formatting * 0.05, 1)
    status, color = categorize(ats)

    # skill "prominence" (mentions) for the breakdown chart
    pool = matched or job_skills
    pool = sorted(pool, key=lambda s: -skill_count(job_description, s))[:5]
    levels = []
    for s in pool:
        you = min(100, 55 + 15 * skill_count(resume_text, s)) if skill_count(resume_text, s) else 0
        req = min(100, 55 + 15 * skill_count(job_description, s))
        levels.append((s, you, req))

    info = {
        "name": guess_name(resume_text), "email": guess_email(resume_text),
        "phone": guess_phone(resume_text), "education": education_text(resume_text),
        "location": guess_location(resume_text),
    }
    return {
        "semantic": semantic, "ats_score": ats, "status": status, "color": color,
        "matched": matched, "missing": missing, "extra": extra, "resume_skills": resume_skills,
        "job_skills": job_skills, "keyword_score": keyword_score, "exp_score": exp_score,
        "edu_score": edu_score, "formatting_score": formatting, "levels": levels, "info": info,
        "have_years": have_y, "need_years": need_y, "have_lvl": have_lvl, "need_lvl": need_lvl,
    }

# ============================================================
# UI HELPERS
# ============================================================
DOC_ICON = ('<svg width="30" height="30" viewBox="0 0 48 48" fill="none" stroke="#2675ee" stroke-width="3" '
            'stroke-linecap="round" stroke-linejoin="round"><path d="M10 6h18l8 8v8"/><path d="M10 6v34h12"/>'
            '<path d="M28 6v8h8"/><path d="M16 22h12M16 29h6"/><circle cx="32" cy="33" r="7"/><path d="M37 38l6 6"/></svg>')
LOGO = DOC_ICON.replace('width="30" height="30"', 'width="46" height="46"').replace("#2675ee", "#3b8cff")

def page_header(title, subtitle):
    st.markdown(f'<div class="ph"><div class="ph-icon">{DOC_ICON}</div><div><div class="page-title">{title}</div>'
                f'<div class="page-subtitle">{subtitle}</div></div></div>', unsafe_allow_html=True)

def nav_button(label, icon, page):
    kind = "primary" if st.session_state.page == page else "secondary"
    try:
        clicked = st.button(label, icon=icon, use_container_width=True, key="nav_" + page, type=kind)
    except TypeError:   # older Streamlit without the icon argument
        clicked = st.button(label, use_container_width=True, key="nav_" + page, type=kind)
    if clicked:
        st.session_state.page = page
        st.rerun()

def render_sidebar():
    with st.sidebar:
        st.markdown(f'<div class="brand">{LOGO}<div><div class="brand-title">AI Resume Matching<br>System</div>'
                    '<div class="brand-subtitle">NLP-Powered &bull; Smarter Hiring</div></div></div>', unsafe_allow_html=True)
        nav_button("Dashboard", ":material/home:", "Dashboard")
        nav_button("Resume Analyzer", ":material/description:", "Analyzer")
        nav_button("Settings", ":material/settings:", "Settings")
        nav_button("Analysis History", ":material/history:", "History")
        nav_button("About", ":material/info:", "About")
        st.markdown('<div style="height:30vh"></div>', unsafe_allow_html=True)
        st.markdown('<div class="side-foot">Built with <span style="color:#ff4d5e !important">&#10084;</span> using Python, NLP,<br>'
                    'Sentence Transformers, Streamlit<br>&amp; Scikit-learn</div>', unsafe_allow_html=True)

def topbar():
    _, c1, c2 = st.columns([8, .5, 2.4])
    with c1:
        if st.button("☀️" if st.session_state.theme == "Light" else "🌙", key="theme_toggle"):
            st.session_state.theme = "Dark" if st.session_state.theme == "Light" else "Light"
            st.rerun()
    with c2:
        initials = "".join(w[0] for w in USER_NAME.split()[:2]).upper()
        st.markdown(f'<div class="user-chip"><div class="avatar">{initials}</div><div class="user-name">{esc(USER_NAME)} &#8964;</div></div>',
                    unsafe_allow_html=True)

def chips(items, kind):
    if not items:
        return '<span class="small-note">None</span>'
    return '<div class="chips">' + "".join(f'<span class="chip chip-{kind}">{esc(pretty(s))}</span>' for s in items) + '</div>'

def join_and(items):
    items = [pretty(i) for i in items]
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " and " + items[-1]

def meter(icon, label, val, color):
    return (f'<div class="meter"><div class="meter-h"><span>{icon} {label}</span><b>{val:.0f}%</b></div>'
            f'<div class="track"><div class="fill" style="width:{min(val,100):.0f}%;background:{color}"></div></div></div>')

# ---------------- result blocks ----------------
def hero_html(r):
    col = C[r["color"]]
    soft = C[SOFT[r["color"]]]
    msg = {"Strong Match": "Excellent! Your resume matches well with the job description.",
           "Moderate Match": "Good start. Your resume partially matches the job description.",
           "Weak Match": "Your resume needs work to match this job description."}[r["status"]]
    deg = r["ats_score"] * 3.6
    return ('<div class="res-hero">'
            f'<div class="donut" style="background:conic-gradient({col} {deg}deg,{C["TRACK"]} 0)"><div class="donut-in">'
            f'<div class="donut-num">{r["ats_score"]:.0f}<span>/100</span></div><div class="donut-lbl">ATS Score</div></div></div>'
            '<div class="hero-body">'
            f'<span class="cat-pill" style="background:{soft};color:{col}">🏆 {r["status"]}</span>'
            f'<div class="hero-msg">{msg}</div>'
            '<div class="mini-stats">'
            f'<div class="mini"><div class="mini-v">{r["keyword_score"]:.0f}%</div><div class="mini-l">Skills Match</div></div>'
            f'<div class="mini"><div class="mini-v">{r["exp_score"]:.0f}%</div><div class="mini-l">Experience Match</div></div>'
            f'<div class="mini"><div class="mini-v">{r["edu_score"]:.0f}%</div><div class="mini-l">Education Match</div></div>'
            '</div></div></div>')

def category_html(r):
    col = C[r["color"]]
    soft = C[SOFT[r["color"]]]
    return ('<div class="cat-row"><span class="cat-title">Match Category</span>'
            f'<span class="cat-pill sm" style="background:{soft};color:{col}">{r["status"]}</span></div>'
            '<div class="legend">'
            f'<span><i style="background:{C["GREEN"]}"></i>Strong Match <small>(70–100%)</small></span>'
            f'<span><i style="background:{C["YELLOW"]}"></i>Moderate Match <small>(40–69%)</small></span>'
            f'<span><i style="background:{C["RED"]}"></i>Weak Match <small>(0–39%)</small></span></div>')

def keyskills_html(r):
    ordered = r["matched"] + [s for s in r["resume_skills"] if s not in r["matched"]]
    return ('<div class="sub-card"><div class="sub-head"><span>🧩 Key Skills Found</span><span class="sub-pill">Top Skills</span></div>'
            + chips(ordered[:12], "green") + '</div>')

def candidate_html(r):
    i = r["info"]
    exp_txt = f'{r["have_years"]:g} Years' if r["have_years"] else "Not specified"
    rows = [("👤", "Name", i["name"]), ("✉️", "Email", i["email"]), ("📱", "Mobile", i["phone"]),
            ("💼", "Experience", exp_txt), ("🎓", "Education", i["education"]), ("📍", "Location", i["location"])]
    body = "".join(f'<div class="info-row"><span>{ic}</span><span class="k">{k}</span><span class="v">{esc(v)}</span></div>' for ic, k, v in rows)
    return f'<div class="sub-card"><div class="sub-head"><span>🪪 Candidate Information</span><span class="sub-pill">Edit</span></div>{body}</div>'

def breakdown_html(r):
    head = '<div class="sub-head"><span>📊 Skills Match Breakdown</span></div>'
    legend = (f'<div class="bd-legend"><span><i style="background:{C["PRIMARY"]}"></i>Your Skills</span>'
              f'<span><i style="background:{C["PURPLE"]}"></i>Required Skills</span></div>')
    if not r["levels"]:
        chart = '<div class="small-note" style="padding:30px 0">No recognised skills were found in the job description.</div>'
    else:
        grid = "".join(f'<div class="bd-grid" style="bottom:{p}%"><span>{p}%</span></div>' for p in range(0, 101, 20))
        groups = ""
        for name, you, req in r["levels"]:
            groups += ('<div class="bd-group">'
                       f'<div class="bd-bar" style="height:{you}%;background:{C["PRIMARY"]}"><span>{you}%</span></div>'
                       f'<div class="bd-bar" style="height:{req}%;background:{C["PURPLE"]}"><span>{req}%</span></div>'
                       f'<div class="bd-name">{esc(pretty(name))}</div></div>')
        chart = f'<div class="bd-wrap">{grid}<div class="bd-bars">{groups}</div></div>'
    note = '<div class="small-note">Bars show how prominently each skill appears (mentions) in the resume vs. the job description.</div>'
    meters = meter("🎯", "Experience Match", r["exp_score"], C["GREEN"]) + meter("🎓", "Education Match", r["edu_score"], C["PRIMARY"])
    return f'<div class="sub-card">{head}{legend}{chart}{note}{meters}</div>'

def summary_texts(r):
    total, m = len(r["job_skills"]), len(r["matched"])
    text = f'Your resume demonstrates a {r["status"].lower()} with the job description. '
    if total:
        text += f"You have {m} of {total} required skills"
    else:
        text += "No recognised skills were found in the job description"
    if r["missing"]:
        text += f", but consider adding {join_and(r['missing'][:3])} to improve your chances further."
    else:
        text += ", with no obvious skill gaps."
    recs = []
    if r["missing"]:
        recs.append(f"Add {join_and(r['missing'][:4])} to your skills section — only if you genuinely have experience with them.")
    recs.append("Highlight your project experience with real-world examples and measurable results.")
    recs.append("Keep your resume concise and focused on relevant skills.")
    if r["formatting_score"] < 60:
        recs.append("Use clear section headings (Summary, Skills, Experience, Education, Projects).")
    return text, recs

def render_results(r):
    main, side = st.columns([2.05, 1], gap="medium")
    with main:
        with st.container(border=True):
            st.markdown('<div class="res-title">📊 Analysis Results</div>', unsafe_allow_html=True)
            st.markdown(hero_html(r), unsafe_allow_html=True)
            st.markdown(category_html(r), unsafe_allow_html=True)
            k1, k2 = st.columns([1, 1.15], gap="small")
            with k1:
                st.markdown(keyskills_html(r), unsafe_allow_html=True)
                st.markdown(candidate_html(r), unsafe_allow_html=True)
            with k2:
                st.markdown(breakdown_html(r), unsafe_allow_html=True)
    with side:
        with st.container(border=True):
            st.markdown('<div class="res-title">📈 Detailed Analysis</div>', unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["Skills Analysis", "Experience Analysis", "Education Analysis"])
            total = len(r["job_skills"])
            with t1:
                st.markdown(
                    f'<div class="da-row"><div class="da-ic" style="background:{C["GREEN"]}">✓</div><div class="da-t">Matched Skills</div>'
                    f'<div class="da-c" style="color:{C["GREEN"]}">{len(r["matched"])}/{total}</div></div>' + chips(r["matched"], "green")
                    + f'<div class="da-row"><div class="da-ic" style="background:{C["RED"]}">✕</div><div class="da-t">Missing Skills</div>'
                    f'<div class="da-c" style="color:{C["RED"]}">{len(r["missing"])}/{total}</div></div>' + chips(r["missing"], "red")
                    + f'<div class="da-row"><div class="da-ic" style="background:{C["PRIMARY"]}">+</div>'
                    f'<div class="da-t">Additional Skills <small>(Found in Resume)</small></div>'
                    f'<div class="da-c">{len(r["extra"])}</div></div>' + chips(r["extra"][:12], "blue"),
                    unsafe_allow_html=True)
            with t2:
                have = f'{r["have_years"]:g} years' if r["have_years"] else "Not specified in resume"
                need = f'{r["need_years"]:g} years' if r["need_years"] else "Not specified in job description"
                st.markdown(f'<div class="insight">💼 Resume experience: <b>{have}</b></div>'
                            f'<div class="insight">📋 Required experience: <b>{need}</b></div>'
                            f'<div class="insight">🎯 Experience match: <b>{r["exp_score"]:.0f}%</b></div>', unsafe_allow_html=True)
            with t3:
                req = LEVEL_NAMES[r["need_lvl"]] if r["need_lvl"] else "Not specified in job description"
                st.markdown(f'<div class="insight">🎓 Resume education: <b>{esc(r["info"]["education"])}</b></div>'
                            f'<div class="insight">📋 Required level: <b>{req}</b></div>'
                            f'<div class="insight">🎯 Education match: <b>{r["edu_score"]:.0f}%</b></div>', unsafe_allow_html=True)
        with st.container(border=True):
            text, recs = summary_texts(r)
            st.markdown(f'<div class="sum-title"><span style="font-size:22px">💡</span>Summary</div><div class="sum-text">{esc(text)}</div>'
                        '<div class="rec-box"><div class="rec-title">🎯 Recommendation</div>'
                        + "".join(f'<div class="rec-item"><b>✓</b><span>{esc(x)}</span></div>' for x in recs) + '</div>',
                        unsafe_allow_html=True)

# ============================================================
# PAGES
# ============================================================
def analyzer():
    page_header("Resume Analyzer", "Upload a resume and job description to analyze the candidate's profile, skills and get ATS score with detailed insights.")
    with st.container(border=True):
        u1, u2, act = st.columns([1, 1, .62], gap="large")
        with u1:
            st.markdown('<div class="sec-title">1. Upload Resume</div>', unsafe_allow_html=True)
            resume_file = st.file_uploader("Resume", type=["pdf", "docx", "txt"], label_visibility="collapsed", key="resume_up")
        with u2:
            st.markdown('<div class="sec-title">2. Upload Job Description</div>', unsafe_allow_html=True)
            jd_file = st.file_uploader("Job Description", type=["pdf", "docx", "txt"], label_visibility="collapsed", key="jd_up")
            with st.expander("Or paste job description text"):
                jd_text = st.text_area("Job Description text", height=120, placeholder="Paste the complete job description here...", label_visibility="collapsed")
        with act:
            st.markdown('<div style="height:34px"></div>', unsafe_allow_html=True)
            if st.button("🔍  Analyze Resume", type="primary", use_container_width=True):
                st.session_state._run_analysis = True
            st.markdown('<div class="act-note">AI will extract key information, match skills and calculate your ATS score.</div>', unsafe_allow_html=True)

    resume_text, job_text = "", ""
    try:
        resume_text = read_file(resume_file)
        job_text = read_file(jd_file) if jd_file else jd_text
    except Exception as e:
        st.error(f"Unable to read file: {e}")

    if resume_text:
        with st.expander("📄 Preview extracted resume (contact details masked)", expanded=False):
            st.markdown(f'<div class="preview">{esc(mask_contact(resume_text)[:8000])}</div>', unsafe_allow_html=True)

    if st.session_state.pop("_run_analysis", False):
        if not resume_file:
            st.warning("Please upload your resume first.")
            return
        if not job_text.strip():
            st.warning("Please upload or paste a job description.")
            return
        if not resume_text.strip():
            st.error("Could not extract readable text from your resume.")
            return
        with st.spinner("🤖 AI is analyzing your resume..."):
            result = analyze_resume(resume_text, job_text)
        result.update({"filename": resume_file.name, "date": datetime.now().strftime("%d %b %Y, %I:%M %p")})
        st.session_state.analysis = result
        st.session_state.history.append({"filename": resume_file.name, "score": result["ats_score"],
                                         "semantic": result["semantic"], "status": result["status"], "date": result["date"]})
        st.success("🎉 Analysis completed successfully!")

    result = st.session_state.analysis
    if not result:
        return
    st.markdown("<br>", unsafe_allow_html=True)
    render_results(result)
    if st.button("🗑️ Clear Current Result"):
        st.session_state.analysis = None
        st.rerun()

def dashboard():
    page_header("Welcome to AI Resume Matching System 👋", "Analyze resumes against job descriptions and get AI-powered match, ATS and skill insights.")
    history = st.session_state.history
    total = len(history)
    avg = sum(x["score"] for x in history) / total if total else 0
    strong = sum(1 for x in history if x["score"] >= 70)
    moderate = sum(1 for x in history if 40 <= x["score"] < 70)
    weak = sum(1 for x in history if x["score"] < 40)
    latest = history[-1] if history else None
    st.markdown('<div class="hero"><div class="hero-title">Find your perfect job match.</div><div class="hero-subtitle">Upload a resume and job description to calculate semantic similarity, ATS compatibility, matched skills and missing requirements.</div></div>', unsafe_allow_html=True)
    cols = st.columns(5)
    stats = [("📄", "Total Resumes", total, "Analyzed"), ("🎯", "Average ATS", f"{avg:.1f}", "Out of 100"),
             ("🏆", "Strong Matches", strong, "70–100"), ("📈", "Moderate Matches", moderate, "40–69"),
             ("⚠️", "Weak Matches", weak, "Below 40")]
    for col, (icon, label, value, note) in zip(cols, stats):
        with col:
            st.markdown(f'<div class="stat-card"><div class="stat-icon">{icon}</div><div class="stat-label">{label}</div><div class="stat-value">{value}</div><div class="small-note">{note}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    a, b = st.columns([1.25, 1])
    with a:
        st.markdown('<div class="card"><div class="card-title">📄 Resume Analyzer</div><div class="card-subtitle">Upload a resume and compare it with a target job description.</div></div>', unsafe_allow_html=True)
        if st.button("🔍 Open Resume Analyzer", type="primary", use_container_width=True):
            st.session_state.page = "Analyzer"
            st.rerun()
    with b:
        if latest:
            badge = {"Strong Match": "GREEN", "Moderate Match": "YELLOW", "Weak Match": "RED"}[latest["status"]]
            st.markdown(f'<div class="card"><div class="card-title">🕘 Latest Analysis</div><div style="font-size:27px;font-weight:850;">{latest["score"]:.0f}/100</div>'
                        f'<span class="cat-pill sm" style="background:{C[SOFT[badge]]};color:{C[badge]}">{latest["status"]}</span></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="card"><div class="card-title">🕘 Latest Analysis</div><div class="small-note">No analysis yet. Start with Resume Analyzer.</div></div>', unsafe_allow_html=True)
    if history:
        left, right = st.columns([1.35, 1])
        with left:
            st.markdown('<div class="card"><div class="card-title">📈 ATS Score Trend</div></div>', unsafe_allow_html=True)
            st.line_chart(pd.DataFrame({"Analysis": range(1, total + 1), "ATS Score": [x["score"] for x in history]}).set_index("Analysis"), height=260)
        with right:
            st.markdown('<div class="card"><div class="card-title">📊 Match Category Distribution</div></div>', unsafe_allow_html=True)
            st.bar_chart(pd.DataFrame({"Category": ["Strong", "Moderate", "Weak"], "Count": [strong, moderate, weak]}).set_index("Category"), height=260)
        rows = [{"Resume": x["filename"], "ATS Score": f'{x["score"]:.0f}/100', "Status": x["status"], "Date": x["date"]} for x in reversed(history[-8:])]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="card" style="text-align:center;padding:45px"><div style="font-size:42px">📊</div><h3>No analytics yet</h3><div class="small-note">Analyze your first resume to populate the dashboard charts.</div></div>', unsafe_allow_html=True)

def history_page():
    page_header("Analysis History", "Review your previous resume matching sessions.")
    h = st.session_state.history
    if not h:
        st.info("No previous analyses found.")
        return
    df = pd.DataFrame(h)
    df = df.rename(columns={"filename": "Resume", "score": "ATS Score", "semantic": "Semantic %", "status": "Status", "date": "Date"})
    st.dataframe(df[["Resume", "ATS Score", "Semantic %", "Status", "Date"]], use_container_width=True, hide_index=True)
    if st.button("🗑️ Clear Analysis History"):
        st.session_state.history = []
        st.rerun()

def settings():
    page_header("Settings", "Configure your Resume Matching experience.")
    st.markdown('<div class="card"><div class="card-title">🎨 Appearance</div><div class="card-subtitle">Use the sun/moon button at the top right to switch between light and dark modes.</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><div class="card-title">🔐 Privacy</div><div class="card-subtitle">Files are processed in your session. Email addresses and phone numbers are masked in the resume preview.</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><div class="card-title">🤖 AI Model</div><div class="card-subtitle">Sentence Transformers • all-mpnet-base-v2</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><div class="card-title">🧮 ATS Score Formula</div><div class="card-subtitle">35% semantic similarity + 35% skills match + 15% experience match + 10% education match + 5% resume structure.</div></div>', unsafe_allow_html=True)
    st.warning("The ATS score is a heuristic compatibility estimate, not a score from a specific commercial ATS platform.")

def about():
    page_header("About", "AI-powered resume matching built with NLP.")
    st.markdown('<div class="card"><div class="card-title">AI Resume Matching System</div><div class="card-subtitle">Compares a resume with a job description using sentence embeddings (semantic similarity), skill extraction, experience and education matching, then produces an ATS-style score with recommendations.</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><div class="card-title">Tech stack</div><div class="card-subtitle">Python • NLP • Sentence Transformers • Streamlit • Scikit-learn compatible workflow</div></div>', unsafe_allow_html=True)

# ============================================================
# ROUTER
# ============================================================
render_sidebar()
topbar()
page = st.session_state.page
if page == "Dashboard":
    dashboard()
elif page == "Analyzer":
    analyzer()
elif page == "History":
    history_page()
elif page == "Settings":
    settings()
elif page == "About":
    about()
st.markdown('<div class="footer">📄 AI Resume Matching System • Python • NLP • Sentence Transformers • Streamlit</div>', unsafe_allow_html=True)
