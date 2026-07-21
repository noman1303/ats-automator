"""
=====================================================================
 ADVANCED ATS RESUME AUTOMATOR v5 — STREAMLIT CLOUD SAFE & ATS SCORING
 (Streamlit + OpenAI-compatible APIs + python-docx + pypdf)
=====================================================================
 Developed by Noman Belim | Fixed for Streamlit Cloud & Gemini API
=====================================================================
"""

import os
import re
import csv
import json
import time
import tempfile
import io
import zipfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------------------------------------------------------
# APP METADATA
# ---------------------------------------------------------------
APP_AUTHOR = "Noman Belim"
APP_VERSION = "v5.1 — Fixed Cloud Deployment & Diagnostics"

JD_SPLIT_MARKER = "===NEXT JD==="
CUSTOM_PATH_LABEL = "— Upload Resume (.docx or .pdf) —"

PROVIDERS = [
    {
        "id": "gemini", 
        "label": "Google Gemini",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "models": ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"],
        "hint": "Free key: aistudio.google.com → Get API key",
    },
]

GEMINI_BASE_URL = PROVIDERS[0]["base_url"]
GEMINI_MODELS = PROVIDERS[0]["models"]
COOLDOWN_SECONDS = 15 * 60

RATE_LIMIT_MARKERS = ("429", "rate", "quota", "exceed", "resource_exhausted", "capacity", "overloaded", "limit")
MODEL_GONE_MARKERS = ("404", "not found", "no longer available", "deprecated", "decommissioned", "does not exist", "invalid model")

# ---------------------------------------------------------------
# STREAMLIT CONFIG & CUSTOM STYLING
# ---------------------------------------------------------------
st.set_page_config(page_title="ATS Resume Automator", page_icon="⚡", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background: radial-gradient(circle at 15% -10%, rgba(139,92,246,0.16), transparent 45%),
                    radial-gradient(circle at 85% 110%, rgba(236,72,153,0.14), transparent 45%),
                    #0e0e17;
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #14141f 0%, #191927 100%);
        border-right: 1px solid rgba(139,92,246,0.18);
    }
    h1 {
        background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 40%, #ec4899 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800 !important;
    }
    h2 {
        color: #c4b5fd !important;
        font-weight: 700 !important;
        border-left: 4px solid #8b5cf6;
        padding-left: 10px;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        border: 1px solid rgba(139,92,246,0.35) !important;
        transition: all 0.18s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(139,92,246,0.35);
        border-color: #ec4899 !important;
    }
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(139,92,246,0.12), rgba(236,72,153,0.10));
        border: 1px solid rgba(139,92,246,0.3);
        border-radius: 14px;
        padding: 14px 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------
# HELPER & NLP KEYWORD EXTRACTION ENGINE
# ---------------------------------------------------------------
STOPWORDS = {
    "the", "and", "a", "an", "to", "of", "in", "for", "on", "with", "as", "is", "at", "by", "or",
    "be", "this", "that", "will", "are", "you", "your", "our", "we", "from", "have", "has", "it",
    "its", "into", "such", "who", "may", "can", "all", "their", "these", "those", "if", "not",
    "than", "then", "them", "they", "he", "she", "his", "her", "which", "about", "including",
    "etc", "per", "up", "out", "over", "under", "job", "role", "position", "candidate", "candidates",
    "applicant", "applicants", "company", "team", "years", "year", "experience", "ability", "strong",
    "must", "required", "requirements", "preferred", "skills", "work", "working", "responsibilities",
    "description", "duties", "successful", "opportunity", "environment"
}

def parse_keys_text(text: str) -> list[str]:
    keys, seen = [], set()
    for line in text.splitlines():
        key = line.strip()
        if not key or key.startswith("#"):
            continue
        if key not in seen:
            seen.add(key)
            keys.append(key)
    return keys

def load_keys_from_file(file_path: str) -> list[str]:
    p = Path(file_path.strip().strip('"'))
    if not p.exists():
        raise FileNotFoundError(f"Key file not found: {p}")
    return parse_keys_text(p.read_text(encoding="utf-8", errors="ignore"))

def gemini_key_short(key: str) -> str:
    if len(key) <= 10:
        return key[:4] + "…"
    return f"{key[:6]}…{key[-4:]}"

def extract_text_from_file(file_path: Path) -> str:
    suffix = file_path.suffix.lower()
    if suffix == ".docx":
        doc = Document(str(file_path))
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(t for t in parts if t.strip())
    elif suffix == ".pdf":
        from pypdf import PdfReader
        reader = PdfReader(str(file_path))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    else:
        raise ValueError("Unsupported file type. Please upload a .docx or .pdf file.")

def extract_jd_keywords(jd_text: str, top_n: int = 30) -> list[str]:
    text = jd_text.lower()
    words = re.findall(r"[a-z][a-z0-9+#./-]*", text)
    words = [w.strip(".-/") for w in words if w.strip(".-/")]

    freq = {}
    for w in words:
        if len(w) < 3 or w in STOPWORDS:
            continue
        freq[w] = freq.get(w, 0) + 1

    bigrams, trigrams = {}, {}
    for i in range(len(words) - 1):
        w1, w2 = words[i], words[i + 1]
        if w1 not in STOPWORDS and w2 not in STOPWORDS and len(w1) >= 3 and len(w2) >= 3:
            phrase = f"{w1} {w2}"
            bigrams[phrase] = bigrams.get(phrase, 0) + 1

    for i in range(len(words) - 2):
        w1, w2, w3 = words[i], words[i + 1], words[i + 2]
        if w1 not in STOPWORDS and w3 not in STOPWORDS and len(w1) >= 3 and len(w3) >= 3:
            phrase = f"{w1} {w2} {w3}"
            trigrams[phrase] = trigrams.get(phrase, 0) + 1

    ranked_trigrams = sorted(((p, c) for p, c in trigrams.items() if c >= 2), key=lambda x: -x[1])
    ranked_bigrams = sorted(((p, c) for p, c in bigrams.items() if c >= 2), key=lambda x: -x[1])
    ranked_words = sorted(freq.items(), key=lambda x: -x[1])

    keywords, seen = [], set()
    for phrase, _ in ranked_trigrams[:5]:
        seen.add(phrase)
        keywords.append(phrase)

    for phrase, _ in ranked_bigrams[:15]:
        if phrase not in seen:
            seen.add(phrase)
            keywords.append(phrase)

    for word, _ in ranked_words:
        if len(keywords) >= top_n:
            break
        if word not in seen and not any(word in k.split() for k in keywords):
            seen.add(word)
            keywords.append(word)

    return keywords[:top_n]

def flatten_profile_text(profile: dict) -> str:
    parts = [
        profile.get("summary", ""),
        profile.get("headline", ""),
        profile.get("name", "")
    ]
    for s in profile.get("skills", []) or []:
        parts.append(s.get("category", ""))
        parts.append(s.get("items", ""))
    for job in profile.get("experience", []) or []:
        parts.append(job.get("title", ""))
        parts.append(job.get("company", ""))
        for b in job.get("bullets", []) or []:
            parts.append(b)
    for e in profile.get("education", []) or []:
        parts.append(e if isinstance(e, str) else str(e))
    return " ".join(parts).lower()

def calculate_ats_score(jd_text: str, profile: dict, top_n: int = 30):
    keywords = extract_jd_keywords(jd_text, top_n=top_n)
    resume_blob = flatten_profile_text(profile)

    matched, missing = [], []
    for kw in keywords:
        if kw in resume_blob:
            matched.append(kw)
        else:
            missing.append(kw)

    score = round(100 * len(matched) / len(keywords)) if keywords else 0
    return score, matched, missing

# ---------------------------------------------------------------
# AI EXECUTION ENGINE
# ---------------------------------------------------------------
def extract_json(raw: str) -> dict:
    raw = re.sub(r"^```(json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in AI reply")
    return json.loads(raw[start:end + 1])

def _try_gemini_key(key: str, prompt: str, cooldowns: dict, notes: list, errors: list):
    now = time.time()
    short = gemini_key_short(key)
    wait = cooldowns.get(f"gemini:{key}", 0) - now
    if wait > 0:
        notes.append(f"Gemini key {short}: on cooldown ({int(wait // 60) + 1} min left)")
        return None

    client = OpenAI(api_key=key, base_url=GEMINI_BASE_URL, timeout=90)
    for model in GEMINI_MODELS:
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
                max_tokens=4096,
            )
            data = extract_json(resp.choices[0].message.content)
            return data, f"Gemini key {short} ({model})"
        except Exception as e:
            msg = str(e)
            errors.append(f"Gemini {short} / {model}: {msg[:140]}")
            low = msg.lower()
            if any(m in low for m in MODEL_GONE_MARKERS):
                continue
            if any(m in low for m in RATE_LIMIT_MARKERS):
                cooldowns[f"gemini:{key}"] = time.time() + COOLDOWN_SECONDS
                return None
            continue
    return None

def call_ai_json(keys: dict, gemini_key_pool: list, prompt: str):
    cooldowns = st.session_state.setdefault("cooldowns", {})
    notes, errors = [], []

    all_keys = list(gemini_key_pool)
    manual_gemini_key = (keys.get("gemini") or "").strip()
    if manual_gemini_key and manual_gemini_key not in all_keys:
        all_keys.append(manual_gemini_key)

    if not all_keys:
        raise ValueError("No Gemini API key provided. Please enter a key or upload a key file in the sidebar.")

    for key in all_keys:
        res = _try_gemini_key(key, prompt, cooldowns, notes, errors)
        if res:
            return res[0], res[1], notes

    raise RuntimeError("All configured AI keys failed or hit rate limits:\n" + "\n".join(errors + notes))

# ---------------------------------------------------------------
# PROMPTS
# ---------------------------------------------------------------
PARSE_PROMPT = """You are a resume parsing engine. Convert the raw text below into strict JSON.
Return ONLY JSON.

Schema:
{{
  "name": "",
  "headline": "",
  "contact_line": "",
  "summary": "",
  "skills": [ {{"category": "", "items": ""}} ],
  "experience": [
    {{
      "title": "",
      "company": "",
      "dates": "",
      "bullets": [""]
    }}
  ],
  "education": [""]
}}

RESUME TEXT:
{resume_text}
"""

TAILOR_PROMPT = """You are an expert Resume Writer & ATS Optimization Specialist.
Your goal is to increase the candidate's ATS Match Score to 85%+ while strictly preserving factuality.

MUST-HAVE KEYWORDS TO INTEGRATE:
{priority_keywords}

RULES FOR ATS OPTIMIZATION:
1. TARGET JOB TITLE: Match the headline directly to the target job title from the JD.
2. SUMMARY REWRITE: High-impact 4-5 sentence summary containing top priority keywords naturally.
3. SKILLS SECTION: Group key skills under existing categories. Include full skill names and acronyms.
4. EXPERIENCE BULLETS: Rewrite bullets using action verbs + exact JD terminology + tools + measurable impact.
   - Do NOT fabricate companies, dates, degrees, or tools never used.
   - Re-align candidate's real tasks to use the JD's phrasing.
   - Keep the exact same number of bullets per role as the original resume.

{cover_letter_stage}

OUTPUT FORMAT (STRICT JSON ONLY):
{{
  "name": "",
  "headline": "",
  "contact_line": "",
  "summary": "",
  "skills": [ {{"category": "", "items": ""}} ],
  "experience": [
    {{
      "title": "",
      "company": "",
      "dates": "",
      "bullets": [""]
    }}
  ],
  "education": [""],
  "job_title_detected": "Exact Job Title from JD",
  "company_detected": "Company Name from JD or Company"
  {cover_letter_key}
}}

(A) CANDIDATE PROFILE:
{profile_json}

(B) JOB DESCRIPTION:
{jd_text}
"""

COVER_LETTER_STAGE = """
COVER LETTER GENERATION:
Write a high-converting, professional cover letter tailored specifically to this role and company using real profile facts.
Return it in the "cover_letter" field as a single string with double newlines between paragraphs.
"""

def load_or_parse_profile(keys: dict, gemini_key_pool: list, resume_path: Path, force: bool = False):
    cache = resume_path.parent / (resume_path.stem + "_profile.json")
    if cache.exists() and not force:
        return json.loads(cache.read_text(encoding="utf-8")), None
    text = extract_text_from_file(resume_path)
    profile, provider, _ = call_ai_json(keys, gemini_key_pool, PARSE_PROMPT.format(resume_text=text))
    cache.write_text(json.dumps(profile, indent=2), encoding="utf-8")
    return profile, provider

def tailor_profile(keys: dict, gemini_key_pool: list, profile: dict, jd_text: str, include_cover_letter: bool = False):
    keywords = extract_jd_keywords(jd_text, top_n=25)
    kw_str = ", ".join(keywords)

    cover_letter_key = ',\n  "cover_letter": "..."' if include_cover_letter else ""
    prompt = TAILOR_PROMPT.format(
        priority_keywords=kw_str,
        cover_letter_stage=COVER_LETTER_STAGE if include_cover_letter else "",
        cover_letter_key=cover_letter_key,
        profile_json=json.dumps(profile, indent=1),
        jd_text=jd_text[:15000],
    )
    return call_ai_json(keys, gemini_key_pool, prompt)

# ---------------------------------------------------------------
# DOCX GENERATION ENGINE
# ---------------------------------------------------------------
def build_docx(profile: dict, out_path: Path):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    # Name
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run(profile.get("name", ""))
    r_name.font.name = "Calibri"
    r_name.font.size = Pt(20)
    r_name.bold = True

    # Headline
    if profile.get("headline"):
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_head = p_head.add_run(profile.get("headline", ""))
        r_head.font.name = "Calibri"
        r_head.font.size = Pt(11)
        r_head.bold = True

    # Contact Line
    if profile.get("contact_line"):
        p_cnt = doc.add_paragraph()
        p_cnt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cnt = p_cnt.add_run(profile.get("contact_line", ""))
        r_cnt.font.name = "Calibri"
        r_cnt.font.size = Pt(9.5)

    def add_heading(title: str):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title.upper())
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.bold = True

    # Summary
    if profile.get("summary"):
        add_heading("Professional Summary")
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(profile["summary"])
        r.font.name = "Calibri"
        r.font.size = Pt(10)

    # Skills
    if profile.get("skills"):
        add_heading("Core Competencies & Technical Skills")
        for sk in profile["skills"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            r_cat = p.add_run(f"•  {sk.get('category', '')}: ")
            r_cat.bold = True
            r_cat.font.name = "Calibri"
            r_cat.font.size = Pt(10)

            r_itm = p.add_run(sk.get("items", ""))
            r_itm.font.name = "Calibri"
            r_itm.font.size = Pt(10)

    # Experience
    if profile.get("experience"):
        add_heading("Professional Experience")
        for job in profile["experience"]:
            p_job = doc.add_paragraph()
            p_job.paragraph_format.space_before = Pt(4)
            p_job.paragraph_format.space_after = Pt(1)

            r_t = p_job.add_run(f"{job.get('title', '')} ")
            r_t.bold = True
            r_t.font.name = "Calibri"
            r_t.font.size = Pt(10.5)

            r_c = p_job.add_run(f"| {job.get('company', '')} ")
            r_c.font.name = "Calibri"
            r_c.font.size = Pt(10)

            if job.get("dates"):
                r_d = p_job.add_run(f"({job.get('dates', '')})")
                r_d.font.name = "Calibri"
                r_d.font.size = Pt(9.5)

            for b in job.get("bullets", []) or []:
                p_b = doc.add_paragraph(style='List Bullet')
                p_b.paragraph_format.space_after = Pt(2)
                r_b = p_b.add_run(b)
                r_b.font.name = "Calibri"
                r_b.font.size = Pt(10)

    # Education
    if profile.get("education"):
        add_heading("Education & Certifications")
        for ed in profile["education"]:
            p_ed = doc.add_paragraph()
            p_ed.paragraph_format.space_after = Pt(2)
            r_e = p_ed.add_run(f"•  {ed}")
            r_e.font.name = "Calibri"
            r_e.font.size = Pt(10)

    doc.save(str(out_path))

def build_cover_letter_docx(name: str, contact: str, title: str, company: str, text: str, out_path: Path):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    p_n = doc.add_paragraph()
    r_n = p_n.add_run(name)
    r_n.font.name = "Calibri"
    r_n.font.size = Pt(16)
    r_n.bold = True

    if contact:
        p_c = doc.add_paragraph()
        r_c = p_c.add_run(contact)
        r_c.font.name = "Calibri"
        r_c.font.size = Pt(9.5)

    doc.add_paragraph()

    for para in text.split("\n\n"):
        if para.strip():
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run(para.strip())
            r.font.name = "Calibri"
            r.font.size = Pt(11)

    doc.save(str(out_path))

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip().replace(" ", "_")

# ---------------------------------------------------------------
# STREAMLIT UI & WORKFLOW
# ---------------------------------------------------------------
st.title("⚡ ATS Resume Automator")
st.caption("Maximize ATS Match Score across Single & Batch Job Applications")

st.sidebar.header("🔑 API Keys & Configuration")

manual_gemini_key = st.sidebar.text_input("Enter Gemini API Key", type="password")
uploaded_key_file = st.sidebar.file_uploader("Or Upload Gemini Keys TXT File", type=["txt"])

gemini_pool = []
if uploaded_key_file:
    gemini_pool = parse_keys_text(uploaded_key_file.getvalue().decode("utf-8", errors="ignore"))

if gemini_pool:
    st.sidebar.success(f"Loaded {len(gemini_pool)} Gemini Keys from file")

keys = {"gemini": manual_gemini_key}

st.sidebar.divider()
st.sidebar.header("📄 Candidate Resume")

uploaded_resume = st.sidebar.file_uploader("Upload Candidate Resume (.docx or .pdf)", type=["docx", "pdf"])
resume_file_path = None

if uploaded_resume:
    temp_dir = Path(tempfile.gettempdir()) / "ats_resumes"
    temp_dir.mkdir(exist_ok=True)
    resume_file_path = temp_dir / uploaded_resume.name
    resume_file_path.write_bytes(uploaded_resume.getvalue())

generate_cover_letter = st.sidebar.checkbox("Generate Cover Letter alongside Resume", value=True)

tab_single, tab_batch = st.tabs(["🎯 Single JD Optimization", "📦 Batch Processing Mode"])

# ---------------------------------------------------------------
# TAB 1: SINGLE MODE
# ---------------------------------------------------------------
with tab_single:
    st.subheader("Single Job Description Optimization")
    jd_input = st.text_area("Paste Job Description (JD) here:", height=250)

    if st.button("🚀 Optimize Resume for ATS", type="primary", use_container_width=True):
        if not resume_file_path or not resume_file_path.exists():
            st.error("Please upload a candidate resume (.docx or .pdf) in the sidebar first.")
        elif not jd_input.strip():
            st.error("Please paste a Job Description.")
        elif not gemini_pool and not manual_gemini_key:
            st.error("Please enter a Gemini API Key or upload a key file in the sidebar.")
        else:
            try:
                with st.status("Optimizing Resume...", expanded=True) as status:
                    st.write(" Parsing original candidate profile...")
                    profile, _ = load_or_parse_profile(keys, gemini_pool, resume_file_path)

                    orig_score, orig_matched, orig_missing = calculate_ats_score(jd_input, profile)
                    st.write(f" Initial Candidate ATS Score: **{orig_score}%**")

                    st.write(" Tailoring content against JD requirements...")
                    tailored, provider_used, _ = tailor_profile(keys, gemini_pool, profile, jd_input, generate_cover_letter)

                    new_score, new_matched, new_missing = calculate_ats_score(jd_input, tailored)
                    st.write(f" Optimized Resume ATS Score: **{new_score}%** (Powered by {provider_used})")

                    status.update(label=" Optimization Complete!", state="complete")

                # Metrics Display
                c1, c2, c3 = st.columns(3)
                c1.metric("Original ATS Match", f"{orig_score}%")
                c2.metric("Optimized ATS Match", f"{new_score}%", delta=f"+{new_score - orig_score}%")
                c3.metric("Matched Keywords", f"{len(new_matched)} / {len(new_matched) + len(new_missing)}")

                st.divider()

                # Output File Generation
                out_dir = Path(tempfile.gettempdir()) / "ats_output" / sanitize_filename(tailored.get("company_detected", "Company"))
                out_dir.mkdir(parents=True, exist_ok=True)

                res_filename = f"{sanitize_filename(profile['name'])}_{sanitize_filename(tailored.get('job_title_detected', 'Tailored'))}_Resume.docx"
                res_path = out_dir / res_filename
                build_docx(tailored, res_path)

                col_res, col_cov = st.columns(2)
                with col_res:
                    st.success("Resume Generated Successfully!")
                    st.download_button(
                        label="📥 Download ATS Resume (.docx)",
                        data=res_path.read_bytes(),
                        file_name=res_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

                if generate_cover_letter and "cover_letter" in tailored:
                    cov_filename = f"{sanitize_filename(profile['name'])}_{sanitize_filename(tailored.get('job_title_detected', 'Role'))}_CoverLetter.docx"
                    cov_path = out_dir / cov_filename
                    build_cover_letter_docx(
                        profile["name"], profile.get("contact_line", ""),
                        tailored.get("job_title_detected", ""), tailored.get("company_detected", ""),
                        tailored["cover_letter"], cov_path
                    )
                    with col_cov:
                        st.success("Cover Letter Generated Successfully!")
                        st.download_button(
                            label="📥 Download Cover Letter (.docx)",
                            data=cov_path.read_bytes(),
                            file_name=cov_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

                # Keyword Breakdown
                st.divider()
                st.subheader("Keyword Match Breakdown")
                col_m, col_u = st.columns(2)
                with col_m:
                    st.markdown("** Matched Keywords in Resume:**")
                    st.write(", ".join([f"`{k}`" for k in new_matched]) if new_matched else "None")
                with col_u:
                    st.markdown("** Unmatched / Omitted Keywords:**")
                    st.write(", ".join([f"`{k}`" for k in new_missing]) if new_missing else "None (100% Coverage)")

            except Exception as ex:
                st.error(f"Execution Error: {str(ex)}")

# ---------------------------------------------------------------
# TAB 2: BATCH PROCESSING MODE
# ---------------------------------------------------------------
with tab_batch:
    st.subheader("Batch Application Processing")
    st.caption("Process multiple JDs separated by `===NEXT JD===` or upload multiple .txt files.")

    batch_text = st.text_area("Paste Multiple JDs (separated by `===NEXT JD===`):", height=200)
    batch_files = st.file_uploader("Or Upload JD Files (.txt)", type=["txt"], accept_multiple_files=True)

    if st.button("⚡ Run Batch ATS Optimization", type="primary", use_container_width=True):
        jds_to_process = []

        if batch_text.strip():
            for block in batch_text.split(JD_SPLIT_MARKER):
                if block.strip():
                    jds_to_process.append(block.strip())

        if batch_files:
            for f in batch_files:
                content = f.getvalue().decode("utf-8", errors="ignore").strip()
                if content:
                    jds_to_process.append(content)

        if not resume_file_path or not resume_file_path.exists():
            st.error("Please upload a candidate resume (.docx or .pdf) in the sidebar first.")
        elif not jds_to_process:
            st.error("No valid JDs provided for batch processing.")
        elif not gemini_pool and not manual_gemini_key:
            st.error("Please enter a Gemini API key in the sidebar.")
        else:
            try:
                zip_buffer = io.BytesIO()
                summary_rows = [["Index", "Company", "Job Title", "Original ATS Score", "Optimized ATS Score", "Status"]]

                profile, _ = load_or_parse_profile(keys, gemini_pool, resume_file_path)
                progress_bar = st.progress(0)

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, jd in enumerate(jds_to_process, 1):
                        st.write(f"Processing JD {idx}/{len(jds_to_process)}...")
                        try:
                            orig_score, _, _ = calculate_ats_score(jd, profile)
                            tailored, _, _ = tailor_profile(keys, gemini_pool, profile, jd, generate_cover_letter)
                            new_score, _, _ = calculate_ats_score(jd, tailored)

                            comp = sanitize_filename(tailored.get("company_detected", f"Company_{idx}"))
                            role = sanitize_filename(tailored.get("job_title_detected", "Tailored_Role"))

                            with tempfile.TemporaryDirectory() as tmpdir:
                                res_p = Path(tmpdir) / f"{comp}_{role}_Resume.docx"
                                build_docx(tailored, res_p)
                                zip_file.write(res_p, arcname=f"{comp}/{res_p.name}")

                                if generate_cover_letter and "cover_letter" in tailored:
                                    cov_p = Path(tmpdir) / f"{comp}_{role}_CoverLetter.docx"
                                    build_cover_letter_docx(
                                        profile["name"], profile.get("contact_line", ""),
                                        tailored.get("job_title_detected", ""), comp,
                                        tailored["cover_letter"], cov_p
                                    )
                                    zip_file.write(cov_p, arcname=f"{comp}/{cov_p.name}")

                            summary_rows.append([idx, comp, role, f"{orig_score}%", f"{new_score}%", "Success"])
                        except Exception as err:
                            summary_rows.append([idx, "Unknown", "Unknown", "N/A", "N/A", f"Failed: {str(err)}"])

                        progress_bar.progress(idx / len(jds_to_process))

                    csv_buffer = io.StringIO()
                    writer = csv.writer(csv_buffer)
                    writer.writerows(summary_rows)
                    zip_file.writestr("batch_summary.csv", csv_buffer.getvalue())

                st.success("Batch Processing Complete!")
                st.download_button(
                    label="📦 Download Batch Results ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"ATS_Batch_Applications_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            except Exception as ex:
                st.error(f"Batch Execution Error: {str(ex)}")
