

"""
=====================================================================
 ADVANCED ATS RESUME AUTOMATOR  v2 — MULTI-PROVIDER FAILOVER
 (Streamlit + OpenAI-compatible APIs + python-docx)
=====================================================================
 WHAT'S NEW IN v2
 - Works with up to 5 AI providers: Gemini, Cerebras, Groq,
   Mistral, OpenRouter. Enter keys for the ones you have.
 - AUTOMATIC FAILOVER: if the first provider is rate-limited or
   fails, the app silently tries the next one. A rate-limited
   provider is skipped for 15 minutes, then retried.
 - Everything else (8-stage tailoring, locked facts, ATS docx,
   Tailored_Resumes folder) is unchanged.

 RUN IT:   streamlit run app.py
=====================================================================
"""

import os
import re
import json
import time
from pathlib import Path
from datetime import datetime

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# ---------------------------------------------------------------
# PROVIDER CHAIN — tried top to bottom. Reorder to change priority.
# All use OpenAI-compatible endpoints, so one client handles all.
# ---------------------------------------------------------------
PROVIDERS = [
    {
        "id": "gemini", "label": "Google Gemini",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "models": ["gemini-3-flash-preview", "gemini-3.1-flash-lite", "gemini-2.5-flash"],
        "hint": "Free key: aistudio.google.com → Get API key",
    },
    {
        "id": "cerebras", "label": "Cerebras",
        "base_url": "https://api.cerebras.ai/v1",
        "models": ["gpt-oss-120b", "zai-glm-4.7"],
        "hint": "Free key: cloud.cerebras.ai → API Keys (~1M tokens/day)",
    },
    {
        "id": "groq", "label": "Groq",
        "base_url": "https://api.groq.com/openai/v1",
        "models": ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"],
        "hint": "Free key: console.groq.com → API Keys",
    },
    {
        "id": "mistral", "label": "Mistral",
        "base_url": "https://api.mistral.ai/v1",
        "models": ["mistral-small-latest"],
        "hint": "Free key: console.mistral.ai (data-training opt-in)",
    },
    {
        "id": "openrouter", "label": "OpenRouter",
        "base_url": "https://openrouter.ai/api/v1",
        "models": ["meta-llama/llama-3.3-70b-instruct:free", "openrouter/free"],
        "hint": "Free key: openrouter.ai → Keys (50 req/day free)",
    },
]

COOLDOWN_SECONDS = 15 * 60   # skip a rate-limited provider for 15 min
RATE_LIMIT_MARKERS = ("429", "rate", "quota", "exceed", "resource_exhausted",
                      "capacity", "overloaded", "limit")

st.set_page_config(page_title="ATS Resume Automator", page_icon="⚡", layout="wide")

# ---------------------------------------------------------------
# 1. TEXT EXTRACTION  (reads the original resume ONE time)
# ---------------------------------------------------------------
def extract_text_from_file(file_path: Path) -> str:
    """Read raw text from a .docx or .pdf resume."""
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
        raise ValueError("Unsupported file type. Please use .docx or .pdf")


# ---------------------------------------------------------------
# 2. MULTI-PROVIDER AI CALL WITH AUTOMATIC FAILOVER
# ---------------------------------------------------------------
def extract_json(raw: str) -> dict:
    """Pull a JSON object out of a model reply, tolerating fences/extra text."""
    raw = re.sub(r"^```(json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in AI reply")
    return json.loads(raw[start:end + 1])


def call_ai_json(keys: dict, prompt: str):
    """
    Try each provider in PROVIDERS order (skipping ones without a key
    or on cooldown). Returns (parsed_json, provider_label, notes).
    """
    cooldowns = st.session_state.setdefault("cooldowns", {})
    now = time.time()
    notes, errors = [], []

    MODEL_GONE_MARKERS = ("404", "not found", "no longer available", "deprecated",
                          "decommissioned", "does not exist", "invalid model")

    for p in PROVIDERS:
        key = (keys.get(p["id"]) or "").strip()
        if not key:
            continue
        wait = cooldowns.get(p["id"], 0) - now
        if wait > 0:
            notes.append(f"{p['label']}: on cooldown ({int(wait // 60) + 1} min left)")
            continue
        client = OpenAI(api_key=key, base_url=p["base_url"], timeout=90)
        rate_limited = False
        for model in p["models"]:
            try:
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                    max_tokens=4096,
                )
                data = extract_json(resp.choices[0].message.content)
                return data, f"{p['label']} ({model})", notes
            except Exception as e:                    # noqa: BLE001
                msg = str(e)
                errors.append(f"{p['label']} / {model}: {msg[:140]}")
                low = msg.lower()
                if any(m in low for m in MODEL_GONE_MARKERS):
                    notes.append(f"{p['label']}: model '{model}' unavailable → trying next model")
                    continue
                if any(m in low for m in RATE_LIMIT_MARKERS):
                    cooldowns[p["id"]] = time.time() + COOLDOWN_SECONDS
                    notes.append(f"{p['label']}: rate-limited → switching to next provider")
                    rate_limited = True
                    break
                notes.append(f"{p['label']} / {model}: error → trying next option")
                continue
        if rate_limited:
            continue

    raise RuntimeError(
        "All configured AI providers failed or are on cooldown.\n\n" +
        "\n".join(errors + notes) +
        "\n\nFixes: wait a few minutes, add another provider key in the sidebar, "
        "or check your internet connection."
    )


# ---------------------------------------------------------------
# 3. STEP ONE — PARSE ORIGINAL RESUME INTO A STRUCTURED PROFILE
#    (cached to _profile.json → only 1 AI call per candidate, ever)
# ---------------------------------------------------------------
PARSE_PROMPT = """You are a resume parsing engine. Convert the resume text below into strict JSON.
Return ONLY JSON, no commentary. Use this exact schema:

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
      "bullets": ["", ""]
    }}
  ],
  "education": ["", ""]
}}

Rules:
- Copy job titles, company names, dates, education, and certifications EXACTLY as written.
- Keep every bullet point of every job.
- "contact_line" = location | phone | email | linkedin joined with " | ".
- "education" = one string per degree/certification line.

RESUME TEXT:
----------------
{resume_text}
----------------
"""


def profile_cache_path(resume_path: Path) -> Path:
    return resume_path.parent / (resume_path.stem + "_profile.json")


def load_or_parse_profile(keys: dict, resume_path: Path, force: bool = False):
    cache = profile_cache_path(resume_path)
    if cache.exists() and not force:
        return json.loads(cache.read_text(encoding="utf-8")), None
    text = extract_text_from_file(resume_path)
    profile, provider, _ = call_ai_json(keys, PARSE_PROMPT.format(resume_text=text))
    cache.write_text(json.dumps(profile, indent=2), encoding="utf-8")
    return profile, provider


# ---------------------------------------------------------------
# 4. STEP TWO — TAILOR THE PROFILE TO A SPECIFIC JD (8-STAGE PROCESS)
# ---------------------------------------------------------------
TAILOR_PROMPT = """You are an elite ATS resume optimization engine used by a professional recruiter.
Tailor the candidate's resume to the Job Description by working through this 8-STAGE PROCESS, in order.

STAGE 1 — DECONSTRUCT THE JD FIRST:
Extract five things: (a) the exact job title, (b) hard skills & tools, (c) process language
(e.g. backlog, risk/issue logs, escalation, status reporting), (d) the domain/industry,
(e) soft requirements (e.g. matrixed environment, self-motivated). Rank keywords by priority:
anything in the job title, in the first paragraph, or repeated 2+ times is TOP PRIORITY and
must appear in the summary, the skills, AND at least one experience bullet.

STAGE 2 — LOCKED vs EDITABLE ZONES:
LOCKED (copy through completely unchanged): name, headline, contact_line, job titles,
company names, dates, education, skill category names and their order, the number of bullets
per job, and every real metric (%, $, team sizes) — numbers are what recruiters trust.
EDITABLE: summary wording, skill items inside existing categories, bullet wording.

STAGE 3 — GAP ANALYSIS (the most important stage). Sort every top JD requirement into 3 buckets:
- DIRECT MATCH: the candidate already does it -> rewrite it using the JD's exact phrase
  (e.g. resume says "risk logs", JD says "risk/issue logs" -> use "risk/issue logs").
- IMPLIED MATCH: the candidate clearly did it but never used the JD's words -> make the
  implied work explicit (e.g. "facilitated stand-ups and retrospectives" implies
  "documented meeting notes, decisions, and follow-up actions").
- NO MATCH: the candidate does not have it -> DO NOT FABRICATE. Never invent employers,
  clients, domain experience (e.g. Medicaid), certifications, degrees, or tools.
  Only emphasize genuinely adjacent real experience, then stop.

STAGE 4 — SUMMARY FORMULA (4-5 sentences):
[Role aligned to JD title] + [years of experience] + [real scale numbers] +
[top 5-6 JD keywords woven in naturally] + [outcome language].
The summary carries the heaviest keyword load.

STAGE 5 — SKILLS REBUILD:
Same categories, same order. Inside each category put JD-priority terms first. Include both
forms of a term where useful — "Microsoft Project (MS Project)", acronym plus spelled-out
"Risk/Issue Logs (RAID)" — because different ATS systems match differently. Only add a skill
if it logically follows from real experience (someone running Jira sprints can honestly
list "Backlog Management").

STAGE 6 — BULLET FORMULA (keep the same bullet count per job):
Each bullet = strong action verb + task in JD terminology + tool + real metric or outcome.
Maximum 1-2 JD keywords per bullet (more = stuffing, and modern ATS penalizes it).
Never open two bullets in the same job with the same verb.
Spread the top keywords across ALL jobs so density looks natural, not dumped in one place.

STAGE 7 — ATS COMPLIANCE CHECK:
Target roughly 70-80% coverage of the top JD keywords across the whole document.
100% coverage is a red flag to recruiters, not a goal. Honestly compute your coverage.

STAGE 8 — HUMAN READ-THROUGH:
Read the result as a recruiter, not a machine. Every line must sound human-written and be
defensible by the candidate in an interview. If a line fails either test, soften it.

OUTPUT: Return ONLY a JSON object — the FULL updated profile in the SAME schema as the
input profile, PLUS these extra top-level keys:
  "job_title_detected": short job title from the JD (used for the file name),
  "company_detected": hiring company from the JD, or "Company" if unknown,
  "keyword_coverage_percent": integer 0-100 from Stage 7,
  "matched_keywords": top 15 JD keywords now present in the resume,
  "missing_keywords": top JD requirements deliberately NOT added (the NO MATCH bucket),
  "gap_analysis": {{
      "direct_matches": ["..."],
      "implied_matches": ["..."],
      "no_match_not_fabricated": ["..."]
  }}

(A) CANDIDATE PROFILE JSON:
{profile_json}

(B) JOB DESCRIPTION:
----------------
{jd_text}
----------------
"""


def tailor_profile(keys: dict, profile: dict, jd_text: str):
    prompt = TAILOR_PROMPT.format(
        profile_json=json.dumps(profile, indent=1),
        jd_text=jd_text[:15000],
    )
    return call_ai_json(keys, prompt)


# ---------------------------------------------------------------
# 5. DOCX GENERATION  (clean, single-column, 100% ATS-parseable)
# ---------------------------------------------------------------
def build_docx(profile: dict, out_path: Path):
    doc = Document()

    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    for side in ("left", "right", "top", "bottom"):
        setattr(section, f"{side}_margin", Inches(0.5))

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(10)
    style.paragraph_format.space_after = Pt(2)

    def para(align=None, before=0, after=2):
        p = doc.add_paragraph()
        if align:
            p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        return p

    def run(p, text, bold=False, italic=False, size=10):
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
        r.font.size = Pt(size)
        r.font.name = "Times New Roman"
        return r

    def section_header(text):
        p = para(before=8, after=2)
        run(p, text.upper(), bold=True, size=10.5)

    p = para(WD_ALIGN_PARAGRAPH.CENTER)
    run(p, profile.get("name", ""), bold=True, size=14)
    if profile.get("headline"):
        p = para(WD_ALIGN_PARAGRAPH.CENTER)
        run(p, profile["headline"], bold=True, size=10.5)
    if profile.get("contact_line"):
        p = para(WD_ALIGN_PARAGRAPH.CENTER, after=4)
        run(p, profile["contact_line"])

    if profile.get("summary"):
        section_header("Professional Summary")
        p = para()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run(p, profile["summary"])

    if profile.get("skills"):
        section_header("Core Skills")
        for s in profile["skills"]:
            p = para()
            run(p, f"{s.get('category', '')}: ", bold=True)
            run(p, s.get("items", ""))

    if profile.get("experience"):
        section_header("Professional Experience")
        usable_width = section.page_width - section.left_margin - section.right_margin
        for job in profile["experience"]:
            p = para(before=6)
            p.paragraph_format.tab_stops.add_tab_stop(usable_width, WD_TAB_ALIGNMENT.RIGHT)
            run(p, f"{job.get('title', '')} — {job.get('company', '')}", bold=True)
            run(p, "\t")
            run(p, job.get("dates", ""), bold=True)
            for b in job.get("bullets", []):
                bp = doc.add_paragraph(style="List Bullet")
                bp.paragraph_format.space_after = Pt(1)
                bp.paragraph_format.left_indent = Inches(0.25)
                r = bp.add_run(b)
                r.font.size = Pt(10)
                r.font.name = "Times New Roman"

    if profile.get("education"):
        section_header("Education & Certifications")
        for line in profile["education"]:
            p = para()
            if "—" in line:
                left, right = line.split("—", 1)
                run(p, left.strip() + " ", bold=True)
                run(p, "— " + right.strip())
            else:
                run(p, line, bold=True)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")[:60]


# ===============================================================
#                      STREAMLIT UI
# ===============================================================
st.title("⚡ ATS Resume Automator")
st.caption("Paste a JD → get a tailored, ATS-clean resume in seconds. "
           "Job titles, companies, dates & education are never changed. No fake facts. "
           "Multiple AI providers with automatic failover.")

with st.sidebar:
    st.header("① One-time Setup")
    st.markdown("**AI provider keys** — fill at least one. "
                "More keys = automatic backup when one hits its limit. "
                "Tried in this order:")
    keys = {}
    for p in PROVIDERS:
        keys[p["id"]] = st.text_input(p["label"], type="password", help=p["hint"],
                                      key=f"key_{p['id']}")
    st.markdown("---")

    resume_input = st.text_input(
        "Original resume path",
        value=r"D:\Noman\BETHLEHEM\Bethlehem_Lulseged_Resume.pdf",
        placeholder=r"C:\noman\Bethlehem_Lulseged_Resume.pdf",
        help="Full path to the candidate's original .docx or .pdf")
    force_reparse = st.checkbox("Re-parse resume (use if you edited the original)", value=False)

    cooldowns = st.session_state.get("cooldowns", {})
    active_cd = {pid: t for pid, t in cooldowns.items() if t > time.time()}
    if active_cd:
        st.markdown("---")
        st.markdown("**⏳ Rate-limited (auto-retry soon):**")
        label_by_id = {p["id"]: p["label"] for p in PROVIDERS}
        for pid, t in active_cd.items():
            mins = int((t - time.time()) // 60) + 1
            st.markdown(f"- {label_by_id.get(pid, pid)} — {mins} min")

st.header("② Paste the Job Description")
jd_text = st.text_area("Job Description", height=280,
                       placeholder="Copy the entire JD from LinkedIn / Indeed / Dice and paste here…")

generate = st.button("🚀 Generate Tailored Resume", type="primary", use_container_width=True)

if generate:
    if not any((v or "").strip() for v in keys.values()):
        st.error("Please enter at least one AI provider key in the sidebar.")
        st.stop()
    if not resume_input or not Path(resume_input.strip().strip('"')).exists():
        st.error("Resume path not found. Paste the FULL path, e.g. C:\\noman\\resume.pdf")
        st.stop()
    if len(jd_text.strip()) < 100:
        st.error("That JD looks too short — paste the full job description.")
        st.stop()

    resume_path = Path(resume_input.strip().strip('"'))
    t0 = time.time()

    try:
        with st.status("Working…", expanded=True) as status:
            st.write("📄 Reading candidate profile…")
            profile, parse_provider = load_or_parse_profile(keys, resume_path,
                                                            force=force_reparse)
            if parse_provider:
                st.write(f"　↳ profile learned via {parse_provider} (one-time)")

            st.write("🧠 Tailoring resume to this JD (8-stage process)…")
            tailored, provider_used, notes = tailor_profile(keys, profile, jd_text)
            for n in notes:
                st.write(f"　↳ {n}")
            st.write(f"　↳ generated via **{provider_used}**")

            st.write("📝 Building ATS-compliant Word document…")
            job_title = tailored.get("job_title_detected", "Role")
            company = tailored.get("company_detected", "Company")
            first_name = safe_filename(profile.get("name", "Candidate").split()[0])
            fname = f"{first_name}_{safe_filename(job_title)}_{safe_filename(company)}.docx"

            out_dir = resume_path.parent / "Tailored_Resumes"
            out_path = out_dir / fname
            if out_path.exists():
                out_path = out_dir / f"{out_path.stem}_{datetime.now():%H%M%S}.docx"

            build_docx(tailored, out_path)
            status.update(label=f"Done in {time.time()-t0:.1f}s ✅ (via {provider_used})",
                          state="complete")

        st.success(f"**Saved:** `{out_path}`")

        col1, col2, col3 = st.columns([2, 1, 2])
        with col1:
            st.subheader("🎯 Detected Role")
            st.write(f"**{job_title}** at **{company}**")
            with open(out_path, "rb") as f:
                st.download_button("⬇️ Download Resume", f, file_name=fname,
                                   use_container_width=True)
        with col2:
            st.subheader("📊 Coverage")
            pct = tailored.get("keyword_coverage_percent")
            st.metric("JD keywords", f"{pct}%" if pct is not None else "—",
                      help="Stage 7 target is 70–80%. 100% looks fake to recruiters.")
        with col3:
            st.subheader("🔑 Keywords Matched")
            kws = tailored.get("matched_keywords", [])
            st.write(" · ".join(kws) if kws else "—")

        gap = tailored.get("gap_analysis", {}) or {}
        with st.expander("🔍 Gap Analysis (Stage 3) — how the JD was matched", expanded=True):
            g1, g2, g3 = st.columns(3)
            with g1:
                st.markdown("**✅ Direct matches**")
                for item in gap.get("direct_matches", []) or ["—"]:
                    st.markdown(f"- {item}")
            with g2:
                st.markdown("**🔁 Implied matches (made explicit)**")
                for item in gap.get("implied_matches", []) or ["—"]:
                    st.markdown(f"- {item}")
            with g3:
                st.markdown("**⛔ Not added (candidate lacks this)**")
                for item in (gap.get("no_match_not_fabricated", [])
                             or tailored.get("missing_keywords", []) or ["—"]):
                    st.markdown(f"- {item}")
            st.caption("⛔ items were deliberately NOT claimed — never state these in the "
                       "application or interview prep either.")

        with st.expander("Preview new summary & skills"):
            st.markdown(f"**Summary:** {tailored.get('summary','')}")
            for s in tailored.get("skills", []):
                st.markdown(f"**{s.get('category','')}:** {s.get('items','')}")

    except Exception as e:                            # noqa: BLE001
        st.error(f"Something went wrong: {e}")
        st.info("Common fixes: check your API keys, add another provider key, "
                "check internet, or tick 'Re-parse resume' if the profile cache "
                "is corrupted.")