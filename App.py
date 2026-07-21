"""
=====================================================================
 ADVANCED ATS RESUME AUTOMATOR  v4 — BATCH MODE + COVER LETTERS
 (Streamlit + OpenAI-compatible APIs + python-docx)
=====================================================================
 Developed by Noman Belim
=====================================================================
 WHAT'S NEW IN v4
 - BATCH MODE: process many job descriptions in a single run instead
   of one at a time. Paste several JDs separated by a marker line, or
   upload one .txt file per JD. The app tailors the resume for every
   JD, then hands back ONE zip containing one company/ subfolder per
   application (built for people applying to a lot of jobs per day).
   A failure on one JD (bad key, malformed JD, etc.) is caught and
   logged per-row instead of stopping the whole batch, and a
   summary.csv is included in the zip so you can see coverage % and
   status for every application at a glance.
 - AUTO COVER LETTERS: an optional checkbox generates a tailored,
   non-fabricated cover letter alongside every resume, in the SAME
   AI call as the resume tailoring (no extra API call, so it doesn't
   cost you extra quota at volume). Works in both single and batch
   mode. Cover letters use only real facts from the candidate's
   profile — same "never fabricate" rule as the resume.
 - Everything from v3 (multi-key Gemini pool, multi-provider
   failover, 8-stage tailoring, locked facts, code-level keyword
   coverage) is unchanged.

 WHAT'S NEW IN v3
 - MULTI-KEY GEMINI POOL: paste ALL your Gemini API keys (one per
   line) into a text file (e.g. gemini_keys.txt). The app loads them
   automatically and rotates through them — when one key gets
   rate-limited, it's put on cooldown and the app instantly tries
   the next key in your file. No more copy-pasting keys by hand.
 - Still supports the other 4 providers (Cerebras, Groq, Mistral,
   OpenRouter) as single-key fallbacks after the Gemini pool is
   exhausted.
 - Everything else (8-stage tailoring, locked facts, ATS docx,
   output/ folder, code-level keyword coverage) is
   unchanged from v2.

 HOW TO USE THE KEY FILE:
   1. Create a plain .txt file, e.g. C:\noman\gemini_keys.txt
   2. Paste one Gemini API key per line (blank lines / lines
      starting with # are ignored):

        AIzaSyABC123...
        AIzaSyDEF456...
        # this one is my backup account
        AIzaSyGHI789...

   3. In the sidebar, point "Gemini keys file" at that path.
   4. Whenever a key gets rate-limited, the app skips it for 15
      minutes and moves to the next key automatically — you never
      touch the app again until all keys in the file are cooling
      down.
   5. Add more keys any time by editing the .txt file and clicking
      "🔄 Reload keys from file" in the sidebar (no restart needed).

 =====================================================================
 HOW TO RUN THIS APP (Windows) — copy-paste these, in order
 =====================================================================
 1) Open Command Prompt.

 2) Go to the folder where this file lives:
        cd D:\Noman

 3) (One-time / whenever packages need installing or updating)
    If plain "pip" isn't recognized, use "python -m pip" instead:
        python -m pip install streamlit openai python-docx pypdf

 4) Run the app.
    If plain "streamlit" isn't recognized (common when Python isn't
    fully added to PATH), run it as a module instead — this always
    works as long as "python" itself works:
        python -m streamlit run app.py

 5) It opens in your browser automatically at something like:
        http://localhost:8501

 6) To stop the app: click back into the Command Prompt window and
    press Ctrl + C.

 TROUBLESHOOTING QUICK REFERENCE:
   - "'pip' is not recognized"       -> use: python -m pip install ...
   - "'streamlit' is not recognized" -> use: python -m streamlit run app.py
   - Check Python is installed at all -> python --version
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
import base64
from pathlib import Path
from datetime import datetime

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# ---------------------------------------------------------------
# APP METADATA / CREDIT
# ---------------------------------------------------------------
APP_AUTHOR = "Noman Belim"
APP_VERSION = "v4"

# ---------------------------------------------------------------
# BATCH MODE — line used to separate multiple pasted JDs in one
# text box. Any JD block may optionally start with a line like
# "COMPANY: Acme Corp" to force the output folder/company name,
# for JDs (e.g. from a recruiter) that don't clearly state the
# hiring company in the JD text itself.
# ---------------------------------------------------------------
JD_SPLIT_MARKER = "===NEXT JD==="
FORCE_COMPANY_PREFIX = "company:"

# ---------------------------------------------------------------
# CANDIDATE PATH PICKER — add your own saved candidates here.
# Dropdown shows these first; "Custom / one-off path" falls back
# to a free-text field for anything not in this list.
# ---------------------------------------------------------------
SAVED_CANDIDATES = {
    "Bethlehem Lulseged": r"D:\Noman\BETHLEHEM\Bethlehem_Lulseged_Resume.pdf",
    # "Another Candidate": r"D:\Noman\OTHER\Another_Candidate_Resume.docx",
}
CUSTOM_PATH_LABEL = "— Custom / one-off path —"

# ---------------------------------------------------------------
# PROVIDER CHAIN — tried top to bottom. Reorder to change priority.
# Gemini is special-cased below to support a POOL of keys from a
# text file; the other providers still use a single key each from
# the sidebar, exactly like v2.
# ---------------------------------------------------------------
PROVIDERS = [
    {
        "id": "gemini", "label": "Google Gemini",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "models": ["gemini-3-flash-preview", "gemini-3.1-flash-lite", "gemini-2.5-flash"],
        "hint": "Free key: aistudio.google.com → Get API key",
    },
    
]

GEMINI_BASE_URL = PROVIDERS[0]["base_url"]
GEMINI_MODELS = PROVIDERS[0]["models"]

COOLDOWN_SECONDS = 15 * 60   # skip a rate-limited provider/key for 15 min
RATE_LIMIT_MARKERS = ("429", "rate", "quota", "exceed", "resource_exhausted",
                      "capacity", "overloaded", "limit")
MODEL_GONE_MARKERS = ("404", "not found", "no longer available", "deprecated",
                      "decommissioned", "does not exist", "invalid model")

st.set_page_config(page_title="ATS Resume Automator", page_icon="⚡", layout="wide")

# ---------------------------------------------------------------
# GLOBAL THEME — "Aurora Violet"
# Dark canvas + violet→pink gradient accents (matches the credit
# badge palette). Paired with .streamlit/config.toml for native
# widget colors (buttons, sliders, checkboxes, focus rings).
# ---------------------------------------------------------------
st.markdown(
    """
    <style>
    /* ---------- page canvas ---------- */
    .stApp {
        background:
            radial-gradient(circle at 15% -10%, rgba(139,92,246,0.16), transparent 45%),
            radial-gradient(circle at 85% 110%, rgba(236,72,153,0.14), transparent 45%),
            #0e0e17;
    }

    /* ---------- sidebar ---------- */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #14141f 0%, #191927 100%);
        border-right: 1px solid rgba(139,92,246,0.18);
    }

    /* ---------- headings ---------- */
    h1 {
        background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 40%, #ec4899 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 800 !important;
    }
    h2 {
        color: #c4b5fd !important;
        font-weight: 700 !important;
        border-left: 4px solid #8b5cf6;
        padding-left: 10px;
    }
    h3 {
        color: #e9d5ff !important;
        font-weight: 650 !important;
    }

    /* ---------- buttons ---------- */
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        border: 1px solid rgba(139,92,246,0.35) !important;
        transition: all 0.18s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(139,92,246,0.35);
        border-color: #ec4899 !important;
    }
    button[kind="primary"] {
        background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 45%, #ec4899 100%) !important;
        box-shadow: 0 3px 14px rgba(139,92,246,0.4) !important;
    }

    /* ---------- inputs ---------- */
    .stTextInput input, .stTextArea textarea {
        border-radius: 8px !important;
        border: 1px solid rgba(139,92,246,0.3) !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #ec4899 !important;
        box-shadow: 0 0 0 2px rgba(236,72,153,0.22) !important;
    }

    /* ---------- expanders ---------- */
    [data-testid="stExpander"] {
        border: 1px solid rgba(139,92,246,0.22) !important;
        border-radius: 12px !important;
        background: rgba(139,92,246,0.05) !important;
        overflow: hidden;
    }

    /* ---------- metrics ---------- */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(139,92,246,0.12), rgba(236,72,153,0.10));
        border: 1px solid rgba(139,92,246,0.3);
        border-radius: 14px;
        padding: 14px 16px;
    }

    /* ---------- alerts (success / error / warning / info) ---------- */
    [data-testid="stAlert"] {
        border-radius: 10px !important;
    }

    /* ---------- dividers ---------- */
    hr {
        border: none !important;
        height: 1px !important;
        background: linear-gradient(90deg, transparent, rgba(139,92,246,0.5), transparent) !important;
        margin: 1.3rem 0 !important;
    }

    /* ---------- status widget ---------- */
    [data-testid="stStatusWidget"] {
        border-radius: 12px !important;
        border: 1px solid rgba(139,92,246,0.25) !important;
    }

    /* ---------- scrollbar ---------- */
    ::-webkit-scrollbar { width: 10px; height: 10px; }
    ::-webkit-scrollbar-track { background: #14141f; }
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #8b5cf6, #ec4899);
        border-radius: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------
# 0. GEMINI KEY LOADER
#    Reads one API key per line from either an uploaded .txt file
#    (works everywhere, including Streamlit Cloud) or a local file
#    path (only works when the app runs on your own computer).
#    Blank lines and lines starting with # are ignored so you can
#    leave yourself notes next to each key.
# ---------------------------------------------------------------
def parse_keys_text(text: str) -> list[str]:
    """Return a de-duplicated, ordered list of keys from raw text (one per line)."""
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
    """
    Return a de-duplicated, ordered list of keys from a LOCAL text file.
    NOTE: this only works when Streamlit is running on the same machine
    where the file lives. On Streamlit Cloud (or any hosted deployment)
    the app runs on a remote server that has no access to your PC's
    C:\\ drive or to a URL you paste in — use the file uploader instead.
    """
    p = Path(file_path.strip().strip('"'))
    if not p.exists():
        raise FileNotFoundError(f"Key file not found: {p}")
    return parse_keys_text(p.read_text(encoding="utf-8", errors="ignore"))


def gemini_key_short(key: str) -> str:
    """Short label for logs/UI so full keys are never printed on screen."""
    if len(key) <= 10:
        return key[:4] + "…"
    return f"{key[:6]}…{key[-4:]}"


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
# 2. MULTI-PROVIDER + MULTI-KEY AI CALL WITH AUTOMATIC FAILOVER
#    Order of attempts:
#      1) Every Gemini key from the key-file pool (round-robin,
#         skipping keys on cooldown), each tried against every
#         Gemini model in turn.
#      2) A single sidebar-entered Gemini key, if the user also
#         typed one directly (kept for backward compatibility).
#      3) The remaining single-key providers (Cerebras, Groq,
#         Mistral, OpenRouter), same as v2.
# ---------------------------------------------------------------
def extract_json(raw: str) -> dict:
    """Pull a JSON object out of a model reply, tolerating fences/extra text."""
    raw = re.sub(r"^```(json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in AI reply")
    return json.loads(raw[start:end + 1])


def _try_gemini_key(key: str, prompt: str, cooldowns: dict, notes: list, errors: list):
    """
    Try one Gemini key across all Gemini models.
    Returns (data, provider_label) on success, or None on failure
    (cooldowns/notes/errors are mutated in place).
    """
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
                temperature=0.3,
                max_tokens=4096,
            )
            data = extract_json(resp.choices[0].message.content)
            return data, f"Gemini key {short} ({model})"
        except Exception as e:                        # noqa: BLE001
            msg = str(e)
            errors.append(f"Gemini {short} / {model}: {msg[:140]}")
            low = msg.lower()
            if any(m in low for m in MODEL_GONE_MARKERS):
                notes.append(f"Gemini key {short}: model '{model}' unavailable → trying next model")
                continue
            if any(m in low for m in RATE_LIMIT_MARKERS):
                cooldowns[f"gemini:{key}"] = time.time() + COOLDOWN_SECONDS
                notes.append(f"Gemini key {short}: rate-limited → switching to next key")
                return None
            notes.append(f"Gemini key {short} / {model}: error → trying next option")
            continue
    return None


def call_ai_json(keys: dict, gemini_key_pool: list, prompt: str):
    """
    Try, in order:
      1) every key in gemini_key_pool (skipping ones on cooldown)
      2) a manually-entered Gemini key from the sidebar (if any)
      3) the other single-key providers
    Returns (parsed_json, provider_label, notes).
    """
    cooldowns = st.session_state.setdefault("cooldowns", {})
    notes, errors = [], []

    # --- 1) Rotate through the Gemini key-file pool -------------
    for key in gemini_key_pool:
        result = _try_gemini_key(key, prompt, cooldowns, notes, errors)
        if result is not None:
            data, label = result
            return data, label, notes

    # --- 2) Fall back to a manually-typed Gemini key, if present -
    manual_gemini_key = (keys.get("gemini") or "").strip()
    if manual_gemini_key and manual_gemini_key not in gemini_key_pool:
        result = _try_gemini_key(manual_gemini_key, prompt, cooldowns, notes, errors)
        if result is not None:
            data, label = result
            return data, label, notes

    # --- 3) Fall back to the other single-key providers ----------
    now = time.time()
    for p in PROVIDERS:
        if p["id"] == "gemini":
            continue  # already handled above via the key pool / manual key
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
        "All configured AI providers/keys failed or are on cooldown.\n\n" +
        "\n".join(errors + notes) +
        "\n\nFixes: wait a few minutes, add more Gemini keys to your key file, "
        "add another provider key in the sidebar, or check your internet connection."
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


def load_or_parse_profile(keys: dict, gemini_key_pool: list, resume_path: Path, force: bool = False):
    cache = profile_cache_path(resume_path)
    if cache.exists() and not force:
        return json.loads(cache.read_text(encoding="utf-8")), None
    text = extract_text_from_file(resume_path)
    profile, provider, _ = call_ai_json(keys, gemini_key_pool, PARSE_PROMPT.format(resume_text=text))
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
{cover_letter_stage}
OUTPUT: Return ONLY a JSON object — the FULL updated profile in the SAME schema as the
input profile, PLUS these extra top-level keys:
  "job_title_detected": short job title from the JD (used for the file name),
  "company_detected": hiring company from the JD, or "Company" if unknown,
  "matched_keywords": top 15 JD keywords now present in the resume,
  "missing_keywords": top JD requirements deliberately NOT added (the NO MATCH bucket),
  "gap_analysis": {{
      "direct_matches": ["..."],
      "implied_matches": ["..."],
      "no_match_not_fabricated": ["..."]
  }}{cover_letter_key}

Note: keyword coverage percentage is calculated separately in code after you respond —
you do not need to compute or report it yourself.

(A) CANDIDATE PROFILE JSON:
{profile_json}

(B) JOB DESCRIPTION:
----------------
{jd_text}
----------------
"""

# Only injected into the prompt when the user checks "generate cover letter" —
# keeps the prompt (and token cost) smaller for people who don't want one.
COVER_LETTER_STAGE = """
STAGE 9 — COVER LETTER (requested for this run):
Write one complete, ready-to-send cover letter for THIS specific role at THIS specific
company, using ONLY real facts from the candidate profile above (same locked-facts rule
as the resume — never invent an employer, metric, or achievement).
- 3-4 short paragraphs, roughly 250-350 words total, in plain prose (no bullet points).
- Paragraph 1: name the role and company, one line on who the candidate is (title + years).
- Paragraphs 2-3: 2-3 of the strongest DIRECT MATCH / IMPLIED MATCH points from Stage 3,
  written as short stories with real metrics — not a restatement of the resume bullets.
- Final paragraph: brief, confident close + availability + thanks.
- Salutation "Dear Hiring Manager," unless the JD clearly names a person.
- Return it as a single string in the "cover_letter" key, with "\\n\\n" between paragraphs.
"""


def tailor_profile(keys: dict, gemini_key_pool: list, profile: dict, jd_text: str,
                   include_cover_letter: bool = False):
    cover_letter_key = (
        '\n  "cover_letter": full tailored cover letter text per the cover-letter stage '
        'below, "\\n\\n"-separated paragraphs.'
        if include_cover_letter else ""
    )
    prompt = TAILOR_PROMPT.format(
        cover_letter_stage=COVER_LETTER_STAGE if include_cover_letter else "",
        cover_letter_key=cover_letter_key,
        profile_json=json.dumps(profile, indent=1),
        jd_text=jd_text[:15000],
    )
    return call_ai_json(keys, gemini_key_pool, prompt)


# ---------------------------------------------------------------
# 4b. CODE-LEVEL KEYWORD COVERAGE  (replaces AI self-reported %)
#    Extracts the highest-frequency meaningful words/phrases from
#    the JD, then checks how many actually appear in the tailored
#    resume text. This is deterministic and can't be "faked" by
#    the model the way a self-reported percentage could be.
# ---------------------------------------------------------------
STOPWORDS = {
    "the", "and", "a", "an", "to", "of", "in", "for", "on", "with", "as", "is",
    "at", "by", "or", "be", "this", "that", "will", "are", "you", "your", "our",
    "we", "from", "have", "has", "it", "its", "into", "such", "who", "may",
    "can", "all", "their", "these", "those", "if", "not", "than", "then",
    "them", "they", "he", "she", "his", "her", "which", "about", "including",
    "etc", "per", "up", "out", "over", "under", "job", "role", "position",
    "candidate", "candidates", "applicant", "applicants", "company", "team",
    "years", "year", "experience", "ability", "strong", "must", "required",
    "requirements", "preferred", "skills", "work", "working", "including",
}


def extract_jd_keywords(jd_text: str, top_n: int = 25) -> list[str]:
    """
    Pull the top_n highest-frequency, meaningful words/short-phrases
    from the JD text. Combines single words and 2-word phrases so
    multi-word terms (e.g. "risk log", "stand up") aren't lost.
    """
    text = jd_text.lower()
    words = re.findall(r"[a-z][a-z0-9+#./-]*", text)
    words = [w.strip(".-/") for w in words if w.strip(".-/")]

    freq = {}
    for w in words:
        if len(w) < 3 or w in STOPWORDS:
            continue
        freq[w] = freq.get(w, 0) + 1

    bigrams = {}
    for i in range(len(words) - 1):
        w1, w2 = words[i], words[i + 1]
        if w1 in STOPWORDS or w2 in STOPWORDS or len(w1) < 3 or len(w2) < 3:
            continue
        phrase = f"{w1} {w2}"
        bigrams[phrase] = bigrams.get(phrase, 0) + 1

    ranked_words = sorted(freq.items(), key=lambda x: (-x[1], x[0]))
    ranked_bigrams = sorted(
        ((p, c) for p, c in bigrams.items() if c >= 2),
        key=lambda x: (-x[1], x[0]),
    )

    keywords, seen = [], set()
    for phrase, _ in ranked_bigrams[: top_n // 2]:
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
    """Collapse the tailored profile into one lowercase text blob for matching."""
    parts = [
        profile.get("summary", ""),
        profile.get("headline", ""),
    ]
    for s in profile.get("skills", []) or []:
        parts.append(s.get("category", ""))
        parts.append(s.get("items", ""))
    for job in profile.get("experience", []) or []:
        parts.append(job.get("title", ""))
        for b in job.get("bullets", []) or []:
            parts.append(b)
    for e in profile.get("education", []) or []:
        parts.append(e)
    return " ".join(parts).lower()


def calculate_keyword_coverage(jd_text: str, tailored_profile: dict, top_n: int = 25):
    """
    Returns (coverage_percent, matched_list, missing_list) computed
    entirely in code from the JD and the FINAL tailored resume text —
    no dependence on the AI's self-reported number.
    """
    keywords = extract_jd_keywords(jd_text, top_n=top_n)
    resume_blob = flatten_profile_text(tailored_profile)

    matched, missing = [], []
    for kw in keywords:
        if kw in resume_blob:
            matched.append(kw)
        else:
            missing.append(kw)

    pct = round(100 * len(matched) / len(keywords)) if keywords else 0
    return pct, matched, missing


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
        return r

    # 1. Name and Contact Info
    p_name = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    run(p_name, profile.get("name", "Candidate Name").upper(), bold=True, size=22)

    p_head = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
    run(p_head, profile.get("headline", ""), bold=True, size=12)

    p_contact = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=12)
    run(p_contact, profile.get("contact_line", ""), size=10)

    # Section Header Helper
    def add_section_header(title):
        p = para(before=12, after=4)
        run(p, title.upper(), bold=True, size=11)
        # Add a simple bottom border using underscores
        p_border = doc.add_paragraph()
        p_border.paragraph_format.space_before = Pt(0)
        p_border.paragraph_format.space_after = Pt(6)
        p_border.add_run("_" * 80).font.size = Pt(8)

    # 2. Professional Summary
    summary = profile.get("summary", "")
    if summary:
        add_section_header("Professional Summary")
        p_sum = para()
        run(p_sum, summary)

    # 3. Technical Skills
    skills = profile.get("skills", [])
    if skills:
        add_section_header("Technical Skills & Core Competencies")
        for skill_grp in skills:
            p_skill = para()
            run(p_skill, skill_grp.get("category", "") + ": ", bold=True)
            run(p_skill, skill_grp.get("items", ""))

    # 4. Professional Experience
    experience = profile.get("experience", [])
    if experience:
        add_section_header("Professional Experience")
        for job in experience:
            p_job = para(before=6, after=2)
            # Use a right-aligned tab stop for the dates to sit flush right
            tab_stops = p_job.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(7.5), WD_TAB_ALIGNMENT.RIGHT)

            run(p_job, job.get("title", ""), bold=True, size=11)
            p_job.add_run("\t")
            run(p_job, job.get("dates", ""), bold=True, size=10)

            p_comp = para(after=4)
            run(p_comp, job.get("company", ""), italic=True, bold=True)

            for bullet in job.get("bullets", []):
                p_bull = para(after=3)
                p_bull.paragraph_format.left_indent = Inches(0.25)
                p_bull.paragraph_format.first_line_indent = Inches(-0.15)
                run(p_bull, "•  " + bullet)

    # 5. Education
    education = profile.get("education", [])
    if education:
        add_section_header("Education & Certifications")
        for edu in education:
            p_edu = para(after=3)
            run(p_edu, edu)

    doc.save(out_path)


# ---------------------------------------------------------------
# 6. COVER LETTER GENERATION (NEW IN v4)
# ---------------------------------------------------------------
def build_cover_letter_docx(profile: dict, out_path: Path):
    doc = Document()
    
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    for side in ("left", "right", "top", "bottom"):
        setattr(section, f"{side}_margin", Inches(1.0))

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    def para(after=12):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        return p

    def run(p, text, bold=False, size=11):
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
        return r

    # Contact Header (matching resume style)
    p_head = para(after=24)
    run(p_head, profile.get("name", "Candidate Name") + "\n", bold=True, size=16)
    run(p_head, profile.get("contact_line", ""), size=10)

    # Date
    p_date = para(after=24)
    run(p_date, datetime.today().strftime("%B %d, %Y"))

    # Body
    cl_text = profile.get("cover_letter", "")
    if not cl_text:
        cl_text = (
            "Dear Hiring Manager,\n\n"
            "I am writing to express my interest in the open position. "
            "Please find my attached resume outlining my background and qualifications.\n\n"
            "Thank you for your time and consideration.\n\n"
            f"Sincerely,\n{profile.get('name', 'Candidate Name')}"
        )

    for paragraph in cl_text.split("\n\n"):
        if paragraph.strip():
            p_body = para()
            run(p_body, paragraph.strip())

    doc.save(out_path)


# ---------------------------------------------------------------
# 7. MAIN STREAMLIT APP LOGIC (UI & BATCH PROCESSING)
# ---------------------------------------------------------------
def sanitize_filename(name: str) -> str:
    """Strip invalid characters so it's safe to use as a folder/file name."""
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()

def main():
    st.title(f"⚡ Advanced ATS Resume Automator {APP_VERSION}")
    st.markdown(f"**Developed by {APP_AUTHOR}** — Batch Mode + Auto Cover Letters")

    st.sidebar.header("⚙️ Settings & API Keys")

    # 1. Gemini Pool
    st.sidebar.subheader("1. Gemini Multi-Key Pool")
    gemini_key_file = st.sidebar.text_input("Gemini keys file (.txt path)", value=r"C:\keys\gemini_keys.txt")
    if st.sidebar.button("🔄 Reload keys from file"):
        st.rerun()

    gemini_pool = []
    if gemini_key_file.strip():
        try:
            gemini_pool = load_keys_from_file(gemini_key_file)
            if gemini_pool:
                st.sidebar.success(f"Active pool: {len(gemini_pool)} keys loaded.")
            else:
                st.sidebar.warning("File found, but no valid keys inside.")
        except Exception as e:
            st.sidebar.warning(f"Key file missing or invalid path: {e}")

    # Fallback Providers
    with st.sidebar.expander("Single-Key Fallbacks (Optional)"):
        manual_keys = {}
        manual_keys["gemini"] = st.text_input("Manual Gemini Key", type="password")
        for p in PROVIDERS:
            if p["id"] == "gemini":
                continue
            manual_keys[p["id"]] = st.text_input(f"{p['label']} Key", type="password")

    st.sidebar.divider()

    # 2. Candidate Selection
    st.sidebar.subheader("2. Candidate Resume")
    candidate_opts = list(SAVED_CANDIDATES.keys()) + [CUSTOM_PATH_LABEL]
    selected_cand = st.sidebar.selectbox("Select Profile", candidate_opts)

    if selected_cand == CUSTOM_PATH_LABEL:
        resume_path_str = st.sidebar.text_input("Full path to original resume (.pdf or .docx)")
    else:
        resume_path_str = SAVED_CANDIDATES[selected_cand]

    st.sidebar.divider()

    # 3. Add-ons
    st.sidebar.subheader("3. Tailoring Options")
    include_cover_letter = st.sidebar.checkbox("Generate Cover Letter alongside Resume", value=True)

    if not resume_path_str:
        st.info("👈 Please select or enter a candidate resume path in the sidebar to begin.")
        st.stop()

    resume_path = Path(resume_path_str.strip('"').strip("'"))
    if not resume_path.exists():
        st.error(f"Original resume not found at: {resume_path}")
        st.stop()

    # Application Tabs
    tab1, tab2 = st.tabs(["📄 Single Application", "📑 Batch Mode"])

    def process_jds(jd_list):
        with st.spinner("Step 1: Parsing original resume into structured profile..."):
            try:
                profile, init_provider = load_or_parse_profile(manual_keys, gemini_pool, resume_path)
            except Exception as e:
                st.error(f"Failed to parse original resume: {e}")
                return

        # Prepare isolated output directory for this run
        out_dir = Path(tempfile.gettempdir()) / f"ats_automator_run_{int(time.time())}"
        out_dir.mkdir(parents=True, exist_ok=True)

        progress = st.progress(0)
        status = st.empty()
        results = []

        for i, jd_raw in enumerate(jd_list):
            status.markdown(f"**Processing Job {i+1} of {len(jd_list)}...**")
            jd_text = jd_raw.strip()

            # Handle explicit company overrides (helpful for recruiter copy-pastes)
            forced_company = None
            if jd_text.lower().startswith(FORCE_COMPANY_PREFIX):
                first_line, rest = jd_text.split("\n", 1)
                forced_company = first_line.split(":", 1)[1].strip()
                jd_text = rest.strip()

            if not jd_text:
                continue

            try:
                # Core Engine Call
                tailored_profile, provider, notes = tailor_profile(
                    manual_keys, gemini_pool, profile, jd_text, include_cover_letter
                )

                if forced_company:
                    tailored_profile["company_detected"] = forced_company

                comp_name = sanitize_filename(tailored_profile.get("company_detected", f"Company_{i+1}"))
                job_title = sanitize_filename(tailored_profile.get("job_title_detected", f"Role_{i+1}"))

                # Strict code-level keyword matching
                cov_pct, matched, missing = calculate_keyword_coverage(jd_text, tailored_profile)

                # Generate outputs
                app_dir = out_dir / comp_name
                app_dir.mkdir(exist_ok=True)

                base_filename = f"{profile.get('name', 'Candidate').replace(' ', '_')}_{comp_name}_{job_title}"
                
                resume_out = app_dir / f"{base_filename}_Resume.docx"
                build_docx(tailored_profile, resume_out)

                if include_cover_letter and "cover_letter" in tailored_profile:
                    cl_out = app_dir / f"{base_filename}_CoverLetter.docx"
                    build_cover_letter_docx(tailored_profile, cl_out)

                results.append({
                    "Company": comp_name,
                    "Role": job_title,
                    "Coverage (%)": cov_pct,
                    "Provider Used": provider,
                    "Status": "✅ Success"
                })

            except Exception as e:
                results.append({
                    "Company": f"Job {i+1}",
                    "Role": "Unknown",
                    "Coverage (%)": 0,
                    "Provider Used": "-",
                    "Status": f"❌ Error: {str(e)}"
                })

            progress.progress((i + 1) / len(jd_list))

        status.success(f"Generation complete! {len(jd_list)} job(s) evaluated.")

        # Summary CSV for the user
        csv_path = out_dir / "summary.csv"
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["Company", "Role", "Coverage (%)", "Provider Used", "Status"])
            writer.writeheader()
            writer.writerows(results)

        # Zip compilation
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(out_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(out_dir)
                    zf.write(file_path, arcname)

        st.dataframe(results, use_container_width=True)
        st.download_button(
            label="📥 Download Application Materials (ZIP)",
            data=buf.getvalue(),
            file_name=f"Automated_Applications_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            type="primary"
        )

    with tab1:
        st.markdown("Tailor a resume for a single job description.")
        single_jd = st.text_area("Paste Job Description", height=400, key="single_jd")
        if st.button("✨ Tailor Single Application", type="primary"):
            if single_jd.strip():
                process_jds([single_jd])
            else:
                st.warning("Please paste a Job Description first.")

    with tab2:
        st.markdown(f"Tailor for multiple roles at once. Paste all JDs below, separated by exactly: `{JD_SPLIT_MARKER}`")
        st.markdown(f"*Tip: If the JD doesn't list the company, start its block with `{FORCE_COMPANY_PREFIX} Company Name`*")
        batch_jds = st.text_area("Paste Batch Job Descriptions", height=400, key="batch_jds")
        if st.button("✨ Process Batch Applications", type="primary"):
            if batch_jds.strip():
                jd_list = [j.strip() for j in batch_jds.split(JD_SPLIT_MARKER) if j.strip()]
                process_jds(jd_list)
            else:
                st.warning("Please paste your Job Descriptions first.")


if __name__ == "__main__":
    main()
