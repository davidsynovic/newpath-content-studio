import hmac
import io
import os
import re
from datetime import datetime

import streamlit as st

try:
    from groq import Groq
except ImportError:
    Groq = None

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# ------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------
APP_NAME = "NewPath Content Studio"
# Used only if Groq's live model list can't be fetched. Order = preference.
PREFERRED_MODELS = [
    "openai/gpt-oss-120b",
    "qwen/qwen3.6-27b",
    "openai/gpt-oss-20b",
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant",
]
NON_CHAT_MARKERS = ("whisper", "guard", "tts", "orpheus", "safeguard", "compound", "embed")
MAX_SOURCE_CHARS = 24000          # keeps requests inside free-tier token limits
MAX_AUDIO_BYTES = 25 * 1024 * 1024  # Groq free-tier file limit
DEPTH_TOKENS = {"Concise": 5000, "Standard": 7500, "Deep-Dive": 10000}
SECTIONS = [
    "EDITORIAL_FEEDBACK",
    "LINKEDIN_POST",
    "EMAIL_NEWSLETTER",
    "EMAIL_SEQUENCE",
    "EXECUTIVE_BRIEFING",
    "CAROUSEL_SLIDES",
]
LINKEDIN_LIMIT = 3000

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------------
# THEME
# ------------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #090D16 !important; color: #F8FAFC !important;
    }
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
        background-color: #0F172A !important; border-right: 1px solid #1E293B !important;
    }
    h1, h2, h3, h4 { color: #FFFFFF !important; letter-spacing: -0.02em; }
    p, span, label, div, li { color: #E2E8F0 !important; }
    textarea, input, [data-baseweb="select"] > div {
        background-color: #1E293B !important; color: #F8FAFC !important;
        border: 1px solid #334155 !important; border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background-color: #0F172A !important; border: 1px dashed #334155 !important; border-radius: 8px !important;
    }
    [data-testid="stChatMessage"] {
        background-color: #111827 !important; border: 1px solid #1E293B !important;
        border-radius: 10px !important; padding: 12px !important;
    }
    button[data-baseweb="tab"] { background-color: transparent !important; color: #94A3B8 !important; font-weight: 600 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { color: #10B981 !important; border-bottom-color: #10B981 !important; }
    .stButton > button, [data-testid="stDownloadButton"] > button {
        background-color: #059669 !important; color: #FFFFFF !important; border-radius: 8px !important;
        border: 1px solid #10B981 !important; font-weight: 600 !important; padding: 0.55rem 1.1rem !important;
    }
    .stButton > button p, [data-testid="stDownloadButton"] > button p { color: #FFFFFF !important; }
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover { background-color: #047857 !important; }
    hr { border-color: #1E293B !important; }
    .np-badge { display:inline-block; padding:2px 10px; border-radius:999px; background:#064E3B;
                border:1px solid #10B981; color:#A7F3D0 !important; font-size:0.78rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------
# SESSION STATE
# ------------------------------------------------------------------
defaults = {
    "assets": None,
    "chat_history": [],
    "source_material": "",
    "last_file_id": None,
    "last_model_used": "",
    "li_text": "",
    "email_text": "",
    "brief_text": "",
    "slides_text": "",
    "sequence_text": "",
    "check_report": "",
    "uploader_nonce": 0,
}
for k, v in defaults.items():
    st.session_state.setdefault(k, v)


def reset_all():
    """Wipe every setting, draft and upload. Keeps only the API key so you don't re-paste it."""
    keep = {"api_key", "uploader_nonce", "authed"}
    nonce = st.session_state.get("uploader_nonce", 0) + 1
    for k in list(st.session_state.keys()):
        if k not in keep:
            del st.session_state[k]
    st.session_state["uploader_nonce"] = nonce  # new key = empty file uploader


# ------------------------------------------------------------------
# API KEY + MODEL DISCOVERY
# ------------------------------------------------------------------
def get_secret(name: str) -> str:
    """Read a value from Streamlit secrets or the environment."""
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.environ.get(name, "")


def secret_key() -> str:
    """Groq key from secrets/environment, so you never have to paste it."""
    return get_secret("GROQ_API_KEY")


def password_ok() -> bool:
    """Simple access gate. Active only when APP_PASSWORD is set (always set it when deployed)."""
    required = get_secret("APP_PASSWORD")
    if not required or st.session_state.get("authed"):
        return True
    st.title(APP_NAME)
    st.caption("Private preview for NewPath Business Advisors.")
    with st.form("login"):
        entered = st.text_input("Access password", type="password")
        go = st.form_submit_button("Unlock")
    if go:
        if hmac.compare_digest(entered.encode(), required.encode()):
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


@st.cache_data(ttl=600, show_spinner=False)
def fetch_live_models(api_key: str):
    """Ask Groq which models are active right now, so a retired model can never break the app."""
    try:
        client = Groq(api_key=api_key)
        ids = [m.id for m in client.models.list().data]
        chat = [i for i in ids if not any(x in i.lower() for x in NON_CHAT_MARKERS)]
        ordered = [m for m in PREFERRED_MODELS if m in chat]
        ordered += sorted(m for m in chat if m not in ordered)
        return ordered or PREFERRED_MODELS, True
    except Exception:
        return PREFERRED_MODELS, False


def friendly_error(err: Exception) -> str:
    msg = str(err)
    low = msg.lower()
    if "invalid_api_key" in low or "401" in low:
        return "Groq rejected the API key. Create a fresh one at console.groq.com/keys and paste it again."
    if "429" in low or "rate_limit" in low:
        return "Free-tier rate limit hit. Wait about a minute, or pick a smaller model."
    if "decommissioned" in low or "model_not_found" in low or "does not exist" in low:
        return "That model has been retired by Groq. Choose another one in the sidebar."
    if "413" in low or "too large" in low or "reduce your message" in low:
        return "The source text is too long for the free tier. Shorten it or use the Concise depth."
    return f"Groq error: {msg}"


# ------------------------------------------------------------------
# FILE EXTRACTION
# ------------------------------------------------------------------
def extract_file(uploaded_file, api_key: str) -> str:
    if uploaded_file is None:
        return ""
    ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
    try:
        if ext in ("txt", "md", "srt", "csv"):
            return uploaded_file.getvalue().decode("utf-8", errors="ignore")
        if ext == "pdf":
            if not pypdf:
                st.error("PDF support needs `pypdf` (pip install pypdf).")
                return ""
            reader = pypdf.PdfReader(io.BytesIO(uploaded_file.getvalue()))
            return "\n".join(p.extract_text() or "" for p in reader.pages)
        if ext == "docx":
            if not docx:
                st.error("Word support needs `python-docx` (pip install python-docx).")
                return ""
            d = docx.Document(io.BytesIO(uploaded_file.getvalue()))
            return "\n".join(p.text for p in d.paragraphs)
        if ext in ("mp3", "wav", "mp4", "m4a"):
            if not api_key or Groq is None:
                st.error("Add your Groq API key first to transcribe audio or video.")
                return ""
            if uploaded_file.size > MAX_AUDIO_BYTES:
                st.error("Audio/video files over 25 MB can't be transcribed on the free tier. Trim or compress it.")
                return ""
            with st.spinner("Transcribing with Groq Whisper..."):
                client = Groq(api_key=api_key)
                result = client.audio.transcriptions.create(
                    file=(uploaded_file.name, uploaded_file.getvalue()),
                    model="whisper-large-v3-turbo",
                )
                return result.text
    except Exception as e:
        st.error(f"Could not read that file: {friendly_error(e)}")
    return ""


# ------------------------------------------------------------------
# GENERATION
# ------------------------------------------------------------------
def build_prompt(source, detail, audience, tones, rules, cta, feedback, previous):
    tone_str = ", ".join(tones) if tones else "Direct & Authoritative"
    refine_block = ""
    if feedback and previous:
        refine_block = f"""
REFINEMENT MODE. The person reviewed the previous drafts and asked: "{feedback}"
Apply that request, keep everything else that worked, and return all 6 sections again.

PREVIOUS DRAFTS:
{previous}
"""
    author_name = (st.session_state.get("author_name") or "").strip() or "the author"
    author_title = (st.session_state.get("author_title") or "").strip()
    author_line = f"{author_name}, {author_title}" if author_title else author_name
    rules_block = (
        f"\nBRAND RULES (HIGHEST PRIORITY. These override any default in this prompt, including names, titles and sign-offs):\n{rules.strip()}\n"
        if rules and rules.strip() else ""
    )
    return f"""You are an elite Executive Content Strategist writing for NewPath Business Advisors.
The content is published under this author: {author_line}. Write in their voice, and use exactly this name and title anywhere the author is named or signs off. Never use any other person's name as the author.
Turn the source material into 5 distinct, publish-ready formats. Rewrite natively for each platform; never copy-paste the source.
Only use facts present in the source. Never invent statistics, client names, or quotes.
{rules_block}
SETTINGS
- Depth: {detail}
- Audience: {audience}
- Tone: {tone_str}
- Call to action to weave in naturally: {cta or "None"}
{refine_block}
SOURCE MATERIAL
{source}

Return EXACTLY these 6 headers, each on its own line, in this order, with nothing before the first header:

### EDITORIAL_FEEDBACK
3 sentences: what works in the raw input, what could be sharper, how you positioned the drafts.

### LINKEDIN_POST
Strong hook, short scannable lines, a few bullets, under {LINKEDIN_LIMIT} characters, 3-4 hashtags at the end.

### EMAIL_NEWSLETTER
Subject line first, then 'Hi [First Name],', a personal story-driven body, one practical takeaway, signed {author_line}, NewPath Business Advisors (unless the brand rules say otherwise).

### EMAIL_SEQUENCE
A 3-email follow-up nurture series to send over about two weeks. For each email give: "Email 1 (Day 0)", "Email 2 (Day 4)", "Email 3 (Day 10)", then 'Subject:' and 'Preview text:' lines, then a short body (under 150 words) with one clear call to action. Escalate naturally: value, then proof or insight, then a soft ask. Sign each as {author_line}. Also give 2 alternative subject lines for Email 1 for A/B testing, labeled "A/B subject options".

### EXECUTIVE_BRIEFING
Markdown client handout: clear headings, strategic takeaways, and a 30-day action checklist using - [ ] items.

### CAROUSEL_SLIDES
5 slides, Slide 1 Cover Hook through Slide 5 Summary + CTA. Give each a short headline and 1-2 supporting lines.
"""


def parse_sections(text: str) -> dict:
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S | re.I).strip()
    pattern = r"^\s*#{1,4}\s*\**(" + "|".join(SECTIONS) + r")\**\s*:?\s*$"
    parts = re.split(pattern, text, flags=re.M)
    out = {s.lower(): "" for s in SECTIONS}
    for i in range(1, len(parts) - 1, 2):
        out[parts[i].lower()] = parts[i + 1].strip()
    return out


def generate(source, api_key, model, detail, audience, tones, rules, cta, feedback="", previous=""):
    if Groq is None:
        st.error("The `groq` package isn't installed. Run: pip install groq")
        return None
    if not api_key:
        st.error("Add your Groq API key in the sidebar (or set GROQ_API_KEY).")
        return None

    models, _ = fetch_live_models(api_key)
    candidates = [model] + [m for m in models if m != model][:3]  # automatic fallback
    prompt = build_prompt(source[:MAX_SOURCE_CHARS], detail, audience, tones, rules, cta, feedback, previous)
    client = Groq(api_key=api_key)
    last_err = None

    for m in candidates:
        try:
            resp = client.chat.completions.create(
                model=m,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_completion_tokens=DEPTH_TOKENS[detail],
            )
            sections = parse_sections(resp.choices[0].message.content or "")
            if not sections["linkedin_post"] and not sections["email_newsletter"]:
                last_err = RuntimeError("The model returned an unexpected format.")
                continue
            st.session_state.last_model_used = m
            return sections
        except Exception as e:
            last_err = e
            low = str(e).lower()
            if "invalid_api_key" in low or "401" in low:
                break  # no point trying other models
    st.error(friendly_error(last_err) if last_err else "Generation failed.")
    return None


def apply_result(result: dict):
    st.session_state.assets = result
    st.session_state.li_text = result["linkedin_post"]
    st.session_state.email_text = result["email_newsletter"]
    st.session_state.brief_text = result["executive_briefing"]
    st.session_state.slides_text = result["carousel_slides"]
    st.session_state.sequence_text = result["email_sequence"]
    st.session_state.check_report = ""  # drafts changed, so the old check is stale


def current_drafts_text() -> str:
    return (
        f"LINKEDIN_POST:\n{st.session_state.li_text}\n\n"
        f"EMAIL_NEWSLETTER:\n{st.session_state.email_text}\n\n"
        f"EMAIL_SEQUENCE:\n{st.session_state.sequence_text}\n\n"
        f"EXECUTIVE_BRIEFING:\n{st.session_state.brief_text}\n\n"
        f"CAROUSEL_SLIDES:\n{st.session_state.slides_text}"
    )


def call_model(prompt, api_key, model, max_tokens, temperature=0.2):
    """One-shot model call with automatic fallback to other live models."""
    if Groq is None or not api_key:
        st.error("Add your Groq API key in the sidebar first.")
        return None
    models, _ = fetch_live_models(api_key)
    candidates = [model] + [m for m in models if m != model][:3]
    client = Groq(api_key=api_key)
    last_err = None
    for m in candidates:
        try:
            resp = client.chat.completions.create(
                model=m,
                messages=[{"role": "user", "content": prompt}],
                temperature=temperature,
                max_completion_tokens=max_tokens,
            )
            text = re.sub(r"<think>.*?</think>", "", resp.choices[0].message.content or "", flags=re.S | re.I).strip()
            if text:
                return text
        except Exception as e:
            last_err = e
            if "invalid_api_key" in str(e).lower() or "401" in str(e):
                break
    st.error(friendly_error(last_err) if last_err else "The model returned nothing.")
    return None


def fact_check(source, drafts, api_key, model):
    prompt = f"""You are a meticulous fact-checker for a business advisory firm.
Compare the DRAFTS against the SOURCE. Find every specific, checkable claim in the drafts: numbers, percentages, dates, names, results, quotes, guarantees, and named outcomes.
Report ONLY claims that are not supported by the source (⚠️) or that contradict it (❌). Ignore generic advice, hashtags, calls to action, and placeholders like [First Name].

Output Markdown:
1. One-line verdict (e.g., "2 unsupported claims found" or "All specific claims are supported by the source").
2. If there are issues, a table with columns: Draft | Claim | Status | Why | Suggested fix.
Never invent issues. If everything is supported, say so and stop.

SOURCE:
{source[:MAX_SOURCE_CHARS]}

DRAFTS:
{drafts}
"""
    return call_model(prompt, api_key, model, 2500, temperature=0.1)


def _clean(s: str) -> str:
    s = re.sub(r"[*_`#]+", "", s)
    s = re.sub(r"^\s*[-•>]+\s*", "", s)
    s = re.sub(r"^(headline|title|supporting( text| line)?|subtext|body|cta)\s*:\s*", "", s, flags=re.I)
    return s.strip()


def parse_slides(text: str):
    parts = re.split(r"(?im)^[#*\s]*(slide\s*\d+[^\n]*)$", text)
    slides = []
    for i in range(1, len(parts) - 1, 2):
        label = _clean(re.sub(r"(?i)^[#*\s]*slide\s*\d+\s*[:\-–—.]?\s*", "", _clean(parts[i])))
        lines = [c for c in (_clean(x) for x in parts[i + 1].splitlines()) if c]
        slides.append((label, lines))
    if not slides:
        lines = [c for c in (_clean(x) for x in text.splitlines()) if c]
        slides = [("", lines)] if lines else []
    return slides


def build_pptx(slides_text: str, logo_bytes=None) -> bytes:
    NAVY, EMERALD, LIGHT, MUTED = "0F172A", "10B981", "E2E8F0", "94A3B8"
    prs = Presentation()
    prs.slide_width = prs.slide_height = Inches(7.5)  # square, LinkedIn carousel format
    name = (st.session_state.get("author_name") or "").strip()
    title = (st.session_state.get("author_title") or "").strip()
    footer = "NewPath Business Advisors" + (f"  |  {name}, {title}" if name and title else f"  |  {name}" if name else "")
    slides = parse_slides(slides_text)
    total = len(slides)

    def textbox(slide, x, y, w, h, paras, size, color, bold=False):
        tf = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
        tf.word_wrap = True
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = para.add_run()
            run.text = p
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor.from_string(color)
            para.space_after = Pt(8)

    for n, (label, lines) in enumerate(slides, start=1):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = RGBColor.from_string(NAVY)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.7), Inches(0.9), Inches(0.08))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor.from_string(EMERALD)
        bar.line.fill.background()
        if label:
            textbox(s, 0.6, 0.9, 5, 0.5, [label.upper()], 13, EMERALD, bold=True)
        headline = lines[0] if lines else ""
        body = lines[1:]
        size = (44 if n == 1 else 36) if len(headline) <= 60 else 28
        textbox(s, 0.6, 1.5, 6.3, 2.4, [headline], size, "FFFFFF", bold=True)
        if body:
            textbox(s, 0.6, 4.0, 6.3, 2.7, body if n == 1 else [f"• {b}" for b in body], 20, LIGHT)
        textbox(s, 0.6, 6.85, 5.2, 0.4, [footer], 11, MUTED)
        textbox(s, 6.1, 6.85, 0.8, 0.4, [f"{n} / {total}"], 11, MUTED)
        if logo_bytes:
            try:
                s.shapes.add_picture(io.BytesIO(logo_bytes), Inches(5.8), Inches(0.5), width=Inches(1.1))
            except Exception:
                pass
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_docx(li, email, seq, brief, slides) -> bytes:
    d = docx.Document()
    d.add_heading("NewPath Business Advisors: Content Bundle", level=0)
    d.add_paragraph(f"Generated {datetime.now():%B %d, %Y}")
    for title, body in [("LinkedIn Post", li), ("Email Newsletter", email), ("Email Sequence", seq),
                        ("Executive Briefing", brief), ("Carousel Slides", slides)]:
        d.add_heading(title, level=1)
        for line in body.splitlines():
            d.add_paragraph(re.sub(r"[#*]+\s?", "", line))
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------------
if not password_ok():
    st.stop()

st.sidebar.title("⚙️ Studio Settings")

stored_key = secret_key()
key_input = st.sidebar.text_input(
    "Groq API Key",
    type="password",
    key="api_key",
    placeholder="gsk_... (or set GROQ_API_KEY)" if not stored_key else "Using saved key",
    help="Free at console.groq.com/keys. Stored only in this session.",
)
api_key = (key_input or stored_key).strip()

models_list, models_live = (PREFERRED_MODELS, False)
if api_key and Groq is not None:
    models_list, models_live = fetch_live_models(api_key)
    if models_live:
        st.sidebar.success(f"Connected. {len(models_list)} live models found.")
    else:
        st.sidebar.error("Key could not be verified. Check it at console.groq.com/keys.")
elif Groq is None:
    st.sidebar.error("Install the Groq package: pip install groq")
else:
    st.sidebar.warning("Get a free key at console.groq.com/keys")

model_choice = st.sidebar.selectbox(
    "AI Model",
    models_list,
    help="Pulled live from Groq. If a model fails, the app automatically tries the next one.",
)
detail_level = st.sidebar.select_slider("Output Depth", ["Concise", "Standard", "Deep-Dive"], value="Standard")

st.sidebar.divider()

target_audience = st.sidebar.selectbox(
    "Target Audience",
    [
        "All Audiences / General Executive",
        "Small Business Owners ($1M–$10M Revenue)",
        "Founders Trapped in Day-to-Day Operations",
        "B2B Service & Agency Leaders",
        "Executive Coaching Prospects",
        "Internal Team & Stakeholders",
    ],
)
brand_tone = st.sidebar.multiselect(
    "Writing Style & Tone",
    ["Direct & No-Fluff", "Data-Driven (KPI Focused)", "Storytelling & Relatable", "Authoritative Executive", "Action-Oriented"],
    default=["Direct & No-Fluff", "Data-Driven (KPI Focused)"],
)
st.sidebar.text_input("Author Name", value="John Sedej", key="author_name",
                      help="Who the posts and emails are written as and signed by.")
st.sidebar.text_input("Author Title", value="CEO", key="author_title",
                      help="e.g., CEO, Senior Business IT Advisor")
cta_text = st.sidebar.text_input("Call to Action", placeholder="e.g., Book a free strategy call at newpath...")
custom_rules = st.sidebar.text_area(
    "Custom Brand Rules",
    placeholder="e.g., Keep hooks punchy. Never use the word 'synergy'.",
    height=80,
    help="Style and content rules. These take priority over everything else.",
)

logo_file = st.sidebar.file_uploader(
    "Brand logo for PowerPoint (optional)",
    type=["png", "jpg", "jpeg"],
    key=f"logo_{st.session_state.uploader_nonce}",
)

st.sidebar.divider()
st.sidebar.button(
    "🔄 Reset Everything",
    on_click=reset_all,
    use_container_width=True,
    help="Clears all settings, source material, uploads and drafts. Your API key is kept.",
)

# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
st.title(APP_NAME)
st.caption("Turn one set of notes into a LinkedIn post, newsletter, email sequence, client handout, and carousel in seconds.")

col_left, col_right = st.columns([1.15, 1], gap="large")

with col_left:
    st.subheader("1. Source Material")
    up = st.file_uploader(
        "Upload a file (.pdf, .docx, .txt, .md, .srt, .mp3, .m4a, .wav, .mp4)",
        type=["pdf", "docx", "txt", "md", "srt", "mp3", "m4a", "wav", "mp4"],
        key=f"uploader_{st.session_state.uploader_nonce}",
    )
    if up is not None:
        fid = (up.name, up.size)
        if fid != st.session_state.last_file_id:  # only extract once per new upload
            text = extract_file(up, api_key)
            if text:
                st.session_state.source_material = text
                st.session_state.last_file_id = fid

    st.text_area(
        "Or paste company updates, notes, or a transcript:",
        key="source_material",
        height=230,
        placeholder="Paste notes here...",
    )
    n_chars = len(st.session_state.source_material)
    st.caption(
        f"{len(st.session_state.source_material.split()):,} words"
        + (f" · only the first {MAX_SOURCE_CHARS:,} characters will be used" if n_chars > MAX_SOURCE_CHARS else "")
    )

    if st.button("⚡ Generate Content", use_container_width=True):
        if not st.session_state.source_material.strip():
            st.error("Paste text or upload a file first.")
        else:
            with st.spinner("Writing your drafts..."):
                result = generate(
                    st.session_state.source_material, api_key, model_choice,
                    detail_level, target_audience, brand_tone, custom_rules, cta_text,
                )
            if result:
                apply_result(result)
                st.session_state.chat_history = [
                    {"role": "assistant", "content": result["editorial_feedback"] or "Drafts are ready."}
                ]

with col_right:
    st.subheader("2. Editorial Coach")
    if st.session_state.assets is None:
        st.info("Add your notes and click **Generate Content**. Then chat here to refine any draft.")
    else:
        if st.session_state.last_model_used:
            st.markdown(f'<span class="np-badge">Model: {st.session_state.last_model_used}</span>', unsafe_allow_html=True)

        quick = {
            "🔥 Punchier hooks": "Make the hooks punchier and more scroll-stopping.",
            "✂️ Shorter": "Cut every draft by about 30% without losing the key points.",
            "🎯 Stronger CTA": "Make the call to action clearer and more compelling.",
            "🤝 Warmer tone": "Make the tone warmer and more personal.",
        }
        qcols = st.columns(len(quick))
        quick_request = None
        for c, (label, prompt_text) in zip(qcols, quick.items()):
            if c.button(label, use_container_width=True):
                quick_request = prompt_text

        box = st.container(height=260)
        with box:
            for msg in st.session_state.chat_history:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])

        typed = st.chat_input("Ask for a change, e.g. 'Make the email less formal'")
        request = typed or quick_request
        if request:
            st.session_state.chat_history.append({"role": "user", "content": request})
            with st.spinner("Updating drafts..."):
                updated = generate(
                    st.session_state.source_material, api_key, model_choice, detail_level,
                    target_audience, brand_tone, custom_rules, cta_text,
                    feedback=request, previous=current_drafts_text(),
                )
            if updated:
                apply_result(updated)
                st.session_state.chat_history.append(
                    {"role": "assistant", "content": f"✅ Updated.\n\n{updated['editorial_feedback']}"}
                )
            st.rerun()

if st.session_state.assets is not None:
    st.divider()
    st.subheader("3. Deliverables")

    tab_li, tab_email, tab_seq, tab_brief, tab_slides, tab_check, tab_export = st.tabs(
        ["💼 LinkedIn", "📧 Newsletter", "📨 Email Sequence", "📋 Executive Handout", "🎠 Carousel",
         "🔍 Source Check", "📥 Export"]
    )

    with tab_li:
        st.text_area("LinkedIn Post", key="li_text", height=320)
        n = len(st.session_state.li_text)
        (st.warning if n > LINKEDIN_LIMIT else st.caption)(f"{n:,} / {LINKEDIN_LIMIT:,} characters")
    with tab_email:
        st.text_area("Email Newsletter", key="email_text", height=340)
        st.caption(f"{len(st.session_state.email_text.split()):,} words")
    with tab_seq:
        st.text_area("3-Email Follow-Up Sequence (with A/B subject options)", key="sequence_text", height=420)
        st.caption("Send Day 0, Day 4, and Day 10. Paste each into your email tool.")
    with tab_brief:
        edit = st.toggle("Edit mode", key="brief_edit")
        if edit:
            st.text_area("Executive Briefing (Markdown)", key="brief_text", height=360)
        else:
            st.markdown(st.session_state.brief_text)
    with tab_slides:
        edit = st.toggle("Edit mode", key="slides_edit")
        if edit:
            st.text_area("Carousel Slides", key="slides_text", height=320)
        else:
            st.markdown(st.session_state.slides_text)
        if HAS_PPTX:
            try:
                st.download_button(
                    "⬇️ Download branded carousel (.pptx)",
                    build_pptx(st.session_state.slides_text, logo_file.getvalue() if logo_file else None),
                    "NewPath_Carousel.pptx",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception as e:
                st.warning(f"Could not build the PowerPoint: {e}")
        else:
            st.caption("Install `python-pptx` to enable PowerPoint export.")
    with tab_check:
        st.write("Scans every draft for numbers, names, results, or promises that **aren't in your source notes**, "
                 "so nothing unverified gets published under your name.")
        if st.button("🔍 Run source check", key="run_check", use_container_width=True):
            with st.spinner("Checking drafts against your source..."):
                report = fact_check(st.session_state.source_material, current_drafts_text(), api_key, model_choice)
            if report:
                st.session_state.check_report = report
        if st.session_state.check_report:
            st.markdown(st.session_state.check_report)
            st.caption("Re-run after you edit or refine the drafts. This is an AI check, not a guarantee. Review before publishing.")
    with tab_export:
        li, em = st.session_state.li_text, st.session_state.email_text
        br, sl = st.session_state.brief_text, st.session_state.slides_text
        sq = st.session_state.sequence_text
        md = (f"# NewPath Content Bundle\n_Generated {datetime.now():%B %d, %Y}_\n\n"
              f"## LinkedIn Post\n{li}\n\n## Email Newsletter\n{em}\n\n"
              f"## Email Sequence\n{sq}\n\n"
              f"## Executive Briefing\n{br}\n\n## Carousel Slides\n{sl}\n")
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Download bundle (.md)", md, "NewPath_Content_Bundle.md",
                           "text/markdown", use_container_width=True)
        if docx:
            c2.download_button(
                "⬇️ Download bundle (.docx)", build_docx(li, em, sq, br, sl), "NewPath_Content_Bundle.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            c2.caption("Install `python-docx` to enable Word export.")
        st.download_button("⬇️ LinkedIn post only (.txt)", li, "linkedin_post.txt", use_container_width=True)
        st.download_button("⬇️ Email only (.txt)", em, "newsletter.txt", use_container_width=True)
        st.download_button("⬇️ Email sequence only (.txt)", sq, "email_sequence.txt", use_container_width=True)

st.divider()
st.caption("AI-generated drafts. Review for accuracy and brand voice before publishing.")
