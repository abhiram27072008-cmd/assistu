import hashlib
import io
import json
import os
import re
import uuid

import streamlit as st
from docx import Document
from fpdf import FPDF
from google import genai
from google.genai import types
from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from streamlit_js_eval import streamlit_js_eval

# ---------------------------------------------------------------- settings
CHAT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-3.5-flash-lite"  # used if the main model is busy
MAX_CHARS = 100_000  # max text read from one file
STORAGE_KEY = "assistu_chats_v1"  # name of the saved data in the browser
MAX_SAVED_BYTES = 4_000_000  # browser storage is small, keep under this
SLIDES_PER_CALL = 5  # long presentations are written in small batches

SYSTEM_PROMPT = ("You are AssistU, a friendly assistant for students. "
                 "Explain simply with examples. "
                 "When files or photos are attached, read them carefully "
                 "and answer based on their content.")
DOC_SYSTEM = ("You are AssistU, helping a student create study material and presentations. "
              "Be accurate, clear and well organised.")

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

st.set_page_config(page_title="AssistU", page_icon="🎓")
st.title("AssistU")
st.caption("Your AI study assistant")

try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    st.error("API key not found. Add GEMINI_API_KEY to your Streamlit secrets.")
    st.stop()
client = genai.Client(api_key=API_KEY)


def is_busy_or_missing(e) -> bool:
    err = str(e)
    return "503" in err or "UNAVAILABLE" in err or "404" in err or "NOT_FOUND" in err


def show_error(e):
    err = str(e)
    if "429" in err or "RESOURCE_EXHAUSTED" in err:
        st.warning("⏳ The free usage limit was reached. Please wait a few minutes "
                   "(or until tomorrow) and try again.")
    elif "503" in err or "UNAVAILABLE" in err:
        st.warning("⚠️ Google's AI is very busy right now. Please try again in a minute.")
    else:
        st.error(f"Something went wrong: {e}")


def stream_reply(contents):
    """Yield the answer piece by piece. If the main model is busy (503),
    switch to the backup model straight away."""
    config = types.GenerateContentConfig(system_instruction=SYSTEM_PROMPT)
    first_err = None
    for model in (CHAT_MODEL, FALLBACK_MODEL):
        started = False
        try:
            for chunk in client.models.generate_content_stream(
                    model=model, contents=contents, config=config):
                if chunk.text:
                    started = True
                    yield chunk.text
            return
        except Exception as e:
            if first_err is None:
                first_err = e
            if started:
                raise  # already showing text: cannot switch models now
            if model == CHAT_MODEL and is_busy_or_missing(e):
                continue  # main model busy or missing: try the backup
            raise first_err
    raise first_err


# ---------------------------------------------------------------- reading files
def read_docx(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
    lines = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            lines.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(lines)


def read_xlsx(data: bytes) -> str:
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"--- Sheet: {ws.title} ---")
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 1000:
                out.append("... (more rows not shown)")
                break
            if any(c is not None for c in row):
                out.append(" | ".join("" if c is None else str(c) for c in row))
    return "\n".join(out)


def read_pptx(data: bytes) -> str:
    prs = Presentation(io.BytesIO(data))
    out = []
    for n, slide in enumerate(prs.slides, 1):
        out.append(f"--- Slide {n} ---")
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                out.append(shape.text_frame.text)
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    out.append(" | ".join(c.text for c in row.cells))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            out.append("Notes: " + slide.notes_slide.notes_text_frame.text)
    return "\n".join(out)


def build_parts(uploaded_files, camera_photo):
    """Turn attachments into Gemini parts. Returns (parts, names, images)."""
    parts, names, images = [], [], []

    for f in uploaded_files or []:
        data = f.getvalue()
        ext = f.name.lower().rsplit(".", 1)[-1]
        try:
            if ext in ("png", "jpg", "jpeg", "webp"):
                mime = f.type or f"image/{'jpeg' if ext == 'jpg' else ext}"
                parts.append(types.Part.from_bytes(data=data, mime_type=mime))
                images.append(data)
            elif ext == "pdf":
                # Gemini reads PDFs directly (scanned PDFs work too)
                parts.append(types.Part.from_bytes(data=data, mime_type="application/pdf"))
            else:
                if ext == "docx":
                    text = read_docx(data)
                elif ext == "xlsx":
                    text = read_xlsx(data)
                elif ext == "pptx":
                    text = read_pptx(data)
                else:  # txt, csv, md
                    text = data.decode("utf-8", errors="replace")
                parts.append(types.Part.from_text(
                    text=f"[Contents of file: {f.name}]\n{text[:MAX_CHARS]}"))
            names.append(f.name)
        except Exception as e:
            st.warning(f"Could not read {f.name}: {e}")

    if camera_photo is not None:
        data = camera_photo.getvalue()
        parts.append(types.Part.from_bytes(data=data, mime_type="image/jpeg"))
        images.append(data)
        names.append("camera photo")

    return parts, names, images


# ---------------------------------------------------------------- PDF helpers
PDF_REPLACE = {
    "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"', "\u2013": "-", "\u2014": "-",
    "\u2212": "-", "\u2022": "-", "\u2026": "...", "\u2192": "->", "\u2190": "<-",
    "\u2248": "~", "\u2264": "<=", "\u2265": ">=", "\u2260": "!=", "\u00a0": " ",
    "\u03c0": "pi", "\u03b1": "alpha", "\u03b2": "beta", "\u03b3": "gamma", "\u0394": "Delta",
    "\u03bc": "mu", "\u03b8": "theta", "\u03bb": "lambda", "\u03c3": "sigma", "\u03a9": "Ohm",
    "\u221e": "infinity", "\u221a": "sqrt", "\u2211": "sum", "\u2713": "[x]",
}


def pdf_safe(text: str) -> str:
    """The built-in PDF fonts only know basic Latin letters, so swap common symbols."""
    for old, new in PDF_REPLACE.items():
        text = text.replace(old, new)
    return text.encode("latin-1", "replace").decode("latin-1")


def _line(pdf, height, text):
    pdf.multi_cell(0, height, text, new_x="LMARGIN", new_y="NEXT")


@st.cache_data(show_spinner=False)
def text_to_pdf(text: str, title: str = "AssistU reply") -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    _line(pdf, 10, pdf_safe(title))
    pdf.ln(2)
    pdf.set_font("Helvetica", size=12)
    _line(pdf, 7, pdf_safe(text.replace("**", "").replace("`", "")))
    return bytes(pdf.output())


def build_report_pdf(title, sections):
    """sections = [(heading, body_text), ...]. Returns (pdf_bytes, page_count)."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 24)
    _line(pdf, 12, pdf_safe(title))
    pdf.set_font("Helvetica", "I", 11)
    _line(pdf, 8, "Created with AssistU")
    for heading, body in sections:
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 16)
        _line(pdf, 9, pdf_safe(heading))
        pdf.ln(1)
        for raw in body.replace("\r", "").split("\n"):
            line = pdf_safe(raw.strip()).replace("**", "").replace("`", "")
            if not line:
                pdf.ln(2)
            elif line.startswith("#"):
                pdf.set_font("Helvetica", "B", 13)
                _line(pdf, 7.5, line.lstrip("# ").strip())
            elif line[:2] in ("- ", "* "):
                pdf.set_font("Helvetica", "", 12)
                pdf.set_x(pdf.l_margin + 5)
                _line(pdf, 6.5, "- " + line[2:].strip())
            else:
                pdf.set_font("Helvetica", "", 12)
                _line(pdf, 6.5, line)
    pages = pdf.page_no()
    return bytes(pdf.output()), pages


# ---------------------------------------------------------------- PowerPoint builder
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
ACCENT = RGBColor(0xF2, 0xA9, 0x1E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xD6, 0xDE, 0xEA)
DARK = RGBColor(0x2B, 0x2B, 0x2B)
GREY = RGBColor(0x80, 0x80, 0x80)


def _rect(slide, x, y, w, h, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _text(slide, x, y, w, h, text, size, color, bold=False, align=None, anchor=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def build_pptx(title, subtitle, slides) -> bytes:
    """slides = [{"title", "bullets", "notes"}, ...] (the title slide is added first)."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    _rect(s, 0, 0, W, H, NAVY)
    _rect(s, Inches(0.9), Inches(4.0), Inches(2.2), Inches(0.08), ACCENT)
    _text(s, Inches(0.9), Inches(1.4), W - Inches(1.8), Inches(2.5), title,
          44 if len(title) < 50 else 34, WHITE, bold=True, anchor=MSO_ANCHOR.BOTTOM)
    _text(s, Inches(0.9), Inches(4.3), W - Inches(1.8), Inches(1.2), subtitle, 22, LIGHT)

    for number, sl in enumerate(slides, start=2):
        s = prs.slides.add_slide(blank)
        _rect(s, 0, 0, W, Inches(1.3), NAVY)
        _rect(s, 0, Inches(1.3), W, Inches(0.06), ACCENT)
        t = sl["title"]
        _text(s, Inches(0.6), Inches(0.15), W - Inches(1.2), Inches(1.0), t,
              30 if len(t) <= 45 else 24, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)

        bullets = [b.replace("**", "").strip() for b in sl["bullets"] if b.strip()]
        total = sum(len(b) for b in bullets)
        size = 24 if total < 350 else 22 if total < 500 else 20 if total < 650 else 18 if total < 800 else 16
        body = s.shapes.add_textbox(Inches(0.8), Inches(1.75), W - Inches(1.6), Inches(5.0))
        tf = body.text_frame
        tf.word_wrap = True
        for j, b in enumerate(bullets):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.space_after = Pt(10)
            run = p.add_run()
            run.text = "\u2022  " + b
            run.font.size = Pt(size)
            run.font.color.rgb = DARK
        _text(s, W - Inches(1.4), H - Inches(0.65), Inches(0.9), Inches(0.4), str(number),
              12, GREY, align=PP_ALIGN.RIGHT)
        if sl.get("notes"):
            s.notes_slide.notes_text_frame.text = sl["notes"]

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- creating long PPTs / PDFs
def gen(contents, json_mode=False):
    """One Gemini call (not streamed). Tries the backup model if the main one is busy."""
    kwargs = {"system_instruction": DOC_SYSTEM}
    if json_mode:
        kwargs["response_mime_type"] = "application/json"
    config = types.GenerateContentConfig(**kwargs)
    first_err = None
    for model in (CHAT_MODEL, FALLBACK_MODEL):
        try:
            resp = client.models.generate_content(model=model, contents=contents, config=config)
            return resp.text or ""
        except Exception as e:
            if first_err is None:
                first_err = e
            if model == CHAT_MODEL and is_busy_or_missing(e):
                continue  # main model busy or missing: try the backup
            raise first_err
    raise first_err


def parse_json(text):
    t = re.sub(r"^```(?:json)?|```$", "", (text or "").strip(), flags=re.M).strip()
    starts = [i for i in (t.find("{"), t.find("[")) if i != -1]
    if not starts:
        raise ValueError("The AI did not return a valid plan.")
    end = max(t.rfind("}"), t.rfind("]"))
    return json.loads(t[min(starts):end + 1])


def gen_json(contents, tries=2):
    last = None
    for _ in range(tries):
        try:
            return parse_json(gen(contents, json_mode=True))
        except ValueError as e:  # bad JSON: ask once more
            last = e
    raise last


def safe_name(title, ext):
    base = re.sub(r"[^A-Za-z0-9]+", "_", title).strip("_")[:40] or "AssistU"
    return f"{base}.{ext}"


def create_pptx(topic, n, src_parts, report):
    def ask(prompt):
        return src_parts + [types.Part.from_text(text=prompt)]

    report(0.03, "Planning the slides...")
    plan = gen_json(ask(
        f"Plan a presentation about: {topic}\n"
        f"Make exactly {n} slides, and the first one is the title slide. Write in the "
        "same language as the topic. Return ONLY JSON like "
        '{"title": "...", "subtitle": "...", "slides": [{"title": "..."}]} '
        f"where slides lists all {n} slides in order."))
    title = str(plan.get("title") or topic[:60])
    subtitle = str(plan.get("subtitle") or "Created with AssistU")
    titles = [str(s.get("title", "")) if isinstance(s, dict) else str(s)
              for s in plan.get("slides", [])]
    titles = (titles + [f"Slide {i + 1}" for i in range(len(titles), n)])[:n]

    slides, partial, error = [], False, None
    idxs = list(range(1, n))  # slide 0 is the title slide
    try:
        for b in range(0, len(idxs), SLIDES_PER_CALL):
            chunk = idxs[b:b + SLIDES_PER_CALL]
            report(0.08 + 0.87 * b / max(len(idxs), 1),
                   f"Writing slides {chunk[0] + 1} to {chunk[-1] + 1} of {n}...")
            items = gen_json(ask(
                f'Presentation: "{title}" about {topic}.\n'
                f"All slide titles in order: {json.dumps(titles)}.\n"
                f"Write the content for slides {chunk[0] + 1} to {chunk[-1] + 1} only "
                "(slide 1 is the title slide). For each slide give 4 to 6 bullet points of "
                "at most 20 words each, factual and useful, plus 2 or 3 sentences of speaker "
                f"notes. Return ONLY a JSON list with exactly {len(chunk)} items like "
                '[{"title": "...", "bullets": ["..."], "notes": "..."}]'))
            if isinstance(items, dict):
                items = items.get("slides", [])
            for k, i in enumerate(chunk):
                it = items[k] if k < len(items) and isinstance(items[k], dict) else {}
                slides.append({"title": str(it.get("title") or titles[i]),
                               "bullets": [str(x) for x in it.get("bullets", [])][:8],
                               "notes": str(it.get("notes", ""))})
    except Exception as e:
        if not slides:
            raise
        partial, error = True, e

    report(0.97, "Building the PowerPoint file...")
    return {"bytes": build_pptx(title, subtitle, slides), "title": title, "mime": PPTX_MIME,
            "name": safe_name(title, "pptx"), "partial": partial, "error": error,
            "count": len(slides) + 1, "total": n,
            "outline": [titles[0]] + [s["title"] for s in slides]}


def create_pdf_doc(topic, n, src_parts, report):
    def ask(prompt):
        return src_parts + [types.Part.from_text(text=prompt)]

    report(0.03, "Planning the document...")
    plan = gen_json(ask(
        f"Plan a document about: {topic}\n"
        f"It should be about {n} pages long. Return ONLY JSON like "
        '{"title": "...", "sections": [{"heading": "...", "covers": "one sentence"}]} '
        f"with exactly {n} sections in a logical order (introduction first, conclusion last). "
        "Write in English."))
    title = str(plan.get("title") or topic[:60])
    secs = [s for s in plan.get("sections", []) if isinstance(s, dict)][:n] or \
           [{"heading": "Overview", "covers": topic}]
    headings = [str(s.get("heading", f"Section {i + 1}")) for i, s in enumerate(secs)]

    done, partial, error = [], False, None
    try:
        for i, s in enumerate(secs):
            report(0.08 + 0.87 * i / len(secs), f"Writing section {i + 1} of {len(secs)}...")
            body = gen(ask(
                f'Document: "{title}" about {topic}.\n'
                f"All sections in order: {json.dumps(headings)}.\n"
                f'Write ONLY section {i + 1}, "{headings[i]}" (it covers: {s.get("covers", "")}). '
                "Length: about 350 words. Use plain text only: short paragraphs, and lines "
                'starting with "- " for bullet points. Do not use markdown symbols like ** or #, '
                "do not repeat the section heading, and use only simple English letters and "
                "punctuation (write maths in words)."))
            done.append((headings[i], body))
    except Exception as e:
        if not done:
            raise
        partial, error = True, e

    report(0.97, "Building the PDF file...")
    pdf_bytes, pages = build_report_pdf(title, done)
    return {"bytes": pdf_bytes, "title": title, "mime": "application/pdf",
            "name": safe_name(title, "pdf"), "partial": partial, "error": error,
            "count": pages, "total": n, "outline": [h for h, _ in done]}


# ---------------------------------------------------------------- saving chats in the browser
def build_payload() -> str:
    """Text of all non-empty chats (photos/files are not saved) as JSON."""
    chats = {
        cid: {"title": c["title"],
              "history": [{"role": m["role"], "text": m["text"], "files": m["files"]}
                          for m in c["history"]]}
        for cid, c in st.session_state.chats.items() if c["history"]
    }
    payload = json.dumps({"chats": chats})
    while len(payload) > MAX_SAVED_BYTES and len(chats) > 1:
        chats.pop(next(iter(chats)))  # drop the oldest chat
        payload = json.dumps({"chats": chats})
    return payload


def restore_chats(raw) -> dict:
    """Read saved chats back from the browser."""
    try:
        stored = json.loads(raw).get("v")
        data = json.loads(stored).get("chats", {}) if stored else {}
    except Exception:
        return {}
    chats = {}
    for cid, c in data.items():
        history = []
        for m in c.get("history", []):
            if m.get("role") not in ("user", "model"):
                continue
            text, files = m.get("text", ""), m.get("files", [])
            model_text = text
            if m.get("role") == "user" and files:
                model_text += (f"\n(Files attached earlier: {', '.join(files)}. "
                               "Their contents are no longer available.)")
            history.append({"role": m["role"], "text": text, "files": files,
                            "images": [], "parts": [types.Part.from_text(text=model_text)]})
        if history:
            chats[cid] = {"title": c.get("title", "Chat"), "history": history}
    return chats


# ---------------------------------------------------------------- state
for name, default in (("uploader_key", 0), ("chats", {}), ("current", None),
                      ("loaded", False), ("save_ok", False)):
    if name not in st.session_state:
        st.session_state[name] = default

# ask the browser for the saved chats (first run gives None, then it reruns)
saved_raw = streamlit_js_eval(
    js_expressions=f"JSON.stringify({{v: localStorage.getItem('{STORAGE_KEY}')}})",
    key="ls_load",
)
if not st.session_state.loaded:
    if saved_raw is None:
        st.info("⏳ Loading your chats...")
        if st.button("Continue without saved chats"):
            st.session_state.loaded = True  # saving stays off so nothing is overwritten
            st.rerun()
        st.stop()
    st.session_state.chats = restore_chats(saved_raw)
    st.session_state.loaded = True
    st.session_state.save_ok = True


def new_chat():
    cur = st.session_state.chats.get(st.session_state.current)
    st.session_state.uploader_key += 1  # clear any attachments
    if cur is not None and not cur["history"]:
        return  # already on an empty chat
    cid = uuid.uuid4().hex[:8]
    st.session_state.chats[cid] = {"title": "New chat", "history": []}
    st.session_state.current = cid


if st.session_state.current not in st.session_state.chats:
    new_chat()

# ---------------------------------------------------------------- side bar: chat list
with st.sidebar:
    st.header("💬 Chats")
    if st.button("➕ New chat", use_container_width=True):
        new_chat()
        st.rerun()
    st.divider()
    for cid in reversed(list(st.session_state.chats)):
        c = st.session_state.chats[cid]
        is_current = cid == st.session_state.current
        if not c["history"] and not is_current:
            continue  # hide empty chats
        col_open, col_del = st.columns([5, 1])
        label = ("🟢 " if is_current else "") + c["title"]
        if col_open.button(label, key=f"open_{cid}", use_container_width=True):
            st.session_state.current = cid
            st.session_state.uploader_key += 1
            st.rerun()
        if col_del.button("🗑️", key=f"del_{cid}", help="Delete this chat"):
            del st.session_state.chats[cid]
            if cid == st.session_state.current:
                st.session_state.current = None
                new_chat()
            st.rerun()

chat = st.session_state.chats[st.session_state.current]
history = chat["history"]

notice = st.session_state.pop("notice", None)
if notice:
    st.warning(notice)

# ---------------------------------------------------------------- tools (main screen)
with st.expander("📎 Upload files  •  📷 Camera", expanded=False):
    st.caption("Photos, PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx), txt, csv")
    uploaded = st.file_uploader(
        "Upload files",
        type=["png", "jpg", "jpeg", "webp", "pdf", "docx", "xlsx", "pptx", "txt", "csv", "md"],
        accept_multiple_files=True,
        key=f"up_{st.session_state.uploader_key}",
    )
    use_camera = st.checkbox("📷 Use camera")
    photo = None
    if use_camera:
        photo = st.camera_input("Take a picture", key=f"cam_{st.session_state.uploader_key}")

with st.expander("📊 Create a PPT or PDF", expanded=False):
    kind = st.radio("What do you want to create?", ["PowerPoint (PPT)", "PDF document"],
                    horizontal=True)
    topic = st.text_area("What should it be about?", height=100,
                         key=f"topic_{st.session_state.uploader_key}",
                         placeholder="Example: How solar panels work, for first-year students")
    if kind.startswith("PowerPoint"):
        amount = st.slider("Number of slides", 3, 40, 10)
    else:
        amount = st.slider("Approximate number of pages", 2, 25, 8)
    st.caption("Tip: attach files in the box above to use them as the source. "
               "Long ones take 1 to 3 minutes, so keep this page open. "
               "PDFs support English letters only.")
    make = st.button("✨ Create", type="primary")

if st.button("🗑️ Clear this chat"):
    chat["history"] = []
    chat["title"] = "New chat"
    st.session_state.uploader_key += 1
    st.rerun()

n_attached = len(uploaded or []) + (1 if photo is not None else 0)
if n_attached:
    st.info(f"📎 {n_attached} attachment(s) ready. They will be sent with your next message.")

# ---------------------------------------------------------------- create PPT / PDF
if make:
    src_parts, src_names, _ = build_parts(uploaded, photo)
    if not topic.strip() and not src_parts:
        st.warning("✏️ Please type a topic first, or attach files to use as the source.")
    else:
        topic_text = topic.strip() or "the attached files"
        bar = st.progress(0.0, text="Starting...")

        def report(fraction, message):
            bar.progress(min(max(fraction, 0.0), 1.0), text=message)

        is_ppt = kind.startswith("PowerPoint")
        try:
            if is_ppt:
                res = create_pptx(topic_text, amount, src_parts, report)
            else:
                res = create_pdf_doc(topic_text, amount, src_parts, report)
        except Exception as e:
            bar.empty()
            show_error(e)
            st.stop()
        bar.empty()

        what = "presentation" if is_ppt else "PDF"
        unit = "slides" if is_ppt else "pages"
        user_text = (f"Create a {amount}-{'slide PowerPoint' if is_ppt else 'page PDF'}: "
                     f"{topic_text}")
        reply_text = "\n".join(
            [f"✅ Your {what} is ready: **{res['title']}** ({res['count']} {unit})", ""] +
            [f"{i}. {t}" for i, t in enumerate(res["outline"], 1)])
        history.append({"role": "user", "text": user_text, "files": src_names, "images": [],
                        "parts": [types.Part.from_text(text=user_text)]})
        history.append({"role": "model", "text": reply_text, "files": [res["name"]],
                        "images": [], "parts": [types.Part.from_text(text=reply_text)],
                        "file": {"name": res["name"], "bytes": res["bytes"], "mime": res["mime"]}})
        if len(history) == 2:
            chat["title"] = topic_text[:28] + ("…" if len(topic_text) > 28 else "")
        if res["partial"]:
            st.session_state.notice = (
                f"⚠️ Stopped early ({res['error']}). Your file has {res['count']} of "
                f"{res['total']} {unit}. Try again later for the full version.")
        st.session_state.uploader_key += 1
        st.rerun()

# ---------------------------------------------------------------- show chat
for i, m in enumerate(history):
    with st.chat_message("user" if m["role"] == "user" else "assistant"):
        for img in m["images"]:
            st.image(img, width=300)
        if m["files"]:
            st.caption("📎 " + ", ".join(m["files"]))
        if m["text"]:
            st.markdown(m["text"])
        if m.get("file"):
            f = m["file"]
            st.download_button(f"⬇️ Download {f['name']}", f["bytes"], file_name=f["name"],
                               mime=f["mime"], key=f"file_{st.session_state.current}_{i}")
        elif m["role"] == "model" and m["text"]:
            st.download_button("⬇️ PDF", text_to_pdf(m["text"]),
                               file_name=f"assistu_reply_{i}.pdf",
                               mime="application/pdf",
                               key=f"pdf_{st.session_state.current}_{i}")

# ---------------------------------------------------------------- chat input
prompt = st.chat_input("Ask AssistU anything...")

if prompt:
    parts, names, images = build_parts(uploaded, photo)

    history.append({
        "role": "user", "text": prompt, "files": names, "images": images,
        "parts": parts + [types.Part.from_text(text=prompt)],
    })
    if len(history) == 1:  # first message names the chat
        chat["title"] = prompt[:28] + ("…" if len(prompt) > 28 else "")

    # show the user's message right away
    with st.chat_message("user"):
        for img in images:
            st.image(img, width=300)
        if names:
            st.caption("📎 " + ", ".join(names))
        st.markdown(prompt)

    try:
        contents = [types.Content(role=m["role"], parts=m["parts"]) for m in history]
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                reply = st.write_stream(stream_reply(contents))  # text appears as it is written
        reply = reply or "(no response)"
        history.append({
            "role": "model", "text": reply, "files": [], "images": [],
            "parts": [types.Part.from_text(text=reply)],
        })
    except Exception as e:
        history.pop()  # remove the failed message
        if not history:
            chat["title"] = "New chat"
        show_error(e)
        st.stop()

    # clear attachments so they are not sent again with the next message
    st.session_state.uploader_key += 1
    st.rerun()

# ---------------------------------------------------------------- save to the browser
# Runs on every normal page run. It writes only when the chats have changed.
if st.session_state.save_ok:
    payload = build_payload()
    save_js = f"localStorage.setItem('{STORAGE_KEY}', {json.dumps(payload)})"
    save_key = "ls_save_" + hashlib.md5(payload.encode("utf-8")).hexdigest()[:12]
else:
    save_js, save_key = "0", "ls_save_off"
streamlit_js_eval(js_expressions=save_js, key=save_key)
