import io
import os
import time

import streamlit as st
from docx import Document
from fpdf import FPDF
from google import genai
from google.genai import types
from openpyxl import load_workbook
from pptx import Presentation

# ---------------------------------------------------------------- settings
CHAT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-2.5-flash"  # used if the main model is busy
MAX_CHARS = 100_000  # max text read from one file

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


def generate_with_retry(model, contents, config=None, fallback=None, tries=3):
    """Retry when Google says the model is busy (503), then try a backup model."""
    last_err = None
    for m in [model] + ([fallback] if fallback else []):
        for attempt in range(tries):
            try:
                return client.models.generate_content(model=m, contents=contents, config=config)
            except Exception as e:
                last_err = e
                busy = "503" in str(e) or "UNAVAILABLE" in str(e)
                if not busy:
                    raise
                time.sleep(2 * (attempt + 1))
    raise last_err


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


# ---------------------------------------------------------------- PDF download
def text_to_pdf(text: str, title: str = "AssistU reply") -> bytes:
    def latin(t):  # built-in PDF fonts only support latin-1
        return t.encode("latin-1", "replace").decode("latin-1")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.multi_cell(0, 10, latin(title))
    pdf.ln(2)
    pdf.set_font("Helvetica", size=12)
    pdf.multi_cell(0, 7, latin(text.replace("**", "").replace("`", "")))
    return bytes(pdf.output())


# ---------------------------------------------------------------- state
if "history" not in st.session_state:
    st.session_state.history = []  # dicts: role, text, parts, images, files
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

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

if st.button("🗑️ Clear chat"):
    st.session_state.history = []
    st.session_state.uploader_key += 1
    st.rerun()

n_attached = len(uploaded or []) + (1 if photo is not None else 0)
if n_attached:
    st.info(f"📎 {n_attached} attachment(s) ready. They will be sent with your next message.")

# ---------------------------------------------------------------- show chat
for i, m in enumerate(st.session_state.history):
    with st.chat_message("user" if m["role"] == "user" else "assistant"):
        for img in m["images"]:
            st.image(img, width=300)
        if m["files"]:
            st.caption("📎 " + ", ".join(m["files"]))
        if m["text"]:
            st.markdown(m["text"])
        if m["role"] == "model" and m["text"]:
            st.download_button("⬇️ PDF", text_to_pdf(m["text"]),
                               file_name=f"assistu_reply_{i}.pdf",
                               mime="application/pdf", key=f"pdf_{i}")

# ---------------------------------------------------------------- chat input
prompt = st.chat_input("Ask AssistU anything...")

if prompt:
    parts, names, images = build_parts(uploaded, photo)

    st.session_state.history.append({
        "role": "user", "text": prompt, "files": names, "images": images,
        "parts": parts + [types.Part.from_text(text=prompt)],
    })

    try:
        contents = [types.Content(role=m["role"], parts=m["parts"])
                    for m in st.session_state.history]
        with st.spinner("Thinking..."):
            resp = generate_with_retry(
                model=CHAT_MODEL,
                fallback=FALLBACK_MODEL,
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction="You are AssistU, a friendly assistant for students. "
                                       "Explain simply with examples. "
                                       "When files or photos are attached, read them carefully "
                                       "and answer based on their content."
                ),
            )
        reply = resp.text or "(no response)"
        st.session_state.history.append({
            "role": "model", "text": reply, "files": [], "images": [],
            "parts": [types.Part.from_text(text=reply)],
        })
    except Exception as e:
        st.session_state.history.pop()  # remove the failed message
        err = str(e)
        if "429" in err or "RESOURCE_EXHAUSTED" in err:
            st.warning("⏳ The free usage limit was reached. Please wait a few minutes "
                       "(or until tomorrow) and try again.")
        elif "503" in err or "UNAVAILABLE" in err:
            st.warning("⚠️ Google's AI is very busy right now. Please try again in a minute.")
        else:
            st.error(f"Something went wrong: {e}")
        st.stop()

    # clear attachments so they are not sent again with the next message
    st.session_state.uploader_key += 1
    st.rerun()
