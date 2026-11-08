import io
import os

import streamlit as st
from fpdf import FPDF
from google import genai
from google.genai import types
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pypdf import PdfReader
from docx import Document

# ---------------------------------------------------------------- config
CHAT_MODEL = "gemini-3.8-flash"
IMAGE_MODEL = "gemini-2.5-flash-image"
MAX_CHARS = 100_000  # cap on text pulled from one file

st.set_page_config(page_title="AssistU", page_icon="🎓")
st.title("AssistU")
st.caption("Your AI study assistant")


def get_client():
    key = None
    try:
        key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        key = os.environ.get("GEMINI_API_KEY")
    if not key:
        st.error("API key not found. Add GEMINI_API_KEY to your Streamlit secrets.")
        st.stop()
    return genai.Client(api_key=key)


client = get_client()

# ---------------------------------------------------------------- file reading
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
                out.append("... (more rows truncated)")
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
    """Turn attachments into Gemini parts. Returns (parts, names, image_bytes_list)."""
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
                # Gemini reads PDFs natively (works for scanned PDFs too)
                parts.append(types.Part.from_bytes(data=data, mime_type="application/pdf"))
            else:
                if ext == "docx":
                    text = read_docx(data)
                elif ext == "xlsx":
                    text = read_xlsx(data)
                elif ext == "pptx":
                    text = read_pptx(data)
                else:  # txt, csv, md, py ...
                    text = data.decode("utf-8", errors="replace")
                text = text[:MAX_CHARS]
                parts.append(types.Part.from_text(text=f"[Contents of file: {f.name}]\n{text}"))
            names.append(f.name)
        except Exception as e:
            st.warning(f"Could not read {f.name}: {e}")

    if camera_photo is not None:
        data = camera_photo.getvalue()
        parts.append(types.Part.from_bytes(data=data, mime_type="image/jpeg"))
        images.append(data)
        names.append("camera photo")

    return parts, names, images


# ---------------------------------------------------------------- PDF creation
def _latin(text: str) -> str:
    # built-in PDF fonts only support latin-1
    return text.encode("latin-1", "replace").decode("latin-1")


def text_to_pdf(text: str, title: str = "AssistU reply") -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.multi_cell(0, 10, _latin(title))
    pdf.ln(2)
    pdf.set_font("Helvetica", size=12)
    pdf.multi_cell(0, 7, _latin(text.replace("**", "").replace("`", "")))
    return bytes(pdf.output())


def image_to_pdf(img_bytes: bytes) -> bytes:
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    w, h = img.size
    max_w, max_h = 190, 277
    scale = min(max_w / w, max_h / h)
    pdf = FPDF()
    pdf.add_page()
    pdf.image(buf, x=10, y=10, w=w * scale, h=h * scale)
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

col_a, col_b = st.columns([3, 2])
with col_a:
    image_mode = st.toggle("🎨 Image generation mode",
                           help="When on, your message is used as an image prompt. "
                                "Attach a photo to edit it instead.")
with col_b:
    if st.button("🗑️ Clear chat", use_container_width=True):
        st.session_state.history = []
        st.session_state.uploader_key += 1
        st.rerun()

n_attached = len(uploaded or []) + (1 if photo is not None else 0)
if n_attached:
    st.info(f"📎 {n_attached} attachment(s) ready. They will be sent with your next message.")
if image_mode:
    st.caption("🎨 Image mode is ON: describe the picture you want.")

# ---------------------------------------------------------------- show history
for i, m in enumerate(st.session_state.history):
    with st.chat_message("user" if m["role"] == "user" else "assistant"):
        for img in m["images"]:
            st.image(img, width=300)
        if m["files"]:
            st.caption("📎 " + ", ".join(m["files"]))
        if m["text"]:
            st.markdown(m["text"])
        if m["role"] == "model":
            c1, c2 = st.columns(2)
            if m.get("generated"):
                c1.download_button("⬇️ PNG", m["images"][0], file_name=f"assistu_{i}.png",
                                   mime="image/png", key=f"png_{i}")
                c2.download_button("⬇️ PDF", image_to_pdf(m["images"][0]),
                                   file_name=f"assistu_{i}.pdf", mime="application/pdf",
                                   key=f"pdfi_{i}")
            elif m["text"]:
                c1.download_button("⬇️ PDF", text_to_pdf(m["text"]),
                                   file_name=f"assistu_reply_{i}.pdf",
                                   mime="application/pdf", key=f"pdft_{i}")

# ---------------------------------------------------------------- input
prompt = st.chat_input("Ask AssistU anything...")

if prompt:
    parts, names, images = build_parts(uploaded, photo)

    st.session_state.history.append({
        "role": "user", "text": prompt, "files": names, "images": images,
        "parts": parts + [types.Part.from_text(text=prompt)],
    })

    try:
        if image_mode:
            with st.spinner("Generating image..."):
                resp = client.models.generate_content(
                    model=IMAGE_MODEL,
                    contents=[prompt] + [p for p in parts if p.inline_data is not None
                                         and p.inline_data.mime_type.startswith("image/")],
                    config=types.GenerateContentConfig(response_modalities=["TEXT", "IMAGE"]),
                )
            out_text, out_img = "", None
            for part in resp.candidates[0].content.parts:
                if part.text:
                    out_text += part.text
                elif part.inline_data:
                    out_img = part.inline_data.data
            if out_img:
                st.session_state.history.append({
                    "role": "model", "text": out_text, "files": [], "images": [out_img],
                    "generated": True,
                    "parts": [types.Part.from_text(text=f"[Generated an image for: {prompt}]")],
                })
            else:
                reply = out_text or "I couldn't generate an image for that prompt."
                st.session_state.history.append({
                    "role": "model", "text": reply, "files": [], "images": [],
                    "parts": [types.Part.from_text(text=reply)],
                })
        else:
            contents = [types.Content(role=m["role"], parts=m["parts"])
                        for m in st.session_state.history]
            with st.spinner("Thinking..."):
                resp = client.models.generate_content(
                    model=CHAT_MODEL,
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
        st.session_state.history.pop()  # drop the failed user turn
        st.error(f"Something went wrong: {e}")
        st.stop()

    # reset the uploader/camera so attachments aren't re-sent next turn
    st.session_state.uploader_key += 1
    st.rerun()
