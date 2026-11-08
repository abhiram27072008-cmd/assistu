import hashlib
import io
import json
import os
import uuid

import streamlit as st
from docx import Document
from fpdf import FPDF
from google import genai
from google.genai import types
from openpyxl import load_workbook
from pptx import Presentation
from streamlit_js_eval import streamlit_js_eval

# ---------------------------------------------------------------- settings
CHAT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-2.5-flash"  # used if the main model is busy
MAX_CHARS = 100_000  # max text read from one file
STORAGE_KEY = "assistu_chats_v1"  # name of the saved data in the browser
MAX_SAVED_BYTES = 4_000_000  # browser storage is small, keep under this

SYSTEM_PROMPT = ("You are AssistU, a friendly assistant for students. "
                 "Explain simply with examples. "
                 "When files or photos are attached, read them carefully "
                 "and answer based on their content.")

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


def stream_reply(contents):
    """Yield the answer piece by piece. If the main model is busy (503),
    switch to the backup model straight away."""
    config = types.GenerateContentConfig(system_instruction=SYSTEM_PROMPT)
    last_err = None
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
            last_err = e
            busy = "503" in str(e) or "UNAVAILABLE" in str(e)
            if started or not busy:
                raise
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
@st.cache_data(show_spinner=False)
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

if st.button("🗑️ Clear this chat"):
    chat["history"] = []
    chat["title"] = "New chat"
    st.session_state.uploader_key += 1
    st.rerun()

n_attached = len(uploaded or []) + (1 if photo is not None else 0)
if n_attached:
    st.info(f"📎 {n_attached} attachment(s) ready. They will be sent with your next message.")

# ---------------------------------------------------------------- show chat
for i, m in enumerate(history):
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

# ---------------------------------------------------------------- save to the browser
# Runs on every normal page run. It writes only when the chats have changed.
if st.session_state.save_ok:
    payload = build_payload()
    save_js = f"localStorage.setItem('{STORAGE_KEY}', {json.dumps(payload)})"
    save_key = "ls_save_" + hashlib.md5(payload.encode("utf-8")).hexdigest()[:12]
else:
    save_js, save_key = "0", "ls_save_off"
streamlit_js_eval(js_expressions=save_js, key=save_key)
