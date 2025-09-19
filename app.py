# app.py
import os, re, tempfile, fitz, docx, requests, copy
import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR

# ---------------- CONFIG ----------------
GEMINI_API_KEY = "AIzaSyBtah4ZmuiVkSrJABE8wIjiEgunGXAbT3Q"  # 🔑 Hardcoded Gemini API key
TEXT_MODEL_NAME = "gemini-2.0-flash"

# ---------------- HELPERS ----------------
def call_gemini(prompt: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{TEXT_MODEL_NAME}:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        resp = requests.post(url, json=payload, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception as e:
        return f"⚠️ Gemini API error: {e}"

def parse_points(points_text: str):
    points, current_title, current_content = [], None, []
    lines = [re.sub(r"[#*>`]", "", ln).rstrip() for ln in points_text.splitlines()]
    for line in lines:
        if not line or "Would you like" in line: continue
        m = re.match(r"^\s*(Slide|Section)\s*(\d+)\s*:\s*(.+)$", line, re.IGNORECASE)
        if m:
            if current_title:
                points.append({"title": current_title, "description": "\n".join(current_content)})
            current_title, current_content = m.group(3).strip(), []
            continue
        if line.strip().startswith("-"):
            text = line.lstrip("-").strip()
            if text: current_content.append(f"• {text}")
        elif line.strip().startswith(("•", "*")) or line.startswith("  "):
            text = line.lstrip("•*").strip()
            if text: current_content.append(f"- {text}")
        else:
            if line.strip(): current_content.append(line.strip())
    if current_title:
        points.append({"title": current_title, "description": "\n".join(current_content)})
    return points

def generate_outline(description: str):
    prompt = f"""Create a PowerPoint outline on: {description}.
Each slide should have a short title and 3–4 bullet points.
Format strictly like this:
Slide 1: <Title>
- Bullet
- Bullet
- Bullet
"""
    outline_text = call_gemini(prompt)
    return parse_points(outline_text)

def extract_text(path: str, filename: str) -> str:
    name = filename.lower()
    if name.endswith(".pdf"):
        text_parts = []
        doc = fitz.open(path)
        try:
            for page in doc: text_parts.append(page.get_text("text"))
        finally: doc.close()
        return "\n".join(text_parts)
    if name.endswith(".docx"):
        d = docx.Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    if name.endswith(".txt"):
        for enc in ("utf-8","utf-16","utf-16-le","utf-16-be","latin-1"):
            try:
                with open(path,"r",encoding=enc) as f: return f.read()
            except UnicodeDecodeError: continue
        with open(path,"r",encoding="utf-8",errors="ignore") as f: return f.read()
    return ""

def split_text(text: str, chunk_size: int = 8000, overlap: int = 300):
    if not text: return []
    chunks, start, n = [], 0, len(text)
    while start < n:
        end = min(start + chunk_size, n)
        chunks.append(text[start:end])
        if end == n: break
        start = max(0, end - overlap)
    return chunks

def summarize_long_text(full_text: str) -> str:
    chunks = split_text(full_text)
    if len(chunks) <= 1:
        return call_gemini(f"Summarize the following text in detail:\n\n{full_text}")
    partial_summaries = []
    for idx,ch in enumerate(chunks, start=1):
        mapped = call_gemini(f"Summarize this part of a longer document:\n\n{ch}")
        partial_summaries.append(f"Chunk {idx}:\n{mapped.strip()}")
    combined = "\n\n".join(partial_summaries)
    return call_gemini(f"Combine these summaries into one clean, well-structured summary:\n\n{combined}")

def sanitize_filename(name: str) -> str:
    return re.sub(r'[^A-Za-z0-9_.-]', '_', name)

def clean_title_text(title: str) -> str:
    if not title: return "Presentation"
    return re.sub(r"\s+", " ", title.strip())

def hex_to_rgb(hex_color: str):
    hex_color = hex_color.lstrip("#")
    return RGBColor(int(hex_color[0:2],16), int(hex_color[2:4],16), int(hex_color[4:6],16))

def create_ppt(title, points, filename="output.pptx", title_size=30, text_size=22,
               font="Calibri", title_color="#5E2A84", text_color="#282828", background_color="#FFFFFF"):
    prs = Presentation()
    title = clean_title_text(title)

    # Title Slide
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = hex_to_rgb(background_color)
    textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
    tf = textbox.text_frame; tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, MSO_VERTICAL_ANCHOR.MIDDLE
    p = tf.add_paragraph()
    p.text, p.font.size, p.font.bold, p.font.name, p.font.color.rgb, p.alignment = title, Pt(title_size), True, font, hex_to_rgb(title_color), PP_ALIGN.CENTER

    # Content Slides
    for idx, item in enumerate(points, start=1):
        key_point, description = clean_title_text(item.get("title","")), item.get("description","")
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = hex_to_rgb(background_color)
        textbox = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8), Inches(1.5))
        tf = textbox.text_frame; tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, MSO_VERTICAL_ANCHOR.MIDDLE
        p = tf.add_paragraph()
        p.text, p.font.size, p.font.bold, p.font.name, p.font.color.rgb, p.alignment = key_point, Pt(title_size), True, font, hex_to_rgb(title_color), PP_ALIGN.LEFT
        if description:
            textbox = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(5), Inches(4))
            tf = textbox.text_frame; tf.word_wrap = True
            for line in description.split("\n"):
                if line.strip():
                    bullet = tf.add_paragraph()
                    bullet.text, bullet.font.size, bullet.font.name, bullet.font.color.rgb, bullet.level = line.strip(), Pt(text_size), font, hex_to_rgb(text_color), 0
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(8), Inches(0.3))
        tf = textbox.text_frame; p = tf.add_paragraph()
        p.text, p.font.size, p.font.name, p.font.color.rgb, p.alignment = "Generated with AI", Pt(10), font, RGBColor(150,150,150), PP_ALIGN.RIGHT

    prs.save(filename); return filename

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="AI Productivity Suite", layout="wide")
st.title("AI Productivity Suite")

defaults = {"messages": [], "outline_chat": None, "summary_text": None, "summary_title": None, "doc_chat_history": []}
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k]=v

# Display chat history
for role, content in st.session_state.messages: 
    with st.chat_message(role): st.markdown(content)
for role, content in st.session_state.doc_chat_history:
    with st.chat_message(role): st.markdown(content)

# ---------------- FILE UPLOAD ----------------
uploaded_file = st.file_uploader("📂 Upload a document", type=["pdf","docx","txt"])
if uploaded_file:
    with st.spinner("Processing uploaded file..."):
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(uploaded_file.getvalue()); tmp_path=tmp.name
        text = extract_text(tmp_path, uploaded_file.name); os.remove(tmp_path)
        if text.strip():
            summary = summarize_long_text(text)
            title = call_gemini(f"Generate a short title for this summary:\n{summary}")
            st.session_state.summary_text, st.session_state.summary_title = summary, title
            st.success(f"✅ Uploaded! Suggested Title: **{title}**")
        else:
            st.error("❌ Unsupported, empty, or unreadable file.")

# ---------------- CHAT ----------------
if prompt := st.chat_input("💬 Type a message..."):
    if st.session_state.summary_text:
        if any(w in prompt.lower() for w in ["ppt","slides","presentation"]):
            slides = generate_outline(st.session_state.summary_text + "\n\n" + prompt)
            st.session_state.outline_chat = {"title": st.session_state.summary_title, "slides": slides}
            st.session_state.doc_chat_history.append(("assistant","✅ Generated PPT outline from document."))
        else:
            st.session_state.doc_chat_history.append(("user",prompt))
            reply = call_gemini(f"Answer using only this doc:\n{st.session_state.summary_text}\n\nQ:{prompt}")
            st.session_state.doc_chat_history.append(("assistant",reply))
    else:
        st.session_state.messages.append(("user",prompt))
        if "ppt" in prompt.lower():
            slides = generate_outline(prompt)
            st.session_state.outline_chat = {"title":"Generated PPT","slides":slides}
            st.session_state.messages.append(("assistant","✅ PPT outline generated!"))
        else:
            reply = call_gemini(prompt)
            st.session_state.messages.append(("assistant",reply))
    st.rerun()

# ---------------- OUTLINE PREVIEW ----------------
if st.session_state.outline_chat:
    outline = st.session_state.outline_chat
    st.subheader(f"📝 Preview Outline: {outline['title']}")
    for idx,slide in enumerate(outline["slides"],start=1):
        with st.expander(f"Slide {idx}: {slide['title']}",expanded=False):
            st.markdown(slide["description"].replace("\n","\n\n"))
    if st.button("✅ Generate PPT"):
        with st.spinner("Generating PPT..."):
            filename = f"{sanitize_filename(outline['title'])}.pptx"
            create_ppt(outline['title'], outline["slides"], filename)
            with open(filename,"rb") as f:
                st.download_button("⬇️ Download PPT", data=f, file_name=filename,
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
