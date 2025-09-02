import os
import re
import io
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from dotenv import load_dotenv
import google.generativeai as genai
from PIL import Image
import streamlit.components.v1 as components

# --- SESSION STATE DEFAULTS ---
if "step" not in st.session_state:
    st.session_state["step"] = 1
if "subject" not in st.session_state:
    st.session_state["subject"] = ""
if "titles" not in st.session_state:
    st.session_state["titles"] = []
if "selected_title" not in st.session_state:
    st.session_state["selected_title"] = ""
if "slide_content" not in st.session_state:
    st.session_state["slide_content"] = []
if "images" not in st.session_state:
    st.session_state["images"] = []



# Load .env if present
load_dotenv()

# ---------- CONFIG ----------
DEFAULT_SLIDES = 5
AUDIENCE_PRESETS = ["Executive", "Technical", "Marketing", "Educational"]
MAX_WORDS_PER_SLIDE = 70
PEXELS_SEARCH_URL = "https://api.pexels.com/v1/search"

# ---------- INIT GEMINI ----------
G_API_KEY = os.getenv("G_API_KEY")
if not G_API_KEY:
    st.error("G_API_KEY not found. Please set environment variable G_API_KEY to use Gemini-Pro.")
else:
    genai.configure(api_key=G_API_KEY)
    MODEL = genai.GenerativeModel("gemini-2.5-flash")

# ---------- PEXELS (safe fetch) ----------
PEXELS_KEY = os.getenv("PEXELS_API_KEY")



def fetch_image_url_safe(query, api_key):
    """
    Return an image URL from Pexels for 'query', or None if not available.
    Handles 401/other codes gracefully (returns None).
    """
    if not api_key:
        return None
    headers = {"Authorization": api_key}
    params = {"query": query, "per_page": 1}
    try:
        resp = requests.get(PEXELS_SEARCH_URL, headers=headers, params=params, timeout=10)
    except Exception as e:
        # network error, timeout etc
        st.warning(f"Pexels request error for '{query}': {e}")
        return None

    if resp.status_code == 200:
        try:
            data = resp.json()
            photos = data.get("photos") or []
            if photos:
                # prefer landscape then original
                src = photos[0].get("src", {})
                return src.get("landscape") or src.get("original")
        except Exception as e:
            st.warning(f"Failed to parse Pexels response for '{query}': {e}")
            return None
    else:
        # handle unauthorized explicitly
        if resp.status_code == 401:
            st.warning("Pexels API returned 401 Unauthorized — check PEXELS_API_KEY. Proceeding without images.")
        else:
            st.warning(f"Pexels API returned status {resp.status_code} for '{query}' — proceeding without image.")
    return None

def download_image_to_path(img_url, keyword):
    """
    Download image bytes from img_url to a temp local path and return path, or None on failure.
    """
    try:
        r = requests.get(img_url, timeout=15)
        r.raise_for_status()
        safe_kw = re.sub(r'[^a-z0-9]', '_', keyword.lower())[:40]
        fname = f"/tmp/pexels_{safe_kw}.jpg"
        with open(fname, "wb") as f:
            f.write(r.content)
        return fname
    except Exception as e:
        st.warning(f"Failed to download image for '{keyword}': {e}")
        return None

# ---------- HELPERS ----------
def safe_filename(s: str) -> str:
    return re.sub(r'[^a-zA-Z0-9_\-]+', '_', s).strip('_')[:80]

def word_count(text: str) -> int:
    return len(text.split())

def parse_lines_to_bullets_and_notes(generated_text: str):
    """
    Parse LLM output into bullets and notes. Fallback to first lines as bullets if no list present.
    """
    lines = [ln.strip() for ln in generated_text.splitlines() if ln.strip()]
    bullets = []
    notes_lines = []
    for ln in lines:
        if re.match(r"^(-|•|\d+\.)\s*", ln):
            bullets.append(re.sub(r"^(-|•|\d+\.)\s*", "", ln).strip())
        else:
            notes_lines.append(ln)
    if not bullets and lines:
        bullets = lines[:4]
        notes_lines = lines[4:]
    notes = " ".join(notes_lines).strip()
    return bullets[:6], notes

# ---------- LLM (Gemini) ----------
def generate_titles(subject, count=6):
    system = "You are an expert presentation author. Produce short, engaging presentation titles (one per line)."
    prompt = f"Generate {count} concise and compelling presentation titles for the subject: \"{subject}\". Each title under 10 words."
    try:
        resp = MODEL.generate_content([system, prompt])
        text = resp.text or ""
        titles = [re.sub(r'^[\-\d\.\)\s]+', '', line).strip() for line in text.splitlines() if line.strip()]
        return titles[:count] if titles else []
    except Exception as e:
        st.error(f"Error generating titles: {e}")
        return []

def generate_slide_text(ppt_title, section_title, audience, include_image_keyword=True):
    system = "You are an expert presentation writer. Use concise bullets and short speaker notes."
    prompt = (
        f"Create slide content for presentation titled '{ppt_title}'.\n"
        f"Slide title: {section_title}\n"
        f"Audience: {audience}\n"
        "- Provide 3 to 5 concise bullet points (<= 20 words each).\n"
        "- Provide a 1-2 sentence speaker note.\n"
    )
    if include_image_keyword:
        prompt += "- At the end, provide one short image keyword (2-4 words) prefixed with 'ImageKeyword:'.\n"
    try:
        resp = MODEL.generate_content([system, prompt])
        text = resp.text or ""
        image_keyword = None
        ik_match = re.search(r"ImageKeyword\s*:\s*(.+)$", text, flags=re.IGNORECASE | re.M)
        if ik_match:
            image_keyword = ik_match.group(1).strip()
            text = re.sub(r"ImageKeyword\s*:\s*.+$", "", text, flags=re.IGNORECASE | re.M).strip()
        bullets, notes = parse_lines_to_bullets_and_notes(text)
        bullets = [b.strip() for b in bullets if b.strip()]
        return {"bullets": bullets, "notes": notes, "image_keyword": image_keyword}
    except Exception as e:
        st.error(f"Error generating slide content: {e}")
        return {"bullets": ["(Unable to generate slide)"], "notes": "", "image_keyword": None}

# ---------- PPT CREATION ----------
def create_pptx_bytes(ppt_title, slide_contents, attach_images=False):
    prs = Presentation()
    # Title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = ppt_title
    try:
        slide.placeholders[1].text = "Generated by AI PPT Wizard"
    except Exception:
        pass

    for s in slide_contents:
        layout = prs.slide_layouts[1]
        slide_obj = prs.slides.add_slide(layout)
        slide_obj.shapes.title.text = s.get("slide_title", "")[:80]
        body = slide_obj.shapes.placeholders[1].text_frame
        body.clear()
        for b in s.get("bullets", []):
            p = body.add_paragraph()
            p.text = b
            p.level = 0
        try:
            slide_obj.notes_slide.notes_text_frame.text = s.get("notes","")
        except Exception:
            pass

        img_path = s.get("image_local_path")
        if attach_images and img_path and os.path.exists(img_path):
            try:
                left = prs.slide_width * 0.5
                top = Inches(1.0)
                width = prs.slide_width * 0.45
                slide_obj.shapes.add_picture(img_path, left, top, width=width)
            except Exception as e:
                st.warning(f"Could not add image to slide '{s.get('slide_title')}': {e}")

    closing = prs.slides.add_slide(prs.slide_layouts[1])
    closing.shapes.title.text = "Conclusion & Next Steps"
    try:
        closing.shapes.placeholders[1].text_frame.text = "Summary and suggested next steps."
    except Exception:
        pass

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# ---------- STREAMLIT UI ----------
st.set_page_config(page_title="AI PPT Wizard (Gemini-Pro)", layout="wide")
st.markdown("""
    <style>
    .stApp { background: linear-gradient(rgba(6,8,12,0.92), rgba(6,8,12,0.92)); color: #e6eef8 !important; }
    .stButton>button { background-color:#2563EB !important; color: white !important; border: none !important; }
    .stSidebar { background-color:#071233 !important; color: #e6eef8 !important; }
    .stTextInput input, textarea { background-color:#0c1320 !important; color: #e6eef8 !important; border:1px solid #233554 !important; }
    </style>
""", unsafe_allow_html=True)

# Sidebar content (instructions + settings)
with st.sidebar:
    st.title("AI PPT Wizard")
    if st.button("See Instructions"):
        st.info(
            "Flow:\n"
            "1) Enter Topic & choose audience\n"
            "2) Generate titles and pick one\n"
            "3) Edit outline/sections\n"
            "4) Preview & edit slide bullets and notes\n"
            "5) Generate PPT and download\n\n"
            "Notes:\n- Use Gemini-Pro (G_API_KEY) for best results.\n- Optional: set PEXELS_API_KEY to auto-insert images."
        )
    st.markdown("---")
    st.subheader("Settings")
    attach_images = st.checkbox("Fetch & attach images via Pexels (optional)", value=False,
                                help="Requires PEXELS_API_KEY in env. If unchecked, only image keywords are shown.")
    st.markdown("---")
    st.subheader("Terms & Disclaimer")
    st.markdown(
        "• The content is AI-generated for demo/informational use only.\n\n"
        "• Confirm facts and consult professionals for critical decisions.\n\n"
        "• Gemini API calls use your API key; manage accordingly."
    )

# One-time popup
if 'popup_shown' not in st.session_state:
    st.session_state['popup_shown'] = True
    components.html("<script>alert('💡 This app uses a dark UI for best visibility.');</script>")

st.title("AI PPT Wizard — Gemini-Pro")
st.write("Guided flow: Topic → Titles → Outline → Edit → Generate PPT")

# Step state

def go_next(): st.session_state['step'] = min(5, st.session_state['step'] + 1)
def go_back(): st.session_state['step'] = max(1, st.session_state['step'] - 1)

# Step 1
if st.session_state['step'] == 1:
    st.header("Step 1 — Topic & Purpose")
    topic = st.text_input("Enter the presentation topic", value=st.session_state.get('topic',''))
    audience = st.selectbox("Audience style", AUDIENCE_PRESETS, index=0)
    slides_count = st.slider("Desired number of slides", min_value=1, max_value=30, value=DEFAULT_SLIDES)
    if st.button("Generate Titles"):
        if not topic:
            st.error("Please enter a topic.")
        else:
            st.session_state['topic'] = topic
            st.session_state['audience'] = audience
            st.session_state['slides_count'] = slides_count
            with st.spinner("Generating titles via Gemini…"):
                titles = generate_titles(topic, count=8)
            if not titles:
                st.error("No titles generated. Try a different topic or check API key.")
            else:
                st.session_state['titles'] = titles
                go_next()
    st.markdown("---")
    st.write("Tip: Keep topics concise and descriptive (e.g., 'AI in Healthcare').")

# Step 2
if st.session_state['step'] == 2:
    st.header("Step 2 — Pick & Edit Title")
    titles = st.session_state.get('titles', [])
    if not titles:
        st.warning("No titles yet — go back and generate.")
        if st.button("Back"): go_back()
    else:
        selected = st.radio("Choose a title", options=titles, index=0)
        custom_title = st.text_input("Or edit the chosen title", value=selected)
        cols = st.columns([1,1,1])
        with cols[0]:
            if st.button("Back"): go_back()
        with cols[1]:
            if st.button("Regenerate Titles"):
                with st.spinner("Regenerating..."):
                    st.session_state['titles'] = generate_titles(st.session_state.get('topic',''), count=8)
                st.experimental_rerun()
        with cols[2]:
            if st.button("Proceed"):
                st.session_state['final_title'] = custom_title or selected
                topic = st.session_state.get('topic','Topic')
                default_sections = [
                    f"Introduction to {topic}",
                    f"Importance of {topic}",
                    f"Key technologies in {topic}",
                    "Case studies",
                    "Conclusion & Recommendations"
                ]
                st.session_state['sections'] = default_sections[:st.session_state.get('slides_count', DEFAULT_SLIDES)]
                go_next()

# Step 3
if st.session_state['step'] == 3:
    st.header("Step 3 — Outline (Edit Sections)")
    sections = st.session_state.get('sections', [])
    if not sections:
        st.warning("No sections configured. Go back to select a title.")
        if st.button("Back"): go_back()
    else:
        edited = []
        for i, s in enumerate(sections):
            new_s = st.text_input(f"Section {i+1}", value=s, key=f"sec_{i}")
            edited.append(new_s)
        extra = st.text_input("Add an extra section (optional)")
        if extra:
            edited.append(extra)
        st.session_state['edited_sections'] = edited
        cols = st.columns([1,1])
        with cols[0]:
            if st.button("Back"): go_back()
        with cols[1]:
            if st.button("Generate Slide Content"):
                final_sections = edited[:st.session_state.get('slides_count', DEFAULT_SLIDES)]
                slide_contents = []
                with st.spinner("Generating slide content via Gemini…"):
                    for sec in final_sections:
                        slide = generate_slide_text(st.session_state.get('final_title','Presentation'), sec, st.session_state.get('audience','Executive'), include_image_keyword=True)
                        slide['slide_title'] = sec
                        slide['image_local_path'] = None
                        # safe pexels flow
                        if attach_images and slide.get('image_keyword'):
                            img_url = fetch_image_url_safe(slide['image_keyword'], PEXELS_KEY)
                            if img_url:
                                local_path = download_image_to_path(img_url, slide['image_keyword'])
                                if local_path:
                                    slide['image_local_path'] = local_path
                        slide_contents.append(slide)
                st.session_state['slide_contents'] = slide_contents
                go_next()

# Step 4
if st.session_state['step'] == 4:
    st.header("Step 4 — Preview & Edit Slides")
    slide_contents = st.session_state.get('slide_contents', [])
    if not slide_contents:
        st.warning("No slide content found. Generate slides first.")
        if st.button("Back"): go_back()
    else:
        for idx, s in enumerate(slide_contents):
            st.markdown(f"### Slide {idx+1}")
            title_in = st.text_input(f"Slide {idx+1} Title", value=s.get('slide_title',''), key=f"title_{idx}")
            bullets = s.get('bullets', [])
            new_bullets = []
            st.write("Bullets:")
            for j, b in enumerate(bullets):
                nb = st.text_input(f"Slide {idx+1} - Bullet {j+1}", value=b, key=f"bullet_{idx}_{j}")
                new_bullets.append(nb)
            if st.button(f"Add bullet to slide {idx+1}", key=f"add_b_{idx}"):
                new_bullets.append("(New bullet)")
                st.experimental_rerun()
            notes = st.text_area(f"Speaker notes for slide {idx+1}", value=s.get('notes',''), key=f"notes_{idx}")
            image_kw = s.get('image_keyword')
            st.write(f"Image suggestion: **{image_kw if image_kw else 'No suggestion'}**")
            if s.get('image_local_path') and os.path.exists(s.get('image_local_path')):
                st.image(s.get('image_local_path'), width=320)
            if word_count(" ".join(new_bullets)) > MAX_WORDS_PER_SLIDE:
                st.warning(f"Slide {idx+1} exceeds recommended {MAX_WORDS_PER_SLIDE} words.")
            st.session_state['slide_contents'][idx]['slide_title'] = title_in
            st.session_state['slide_contents'][idx]['bullets'] = new_bullets
            st.session_state['slide_contents'][idx]['notes'] = notes

        cols = st.columns([1,1,1])
        with cols[0]:
            if st.button("Back"): go_back()
        with cols[1]:
            if st.button("Regenerate this slide (not implemented)"):
                st.info("Partial regeneration placeholder. You can edit bullets manually or regenerate whole flow.")
        with cols[2]:
            if st.button("Finalize & Create PPT"):
                ppt_title = st.session_state.get('final_title','AI_Presentation')
                bio = create_pptx_bytes(ppt_title, st.session_state['slide_contents'], attach_images=attach_images)
                filename = safe_filename(ppt_title) + ".pptx"
                st.success("PPT created!")
                st.download_button("Download PPT", data=bio, file_name=filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                st.session_state['step'] = 5

# Step 5
if st.session_state['step'] == 5:
    st.header("Step 5 — Finished")
    st.success("Presentation generated — check your downloads.")
    st.write("You can go back and tweak slides or start a new presentation.")
    if st.button("Start New Presentation"):
        keys = ['topic','audience','slides_count','titles','final_title','sections','edited_sections','slide_contents']
        for k in keys:
            if k in st.session_state: del st.session_state[k]
        st.session_state['step'] = 1
        st.experimental_rerun()
