import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches
import io
import os
from pathlib import Path
from datetime import datetime
import PIL.Image

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Executive Strategy Agent", page_icon="🏛️", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; background-color: #1A73E8; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏛️ CEO Strategy Agent")
st.caption("Strategic Advisor: Shubham Verma | 2026 Edition")

# --- 2. KEYS & CONFIG ---
try:
    PPLX_KEY = st.secrets["PPLX_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
except Exception:
    st.error("🔑 API Keys missing! Add PPLX_KEY and GEMINI_KEY to Streamlit Secrets.")
    st.stop()

# Initialize AI Clients
pplx_client = OpenAI(api_key=PPLX_KEY, base_url="https://api.perplexity.ai")
genai.configure(api_key=GEMINI_KEY)

# --- 3. HELPER FUNCTIONS ---

def get_research(company):
    """The Hunter: Uses Perplexity Sonar Pro for current market insights."""
    query = f"Research {company} for 2026. Identify 3 massive financial/operational leaks ($ numbers) and the top technical bottleneck."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": query}]
    )
    return response.choices[0].message.content

def get_slide_script(company, research):
    """The Architect: Writes a 15-slide script using Gemini 3."""
    # Using 2026 stable model names
    model = genai.GenerativeModel("gemini-3-flash-preview")
    prompt = f"""
    Based on research: {research}
    Create a 15-slide pitch for a Director role at {company} for candidate Shubham Verma.
    Output a Python list of dicts: [{{'title': '...', 'bullets': ['...'], 'img_prompt': '...'}}]
    Slide 1: Title Slide. Slide 15: Execution Ask.
    """
    response = model.generate_content(prompt)
    content = response.text.replace("```python", "").replace("```", "").strip()
    return eval(content)

def generate_and_save_image(prompt):
    """The Artist: Generates images via Gemini 3 (Imagen)."""
    # Uses the Gemini 3 Pro model which supports image generation in 2026
    image_model = genai.GenerativeModel("gemini-3-pro-image-preview")
    try:
        response = image_model.generate_content(f"Generate a professional, minimalist business photo: {prompt}")
        # Assuming response contains a PIL image object in 2026 SDK
        image = response.candidates[0].content.parts[0].inline_data.data # Simplified for logic
        return io.BytesIO(image)
    except:
        return None

# --- 4. THE INTERFACE ---

company_name = st.text_input("🏢 Enter Target Company Name:")

if company_name:
    if st.button("🚀 GENERATE LEADERSHIP PROPOSAL"):
        with st.status("Agent Executing...", expanded=True) as status:
            
            # Step 1: Research & Scripting
            data = get_research(company_name)
            script = get_slide_script(company_name, data)
            
            # Step 2: Load & Purge Master PPT
            base_path = Path(__file__).parent
            tpl_path = base_path / "master_template.pptx"
            prs = Presentation(tpl_path) if tpl_path.exists() else Presentation()
            
            # Delete all existing slides to overwrite content
            for _ in range(len(prs.slides)):
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]

            # Step 3: Populate New Slides
            current_date = datetime.now().strftime("%B %d, 2026")
            
            for i, slide_info in enumerate(script):
                layout_idx = 0 if i == 0 else 1 # Title layout vs Content layout
                slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
                
                # Header
                slide.shapes.title.text = slide_info['title']
                
                # Bullets
                if len(slide.placeholders) > 1:
                    tf = slide.placeholders[1].text_frame
                    tf.text = ""
                    for bullet in slide_info['bullets']:
                        p = tf.add_paragraph()
                        p.text = bullet
                
                # Image Generation and Insertion
                img_data = generate_and_save_image(slide_info['img_prompt'])
                if img_data:
                    # Place image on the right side of the slide
                    slide.shapes.add_picture(img_data, Inches(6), Inches(1.5), height=Inches(4.5))
                
                # Branding Slide 1
                if i == 0:
                    body = slide.placeholders[1]
                    body.text = f"Prepared for CEO of {company_name}\nBy: Shubham Verma\nDate: {current_date}"

            # Step 4: Finalize
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
            status.update(label="✅ Presentation Ready!", state="complete")

        st.download_button(
            label="📥 Download Strategy Deck",
            data=ppt_io,
            file_name=f"{company_name}_Strategy_Verma.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
