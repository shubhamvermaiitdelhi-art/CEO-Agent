import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches
import io
import os
from pathlib import Path
from datetime import datetime

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Executive Strategy Agent", page_icon="🏛️", layout="centered")

# Custom UI Styling
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; background-color: #1A73E8; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏛️ CEO Strategy Agent")
st.caption("Developed for Shubham Verma | 2026 Leadership Edition")

# --- 2. KEYS & CONFIG ---
try:
    PPLX_KEY = st.secrets["PPLX_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
except Exception:
    st.error("🔑 API Keys missing! Add PPLX_KEY and GEMINI_KEY to Streamlit Secrets.")
    st.stop()

# Clients
pplx_client = OpenAI(api_key=PPLX_KEY, base_url="https://api.perplexity.ai")
genai.configure(api_key=GEMINI_KEY)

# --- 3. CORE ENGINE ---

def get_research(company):
    """The Hunter: Uses Perplexity Sonar Pro for live 2026 data."""
    query = f"Find the top 3 massive strategic/financial failures for {company} in 2025-2026. Give specific $ numbers and technical bottlenecks."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": query}]
    )
    return response.choices[0].message.content

def get_slide_script(company, research):
    """The Architect: Writes a 15-slide script with image prompts."""
    model = genai.GenerativeModel("gemini-1.5-pro")
    prompt = f"""
    Research: {research}
    Create a 15-slide pitch for a Director role at {company} by Shubham Verma.
    Format your response STRICTLY as a Python list of dictionaries:
    [
      {{"title": "Title", "bullets": ["p1", "p2"], "img_prompt": "Professional 16:9 4k business photo of..."}},
      ...
    ]
    Include 15 slides.
    """
    response = model.generate_content(prompt)
    content = response.text.replace("```python", "").replace("```", "").strip()
    return eval(content)

def generate_image(prompt):
    """The Artist: Generates 16:9 business images using Imagen 3."""
    # Note: Imagen 3 API via Gemini 1.5/3 Pro
    model = genai.GenerativeModel("gemini-1.5-pro")
    # This triggers the internal image generation tool
    response = model.generate_content(f"Generate a professional 16:9 business image: {prompt}")
    # In a real API environment, we'd extract the image bytes
    # For this script, we'll return a placeholder logic or use a stable API path
    return response

# --- 4. THE INTERFACE ---

company_name = st.text_input("🏢 Enter Target Company Name:")

if company_name:
    if st.button("🚀 GENERATE LEADERSHIP PROPOSAL"):
        with st.status("Agent Executing...", expanded=True) as status:
            
            data = get_research(company_name)
            script = get_slide_script(company_name, data)
            
            # 1. LOAD & CLEAR TEMPLATE
            base_path = Path(__file__).parent
            tpl_path = base_path / "master_template.pptx"
            prs = Presentation(tpl_path) if tpl_path.exists() else Presentation()
            
            # DELETE ALL EXISTING SLIDES (Clean Start)
            xml_slides = prs.slides._sldIdLst
            for i in range(len(xml_slides)):
                del xml_slides[0]

            # 2. ADD NEW CONTENT
            current_date = datetime.now().strftime("%B %d, %2026")
            
            for i, slide_info in enumerate(script):
                layout = prs.slide_layouts[0] if i == 0 else prs.slide_layouts[1]
                slide = prs.slides.add_slide(layout)
                
                # Title
                slide.shapes.title.text = slide_info['title']
                
                # Main Text
                if len(slide.placeholders) > 1:
                    tf = slide.placeholders[1].text_frame
                    tf.text = ""
                    for point in slide_info['bullets']:
                        p = tf.add_paragraph()
                        p.text = point
                
                # 3. INSERT SHUBHAM'S CREDENTIALS (Slide 1)
                if i == 0:
                    subtitle = slide.placeholders[1]
                    subtitle.text = f"Presented by: Shubham Verma\nStrategy Director Candidate\nDate: {current_date}"

            # 4. SAVE TO MEMORY
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
            status.update(label="✅ Deck Complete!", state="complete")

        st.download_button(
            label="📥 Download Strategy Presentation",
            data=ppt_io,
            file_name=f"{company_name}_Strategy_Verma.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
