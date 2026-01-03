import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches
import io
from pathlib import Path
from datetime import datetime

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Executive Strategy Agent", page_icon="🏛️", layout="centered")

# CSS to give the app a BCG/McKinsey "clean" feel
st.markdown("""
    <style>
    .main { background-color: #ffffff; color: #00453c; }
    .stButton>button { background-color: #00453c; color: white; border-radius: 4px; border: none; font-weight: bold; }
    .stTextInput>div>div>input { border-radius: 4px; }
    </style>
    """, unsafe_allow_html=True)

st.title("Strategic Advisor Agent")
st.write("Candidate: **Shubham Verma** | Strategy & AI Leadership")

# --- 2. KEYS & CONFIG ---
try:
    PPLX_KEY = st.secrets["PPLX_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
except:
    st.error("🔑 Keys missing in Streamlit Secrets!")
    st.stop()

# Clients
pplx_client = OpenAI(api_key=PPLX_KEY, base_url="https://api.perplexity.ai")
genai.configure(api_key=GEMINI_KEY)

# --- 3. HELPER FUNCTIONS ---

def get_research(company):
    """Hunter: Real-time BCG-style problem identification."""
    query = f"Provide a SWOT analysis for {company} in 2026 focusing on financial leaks and technical debt. Identify one $100M+ problem."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": query}]
    )
    return response.choices[0].message.content

def get_slide_script(company, research):
    """Architect: Writes the 15-slide script."""
    # Using the most robust model for logic
    model = genai.GenerativeModel("gemini-1.5-pro")
    prompt = f"""
    You are Shubham Verma, a Strategy Director candidate. 
    Research: {research}
    Create a 15-slide pitch deck for {company}. 
    Format: Python list of dicts: [{{'title': '...', 'bullets': ['...'], 'img_prompt': '...'}}]
    Slides: 1:Title, 2:Executive Summary, 3-5:Situation, 6-10:AI Solution, 11-14:ROI, 15:The Ask.
    """
    response = model.generate_content(prompt)
    return eval(response.text.replace("```python", "").replace("```", "").strip())

def get_best_layout(prs, keywords):
    """Finds BCG master layouts by keyword search."""
    for layout in prs.slide_layouts:
        for word in keywords:
            if word.lower() in layout.name.lower():
                return layout
    return prs.slide_layouts[1] # Default fallback

# --- 4. EXECUTION ---

target_company = st.text_input("🏢 Enter Company Name:")

if target_company:
    if st.button("🚀 GENERATE BCG-STYLE PROPOSAL"):
        with st.status("Agent Executing...", expanded=True) as status:
            
            # Data & Script
            data = get_research(target_company)
            script = get_slide_script(target_company, data)
            
            # Load Template
            tpl_path = Path(__file__).parent / "master_template.pptx"
            prs = Presentation(tpl_path) if tpl_path.exists() else Presentation()
            
            # 1. DELETE ALL EXISTING SLIDES (Clean Start for your BCG Template)
            for _ in range(len(prs.slides)):
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]

            # 2. POPULATE
            today = datetime.now().strftime("%B %d, 2026")
            
            for i, s_data in enumerate(script):
                # Choose layout based on BCG naming conventions
                if i == 0:
                    layout = get_best_layout(prs, ["Title", "Cover", "Intro"])
                else:
                    layout = get_best_layout(prs, ["Content", "Bullet", "Body"])
                
                slide = prs.slides.add_slide(layout)
                
                # Set Title
                if slide.shapes.title:
                    slide.shapes.title.text = s_data['title']
                
                # Find the main content placeholder (not the title)
                body_ph = None
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx != 0:
                        body_ph = ph
                        break
                
                if body_ph and body_ph.has_text_frame:
                    tf = body_ph.text_frame
                    tf.text = ""
                    for bullet in s_data['bullets']:
                        p = tf.add_paragraph()
                        p.text = bullet
                
                # Personal Branding Slide 1
                if i == 0 and body_ph:
                    body_ph.text = f"Presented to the CEO of {target_company}\nBy: Shubham Verma\nStrategy Director Candidate\nDate: {today}"

            # 3. SAVE
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
            status.update(label="✅ Pitch Deck Complete!", state="complete")

        st.download_button(
            label="📥 Download Strategy Deck",
            data=ppt_io,
            file_name=f"{target_company}_Pitch_Verma.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
