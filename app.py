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

st.markdown("""
    <style>
    .main { background-color: #ffffff; color: #00453c; }
    .stButton>button { background-color: #00453c; color: white; border-radius: 4px; font-weight: bold; width: 100%; height: 3em; }
    </style>
    """, unsafe_allow_html=True)

st.title("Strategic Advisor Agent")
st.write("Candidate: **Shubham Verma** | Strategy & AI Leadership")

# --- 2. KEYS & CONFIG ---
try:
    PPLX_KEY = st.secrets["PPLX_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
except:
    st.error("🔑 Keys missing in Streamlit Secrets! Please add PPLX_KEY and GEMINI_KEY.")
    st.stop()

# Clients
pplx_client = OpenAI(api_key=PPLX_KEY, base_url="https://api.perplexity.ai")
genai.configure(api_key=GEMINI_KEY)

# --- 3. SMART MODEL RESOLVER (FIXES 404 ERROR) ---
def get_best_available_model():
    """Finds the best model your API key is allowed to use."""
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        # Priority List for 2026
        priorities = ['models/gemini-1.5-pro', 'models/gemini-1.5-flash', 'models/gemini-pro']
        
        for p in priorities:
            if p in available_models:
                return genai.GenerativeModel(p)
        
        # If none of the above are found, just pick the first one available
        return genai.GenerativeModel(available_models[0])
    except Exception as e:
        st.error(f"Failed to list models: {e}")
        st.stop()

# --- 4. HELPER FUNCTIONS ---

def get_research(company):
    """Hunter: Real-time strategic problem identification."""
    query = f"Provide a SWOT analysis for {company} in 2026. Identify one specific $100M+ operational problem or financial leak."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": query}]
    )
    return response.choices[0].message.content

def get_slide_script(company, research):
    """Architect: Writes the 15-slide script using the best found model."""
    model = get_best_available_model()
    prompt = f"""
    You are Shubham Verma, a Strategy Director candidate. 
    Research Data: {research}
    Create a 15-slide pitch deck for {company}. 
    Format: STRICTLY a Python list of dicts: [{{'title': '...', 'bullets': ['...']}}]
    Include 15 slides: Title, Summary, Context, AI Solution, ROI, and The Ask.
    """
    response = model.generate_content(prompt)
    # Cleaning common AI formatting junk
    clean_text = response.text.replace("```python", "").replace("```", "").replace("json", "").strip()
    return eval(clean_text)

def get_best_layout(prs, keywords):
    """Finds BCG layouts by searching names."""
    for layout in prs.slide_layouts:
        if any(word.lower() in layout.name.lower() for word in keywords):
            return layout
    return prs.slide_layouts[1] # Standard Content layout fallback

# --- 5. EXECUTION ---

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
            
            # 1. CLEAR OLD CONTENT
            for _ in range(len(prs.slides)):
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]

            # 2. POPULATE SLIDES
            today = datetime.now().strftime("%B %d, 2026")
            
            for i, s_data in enumerate(script):
                layout = get_best_layout(prs, ["Title", "Cover"]) if i == 0 else get_best_layout(prs, ["Content", "Bullet"])
                slide = prs.slides.add_slide(layout)
                
                # Title
                if slide.shapes.title:
                    slide.shapes.title.text = s_data['title']
                
                # Find main text body
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
                        p.text = str(bullet)
                
                # Title Slide Personalization
                if i == 0 and body_ph:
                    body_ph.text = f"Presented to: CEO of {target_company}\nBy: Shubham Verma\nCandidate: Strategy & AI Leadership\nDate: {today}"

            # 3. SAVE & DOWNLOAD
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
            status.update(label="✅ BCG Deck Prepared!", state="complete")

        st.download_button(
            label="📥 Download Strategy Deck",
            data=ppt_io,
            file_name=f"{target_company}_Strategy_Verma.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
