import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from pptx import Presentation
import io

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Executive Strategy Agent", page_icon="🏛️", layout="centered")

# CSS to make it look professional
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏛️ CEO Strategy Agent")
st.subheader("Turn a company name into a 15-slide leadership pitch.")

# --- 2. KEYS & CONFIG ---
# These should be in your Streamlit Secrets (Advanced Settings)
try:
    PPLX_KEY = st.secrets["PPLX_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
except:
    st.error("Missing API Keys! Add PPLX_KEY and GEMINI_KEY to Streamlit Secrets.")
    st.stop()

# Initialize AI Clients
pplx_client = OpenAI(api_key=PPLX_KEY, base_url="https://api.perplexity.ai")
genai.configure(api_key=GEMINI_KEY)
gemini_model = genai.GenerativeModel('gemini-1.5-pro')

# --- 3. HELPER FUNCTIONS ---

def get_research(company):
    """The Hunter: Uses Perplexity to find real-time problems."""
    prompt = f"Find the top 3 massive strategic/financial failures for {company} in 2025-2026. Focus on missed targets and technical debt. Give specific $ numbers."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

def get_slide_script(company, research):
    """The Architect: Uses Gemini to write the narrative."""
    prompt = f"""
    Based on this research: {research}
    Create a 15-slide executive pitch for a leadership role at {company}.
    Format the output strictly as a Python list of dictionaries:
    [
      {{"title": "Slide Title", "bullets": ["Point 1", "Point 2", "Point 3"]}},
      ...
    ]
    Slide 1 must be a Title Slide. Slide 15 must be 'The Ask' for a leadership role.
    """
    response = gemini_model.generate_content(prompt)
    # Clean up the response to ensure it's valid Python code
    content = response.text.replace("```python", "").replace("```", "").strip()
    return eval(content)

# --- 4. THE INTERFACE ---

company_name = st.text_input("🏢 Enter Company Name (e.g., Nike, Tesla, Uber):")

if company_name:
    if st.button("🚀 Generate Executive Pitch"):
        with st.status("Agent is working...", expanded=True) as status:
            
            # Step 1: Research
            st.write("🔍 Searching for company pain points...")
            research_data = get_research(company_name)
            
            # Step 2: Strategy
            st.write("🧠 Architecting the AI solution...")
            slide_script = get_slide_script(company_name, research_data)
            
            # Step 3: Build PPT in Memory (Fixes FileNotFoundError)
            st.write("🎨 Applying your professional template...")
            ppt_buffer = io.BytesIO()
            
            try:
                # IMPORTANT: master_template.pptx must be in your GitHub folder!
                prs = Presentation("master_template.pptx")
            except:
                st.warning("Master template not found. Using basic layout.")
                prs = Presentation()

            for slide_info in slide_script:
                slide_layout = prs.slide_layouts[1] # Title and Content layout
                slide = prs.slides.add_slide(slide_layout)
                slide.shapes.title.text = slide_info['title']
                
                tf = slide.placeholders[1].text_frame
                for point in slide_info['bullets']:
                    p = tf.add_paragraph()
                    p.text = point
            
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)
            
            status.update(label="✅ Strategy Deck Ready!", state="complete")

        # The Download Button (Uses the memory buffer)
        st.download_button(
            label="📂 Download 15-Slide Presentation",
            data=ppt_buffer,
            file_name=f"{company_name}_Executive_Strategy.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
