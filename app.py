import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches
import io
import matplotlib.pyplot as plt
import pandas as pd
from pathlib import Path
from datetime import datetime

# --- CONFIG & BCG BRANDING ---
st.set_page_config(page_title="Executive Strategy Agent", layout="centered")
BCG_TEAL = "#00453c"
BCG_GREY = "#545454"

# --- AI CLIENTS ---
try:
    pplx_client = OpenAI(api_key=st.secrets["PPLX_KEY"], base_url="https://api.perplexity.ai")
    genai.configure(api_key=st.secrets["GEMINI_KEY"])
    # Using stable 2026 model identifiers
    strategy_model = genai.GenerativeModel("gemini-1.5-pro")
except Exception as e:
    st.error("Credential Error: Check your Streamlit Secrets.")
    st.stop()

# --- HELPER: DATA & CHARTS ---
def get_financial_csv(company):
    """Hunter: Fetches real numbers for charting."""
    query = f"Provide a CSV table of {company}'s revenue and net income for the last 5 years. Use 'Year,Revenue,Income' as headers. Output CSV only."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": query}]
    )
    return response.choices[0].message.content

def build_bcg_chart(csv_data):
    """Analyst: Generates a chart matching your BCG template."""
    try:
        df = pd.read_csv(io.StringIO(csv_data.strip()))
        fig, ax = plt.subplots(figsize=(6, 4))
        df.plot(kind='bar', x='Year', ax=ax, color=[BCG_TEAL, BCG_GREY])
        plt.title("Financial Performance Analysis (2021-2025)", fontsize=12, fontweight='bold', color=BCG_TEAL)
        plt.grid(axis='y', linestyle='--', alpha=0.3)
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', transparent=True, dpi=300)
        return img_buf
    except:
        return None

# --- MAIN INTERFACE ---
st.title("🏛️ Strategic Leadership Agent")
st.write("Candidate: **Shubham Verma** | Strategy & AI Leadership")

target_company = st.text_input("🏢 Enter Target Company Name:", placeholder="e.g. Nike, Tesla, Starbucks")

if target_company and st.button("🚀 GENERATE DATA-DRIVEN PROPOSAL"):
    with st.status("Agent Executing...", expanded=True) as status:
        
        # 1. DATA & CHARTS
        st.write("🔍 Extracting financial data from Perplexity...")
        csv_raw = get_financial_csv(target_company)
        chart_buffer = build_bcg_chart(csv_raw)
        
        # 2. STRATEGY SCRIPT
        st.write("🧠 Architecting AI strategy with Gemini 1.5 Pro...")
        research_query = f"Analyze {target_company} for 2026. Find the #1 technical debt problem costing them millions."
        research_res = pplx_client.chat.completions.create(model="sonar-pro", messages=[{"role": "user", "content": research_query}]).choices[0].message.content
        
        prompt = f"Using this research: {research_res}, write a 15-slide pitch for a Director role at {target_company}. Format: Python list of dicts [{{'title': '...', 'bullets': ['...']}}]."
        script_raw = strategy_model.generate_content(prompt).text
        script = eval(script_raw.strip("`python\n").strip("`"))
        
        # 3. PPT ASSEMBLY
        st.write("🎨 Applying BCG Master Template...")
        tpl_path = Path(__file__).parent / "master_template.pptx"
        prs = Presentation(tpl_path) if tpl_path.exists() else Presentation()
        
        today = datetime.now().strftime("%B %d, 2026")
        
        for i, s_data in enumerate(script):
            if i >= len(prs.slides): break
            slide = prs.slides[i]
            
            # Update Title
            if slide.shapes.title:
                slide.shapes.title.text = s_data['title']
            
            # Replace Content Placeholders
            for shape in slide.shapes:
                if shape.has_text_frame and shape != slide.shapes.title:
                    if len(shape.text_frame.text) > 2 or "Click to add" in shape.text_frame.text:
                        shape.text_frame.text = ""
                        for b in s_data['bullets']:
                            p = shape.text_frame.add_paragraph()
                            p.text = str(b)
                        break

            # Inject Chart into Slide 3
            if i == 2 and chart_buffer:
                slide.shapes.add_picture(chart_buffer, Inches(5.5), Inches(2), width=Inches(4))

            # Branding Slide 1
            if i == 0:
                for shape in slide.shapes:
                    if shape.has_text_frame and shape != slide.shapes.title:
                        shape.text_frame.text = f"Presented to: CEO of {target_company}\nBy: Shubham Verma\nStrategy Director Candidate\n{today}"
                        break
        
        # 4. FINALIZE
        ppt_out = io.BytesIO()
        prs.save(ppt_out)
        ppt_out.seek(0)
        status.update(label="✅ Strategy Deck Ready!", state="complete")

    st.download_button("📥 Download BCG Deck", ppt_out, f"{target_company}_Strategy_Verma.pptx")
