import streamlit as st
from pathlib import Path
import io
import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import google.generativeai as genai
from openai import OpenAI

# --- 1. CONFIG & BRANDING ---
st.set_page_config(page_title="Strategic Leadership Agent", layout="wide")
BCG_TEAL = "#00453c"
BCG_GREY = "#545454"

# --- 2. AUTHENTICATION ---
try:
    PPLX_KEY = st.secrets["PPLX_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
except KeyError:
    st.error("Missing API keys in Streamlit Secrets (PPLX_KEY, GEMINI_KEY)")
    st.stop()

# Clients
pplx_client = OpenAI(api_key=PPLX_KEY, base_url="https://api.perplexity.ai")
genai.configure(api_key=GEMINI_KEY)

# --- 3. CORE LOGIC ---

def get_best_model():
    """Finds the best available stable model for 2026."""
    models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]
    for m in models:
        try:
            model = genai.GenerativeModel(m)
            # Simple test call
            model.generate_content("test", generation_config={"max_output_tokens": 1})
            return model
        except:
            continue
    return genai.GenerativeModel("gemini-1.5-flash")

def get_research_and_data(company):
    """Hunter: Fetches strategic pain points and financial CSV data."""
    research_query = f"Find the top 3 strategic failures for {company} in 2025. Identify a $100M+ problem. Then, provide a CSV table of their 5-year revenue (2020-2025)."
    response = pplx_client.chat.completions.create(
        model="sonar-pro",
        messages=[{"role": "user", "content": research_query}]
    )
    return response.choices[0].message.content

def create_bcg_chart(research_text):
    """Analyst: Extracts CSV from research and builds a BCG-themed chart."""
    try:
        # Extract CSV block from text
        csv_str = research_text.split("```csv")[1].split("```")[0].strip()
        df = pd.read_csv(io.StringIO(csv_str))
        fig, ax = plt.subplots(figsize=(6, 4))
        df.plot(kind='bar', x=df.columns[0], ax=ax, color=BCG_TEAL)
        plt.title("Financial Trajectory Analysis", color=BCG_TEAL, fontweight='bold')
        plt.tight_layout()
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', transparent=True, dpi=300)
        return img_buf
    except:
        return None

# --- 4. INTERFACE ---
st.title("🏛️ Executive Strategy Agent")
st.markdown(f"**Candidate:** Shubham Verma | **Status:** 2026 Leadership Suite Active")

target_company = st.text_input("Enter Company Name:", placeholder="e.g. Nike")

if target_company and st.button("🚀 EXECUTE STRATEGY DECK"):
    with st.status("Agent processing...", expanded=True) as status:
        
        # 1. Intelligence Gathering
        raw_data = get_research_and_data(target_company)
        chart_buf = create_bcg_chart(raw_data)
        
        # 2. Narrative Architecture
        model = get_best_model()
        prompt = f"Using this data: {raw_data}, write 15 slides for {target_company}. Format: Python list of dicts: [{{'title': '...', 'bullets': ['...']}}]."
        script_res = model.generate_content(prompt)
        script = eval(script_res.text.strip("`python\n").strip("`"))

        # 3. PPT Assembly
        tpl_path = Path(__file__).parent / "master_template.pptx"
        prs = Presentation(tpl_path) if tpl_path.exists() else Presentation()
        
        # Overwrite Logic
        for i, s_data in enumerate(script):
            if i >= len(prs.slides): break
            slide = prs.slides[i]
            
            # Title
            if slide.shapes.title:
                slide.shapes.title.text = s_data['title']
            
            # Smart Placeholder Discovery (Fixes KeyError)
            body_shape = next((sh for sh in slide.placeholders if sh.placeholder_format.idx != 0), None)
            if body_shape and body_shape.has_text_frame:
                tf = body_shape.text_frame
                tf.text = ""
                for b in s_data['bullets']:
                    p = tf.add_paragraph()
                    p.text = str(b)

            # Insert Chart on Slide 3
            if i == 2 and chart_buf:
                slide.shapes.add_picture(chart_buf, Inches(5), Inches(2), width=Inches(4.5))

            # Shubham's Signature (Slide 1)
            if i == 0:
                branding = f"Prepared for {target_company} CEO\nBy: Shubham Verma\nDate: {datetime.now().strftime('%d %B %2026')}"
                st.sidebar.success(f"Branding Slide 1 for {target_company}")

        # 4. Finalization
        final_ppt = io.BytesIO()
        prs.save(final_ppt)
        st.download_button("📥 Download BCG Pitch Deck", final_ppt.getvalue(), f"{target_company}_Strategy_Verma.pptx")
        status.update(label="✅ Deck Complete", state="complete")
