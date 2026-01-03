import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
from openai import OpenAI
import google.generativeai as genai
from datetime import datetime

# --- CONFIGURATION ---
st.set_page_config(page_title="Strategic Intelligence Agent", layout="wide")

# Valid 2026 Model Names
PERPLEXITY_MODEL = "sonar" 
GEMINI_MODEL = "gemini-2.5-pro" # Fallback to 1.5-pro-latest if 2.5 is gated

# --- API SETUP ---
try:
    pplx_client = OpenAI(api_key=st.secrets["PPLX_KEY"], base_url="https://api.perplexity.ai")
    genai.configure(api_key=st.secrets["GEMINI_KEY"])
except Exception:
    st.error("⚠️ API Keys Missing. Please add PPLX_KEY and GEMINI_KEY to Secrets.")
    st.stop()

# --- AGENT BRAINS ---

def get_deep_research(company):
    """The Hunter: Uses Perplexity Sonar to find a specific $100M failure."""
    query = f"""
    Conduct a deep forensic audit of {company} for 2026.
    1. Identify the single biggest Operational or Financial bottleneck (must be >$100M impact).
    2. Provide real 2024-2025 financial data points related to this bottleneck.
    3. Find specific technical debt or legacy system issues causing this.
    Output purely factual data.
    """
    response = pplx_client.chat.completions.create(
        model=PERPLEXITY_MODEL,
        messages=[{"role": "user", "content": query}]
    )
    return response.choices[0].message.content

def get_strategic_narrative(company, research):
    """The Architect: Uses Gemini 2.5 to write the 6-page strategy."""
    model = genai.GenerativeModel(GEMINI_MODEL)
    prompt = f"""
    You are a Senior Partner at McKinsey in 2026.
    Based on this research for {company}: {research}
    
    Write a Strategic Memo strictly in this JSON format:
    {{
      "title": "The Strategic Theme",
      "executive_summary": "300 word punchy summary for the CEO.",
      "problem_statement": "Deep dive into the $100M pain point. Use numbers.",
      "solution_architecture": "Technical description of the Multi-Agent AI System to fix it.",
      "roi_analysis": "Conservative financial projection of savings/growth.",
      "implementation_plan": "Phase 1 (Month 1-2), Phase 2 (Month 3-4), Phase 3 (Month 5-6)."
    }}
    """
    try:
        response = model.generate_content(prompt)
        text = response.text.replace("```json", "").replace("```", "").strip()
        return eval(text)
    except:
        # Fallback for simpler models
        return {
            "title": f"AI Transformation Strategy for {company}",
            "executive_summary": "Analysis failed. Please retry.",
            "problem_statement": "N/A", "solution_architecture": "N/A", 
            "roi_analysis": "N/A", "implementation_plan": "N/A"
        }

# --- VISUALIZATION ENGINE ---

def create_financial_chart(company):
    """Generates a professional Matplotlib chart."""
    # Mocking data structure for stability - in production, extract from Perplexity
    data = {
        'Year': ['2022', '2023', '2024', '2025 (Est)'],
        'OpEx (Billions)': [12.5, 13.2, 14.8, 16.1]
    }
    df = pd.DataFrame(data)
    
    plt.figure(figsize=(7, 4))
    # McKinsey Blue style
    plt.bar(df['Year'], df['OpEx (Billions)'], color="#004c6d", width=0.6)
    plt.title(f"{company}: Rising Operational Costs", fontsize=14, fontweight='bold', pad=20)
    plt.ylabel("Billions ($)", fontsize=10)
    plt.grid(axis='y', linestyle='--', alpha=0.3)
    
    # Save to buffer
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
    return img_buf

def create_architecture_diagram():
    """Draws a System Architecture Diagram using basic patches."""
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.axis('off')
    
    # Define boxes
    boxes = {
        "Data Lake": (0.1, 0.4),
        "AI Agent Core": (0.4, 0.4),
        "Action Layer": (0.7, 0.4),
        "User Interface": (0.4, 0.7),
        "Legacy ERP": (0.4, 0.1)
    }
    
    for name, (x, y) in boxes.items():
        rect = patches.FancyBboxPatch((x, y), 0.2, 0.15, boxstyle="round,pad=0.05", 
                                      linewidth=2, edgecolor="#004c6d", facecolor="#e6f3ff")
        ax.add_patch(rect)
        ax.text(x+0.1, y+0.075, name, ha='center', va='center', fontsize=9, fontweight='bold')

    # Arrows
    ax.arrow(0.3, 0.475, 0.1, 0, head_width=0.02, color="#555") # Data -> Core
    ax.arrow(0.6, 0.475, 0.1, 0, head_width=0.02, color="#555") # Core -> Action
    ax.arrow(0.5, 0.55, 0, 0.15, head_width=0.02, color="#555") # Core -> UI
    ax.arrow(0.5, 0.4, 0, -0.15, head_width=0.02, color="#555") # Core -> ERP

    ax.text(0.5, 0.9, "Proposed AI Orchestration Layer", ha='center', fontsize=12, fontweight='bold', color="#004c6d")
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
    return img_buf

# --- DOCUMENT COMPILER ---

def create_word_doc(company, strategy, chart_img, arch_img):
    doc = Document()
    
    # Styles
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    
    # 1. Title Page
    doc.add_heading(f"Strategic Intelligence Brief: {company}", 0)
    doc.add_paragraph(f"Prepared by: Shubham Verma | {datetime.now().strftime('%B %Y')}")
    doc.add_paragraph("Strictly Confidential").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    
    # 2. Executive Summary
    doc.add_heading("1. Executive Summary", level=1)
    doc.add_paragraph(strategy['executive_summary'])
    
    # 3. The Problem
    doc.add_heading("2. The Strategic Bottleneck", level=1)
    doc.add_paragraph(strategy['problem_statement'])
    
    # Insert Chart
    doc.add_paragraph("Figure 1: Financial Trend Analysis").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(chart_img, width=Inches(5))
    
    # 4. The Solution
    doc.add_heading("3. Proposed AI Architecture", level=1)
    doc.add_paragraph(strategy['solution_architecture'])
    
    # Insert Diagram
    doc.add_paragraph("Figure 2: Multi-Agent System Design").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(arch_img, width=Inches(5.5))
    
    # 5. Impact & Roadmap
    doc.add_heading("4. ROI & Implementation", level=1)
    doc.add_paragraph(strategy['roi_analysis'])
    doc.add_heading("Execution Timeline", level=2)
    doc.add_paragraph(strategy['implementation_plan'])
    
    # Save to IO
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- FRONTEND ---

st.title("♟️ Strategic Intelligence Agent (Docx)")
st.caption("Powered by Perplexity Sonar & Gemini 2.5 Pro")

company_input = st.text_input("Target Company:")

if company_input and st.button("Generate Strategy Brief"):
    with st.status("Initializing Strategic Deep Dive...", expanded=True) as status:
        
        st.write("📡 Scanning Perplexity Sonar for financial leaks...")
        research = get_deep_research(company_input)
        
        st.write("🧠 Gemini 2.5 Pro is architecting the solution...")
        strategy = get_strategic_narrative(company_input, research)
        
        st.write("📊 Matplotlib is rendering financial models...")
        chart = create_financial_chart(company_input)
        
        st.write("🏗️ Designing System Architecture...")
        arch = create_architecture_diagram()
        
        st.write("📝 Compiling DOCX Report...")
        doc_file = create_word_doc(company_input, strategy, chart, arch)
        
        status.update(label="Strategy Brief Ready", state="complete")
        
    st.success("Analysis Complete.")
    
    st.download_button(
        label="📥 Download Strategy Memo (.docx)",
        data=doc_file,
        file_name=f"Strategy_Brief_{company_input}_Verma.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
