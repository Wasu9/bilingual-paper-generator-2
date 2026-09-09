import streamlit as st
import pdfplumber
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from deep_translator import GoogleTranslator
import io
import re

# Page Configuration
st.set_page_config(
    page_title="PaperCraft AI | Bilingual Paper Generator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Premium UI
st.markdown("""
    <style>
    /* Main Background & Fonts */
    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        font-family: 'Inter', sans-serif;
    }
    
    /* Hero Header Banner */
    .hero-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 2.5rem 2rem;
        border-radius: 20px;
        color: white;
        text-align: center;
        box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.4);
        margin-bottom: 2rem;
    }
    .hero-title {
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        letter-spacing: -0.5px;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        opacity: 0.9;
        font-weight: 300;
    }

    /* Cards Styling */
    .custom-card {
        background: white;
        padding: 1.8rem;
        border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
        margin-bottom: 1.5rem;
    }
    
    /* Feature Badge */
    .badge {
        background: #dbeafe;
        color: #1e40af;
        padding: 0.4rem 0.8rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 1rem;
    }

    /* Buttons Styling */
    .stButton>button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.05rem;
        font-weight: 600;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.4);
    }

    /* Download Button */
    .stDownloadButton>button {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        color: white !important;
        border: none !important;
        padding: 0.75rem 2rem !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 12px rgba(5, 150, 105, 0.3) !important;
        width: 100% !important;
    }
    
    /* Hide Streamlit Default Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Top Banner
st.markdown("""
    <div class="hero-header">
        <div class="hero-title">🎓 PaperCraft AI</div>
        <div class="hero-subtitle">Convert English Question Papers to Professional Bilingual (Hindi + English) Word Format</div>
    </div>
""", unsafe_allow_html=True)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def extract_text_from_pdf(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text

def parse_questions(text):
    raw_blocks = re.split(r'\n(?=(?:Q\d+[\.\:]|\d+[\.\:]))', text)
    questions = [b.strip() for b in raw_blocks if b.strip()]
    return questions

def translate_to_hindi(text):
    try:
        translator = GoogleTranslator(source='en', target='hi')
        chunks = text.split('\n')
        translated_chunks = [translator.translate(chunk) if chunk.strip() else "" for chunk in chunks]
        return '\n'.join(translated_chunks)
    except Exception as e:
        return text

def create_bilingual_docx(questions, progress_bar):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
    header_para = doc.add_paragraph()
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header_para.add_run("BILINGUAL QUESTION PAPER")
    run.bold = True
    run.font.size = Pt(14)
    
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "English"
    hdr_cells[1].text = "हिन्दी अनुवाद"
    
    for i in range(2):
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(11)
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)
        
    hdr_cells[0].width = Inches(3.6)
    hdr_cells[1].width = Inches(3.6)

    total_q = len(questions)
    translated_data = []

    for idx, q_text in enumerate(questions):
        hi_text = translate_to_hindi(q_text)
        translated_data.append((q_text, hi_text))
        
        row_cells = table.add_row().cells
        row_cells[0].width = Inches(3.6)
        row_cells[1].width = Inches(3.6)
        
        set_cell_margins(row_cells[0], top=100, bottom=100, left=150, right=150)
        set_cell_margins(row_cells[1], top=100, bottom=100, left=150, right=150)

        p_en = row_cells[0].paragraphs[0]
        p_en.paragraph_format.space_after = Pt(4)
        run_en = p_en.add_run(q_text)
        run_en.font.name = 'Calibri'
        run_en.font.size = Pt(10)

        p_hi = row_cells[1].paragraphs[0]
        p_hi.paragraph_format.space_after = Pt(4)
        run_hi = p_hi.add_run(hi_text)
        run_hi.font.name = 'Mangal'
        run_hi.font.size = Pt(10)
        
        progress_bar.progress((idx + 1) / total_q)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer, translated_data

# Layout Columns
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("""
        <div class="custom-card">
            <span class="badge">STEP 1</span>
            <h3 style="margin-top:0; color:#1e293b;">📄 Upload Question Paper</h3>
            <p style="color:#64748b; font-size:0.95rem;">Select your English PDF paper to convert into a 2-column bilingual layout.</p>
        </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("", type=["pdf"])

with col2:
    st.markdown("""
        <div class="custom-card">
            <span class="badge">STEP 2</span>
            <h3 style="margin-top:0; color:#1e293b;">⚡ Generate & Download</h3>
            <p style="color:#64748b; font-size:0.95rem;">Click process to translate and format into MS Word 2-Column Table.</p>
        </div>
    """, unsafe_allow_html=True)
    
    if uploaded_file is not None:
        raw_text = extract_text_from_pdf(uploaded_file)
        questions = parse_questions(raw_text)
        
        st.success(f"✅ Found {len(questions)} Questions in PDF!")
        
        if st.button("🚀 Generate Bilingual Word File"):
            st.write("Translating & Formatting...")
            progress_bar = st.progress(0)
            
            docx_buffer, translated_data = create_bilingual_docx(questions, progress_bar)
            
            st.session_state['docx_buffer'] = docx_buffer
            st.session_state['translated_data'] = translated_data
            
        if 'docx_buffer' in st.session_state:
            st.download_button(
                label="📥 Download Word Document (.docx)",
                data=st.session_state['docx_buffer'],
                file_name="Bilingual_Question_Paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.info("👈 Please upload a PDF file on the left to get started.")

# Live Preview Section
if 'translated_data' in st.session_state:
    st.markdown("---")
    st.markdown("<h3 style='text-align: center; color: #1e293b;'>👀 Live Paper Preview</h3>", unsafe_allow_html=True)
    
    for idx, (en, hi) in enumerate(st.session_state['translated_data'][:5]):
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            st.info(f"**English:**\n\n{en}")
        with p_col2:
            st.success(f"**हिन्दी:**\n\n{hi}")
    
    if len(st.session_state['translated_data']) > 5:
        st.caption("Showing preview of first 5 questions. Full paper is available in the downloaded Word file.")
