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

st.set_page_config(page_title="Bilingual Paper Generator", layout="wide")

st.title("📄 Bilingual Exam Paper Generator")
st.write("अपनी English PDF अपलोड करें और 2-Column (English & Hindi) Word Document डाउनलोड करें।")

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

def create_bilingual_docx(questions):
    doc = Document()
    
    sections = doc.sections
    for section in sections:
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

    progress_bar = st.progress(0)
    total_q = len(questions)

    for idx, q_text in enumerate(questions):
        hi_text = translate_to_hindi(q_text)
        
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
    return buffer

uploaded_file = st.file_uploader("अपनी PDF फ़ाइल चुनें", type=["pdf"])

if uploaded_file is not None:
    st.info("PDF से टेक्स्ट निकाला जा रहा है...")
    raw_text = extract_text_from_pdf(uploaded_file)
    questions = parse_questions(raw_text)
    
    st.success(f"कुल {len(questions)} प्रश्न पाए गए।")
    
    if st.button("Bilingual Word Document बनाएँ"):
        with st.spinner("अनुवाद और Formatting चालू है..."):
            docx_buffer = create_bilingual_docx(questions)
            
            st.download_button(
                label="📥 Word Paper (.docx) डाउनलोड करें",
                data=docx_buffer,
                file_name="bilingual_paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
