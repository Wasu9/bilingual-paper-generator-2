import io, re, hashlib
import fitz
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from deep_translator import GoogleTranslator

st.set_page_config(page_title='PaperCraft AI', page_icon='🎓', layout='wide')
st.markdown('''<style>.stApp{background:linear-gradient(135deg,#f8fafc,#e2e8f0)}.hero{background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:28px;border-radius:20px;color:white;text-align:center;margin-bottom:22px}.hero h1{margin:0}.card{background:white;padding:20px;border-radius:16px;border:1px solid #e2e8f0;box-shadow:0 4px 20px #0000000d}</style>''', unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>🎓 PaperCraft AI</h1><div>English PDF → Professional Bilingual English + Hindi Word Paper</div></div>', unsafe_allow_html=True)

FORMULA_CHARS = set('∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω')
CHEM = re.compile(r'\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|ATP|DNA|RNA)\b', re.I)

@st.cache_data(show_spinner=False)
def pdf_state(data): return hashlib.sha256(data).hexdigest()

@st.cache_data(show_spinner=False)
def analyze_pdf(data):
    pdf=fitz.open(stream=data,filetype='pdf'); pages=[]
    for pno,page in enumerate(pdf):
        blocks=page.get_text('blocks', sort=True)
        words=page.get_text('words', sort=True)
        text='\n'.join(b[4].strip() for b in blocks if b[4].strip())
        pages.append({'page':pno,'rect':page.rect,'blocks':blocks,'words':words,'text':text,'images':page.get_images(full=True)})
    return pages

def clean_lines(text):
    out=[]
    for line in text.splitlines():
        s=re.sub(r'\s+',' ',line).strip()
        if not s: continue
        if re.fullmatch(r'(?:page\s*)?\d+(?:\s*of\s*\d+)?',s,re.I): continue
        if re.search(r'^(shaheen|www\.|www\s|test\s*series|question\s*paper\s*series)',s,re.I) and len(s)<100: continue
        out.append(s)
    return out

def parse_questions(pages):
    lines=[]
    for pg in pages: lines += clean_lines(pg['text'])
    starts=[]
    pat=re.compile(r'^(?:Q(?:uestion)?\s*)?(\d{1,3})\s*[\.)\:]',re.I)
    for i,l in enumerate(lines):
        if pat.match(l): starts.append(i)
    if not starts: return []
    qs=[]
    for n,st_i in enumerate(starts):
        en=starts[n+1] if n+1<len(starts) else len(lines)
        block='\n'.join(lines[st_i:en]).strip()
        if len(block)>=5: qs.append(block)
    return qs

def is_formula(s):
    return bool(FORMULA_CHARS.intersection(s) or CHEM.search(s) or re.search(r'\b(?:[A-Za-z]+\s*[=<>]\s*[-+\w()./]+|\d+\s*/\s*\d+|[A-Za-z]\s*[_^]\s*\d)\b',s))

def protect(text):
    saved=[]
    def repl(m): saved.append(m.group(0)); return f'__PAPERCRAFT_{len(saved)-1}__'
    pattern=re.compile(r'\\(?:[A-Za-z]+\{[^}]*\}|[A-Za-z]+)|\$[^$]+\$|(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|vec|frac)\b|[A-Za-z]+\d+[A-Za-z\d]*|\d+(?:\.\d+)?\s*(?:N|J|W|Pa|V|A|Hz|kg|g|m|cm|mm|s|min|mol|K)\b|[∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω]|https?://\S+',re.I)
    return pattern.sub(repl,text),saved

@st.cache_data(show_spinner=False)
def translate_text(text):
    protected,saved=protect(text)
    try: translated=GoogleTranslator(source='en',target='hi').translate(protected)
    except Exception: translated=protected
    for i,v in enumerate(saved): translated=translated.replace(f'__PAPERCRAFT_{i}__',v)
    return translated

def add_run(par, text, hindi=False, bold=False):
    r=par.add_run(text); r.font.name='Nirmala UI' if hindi else 'Calibri'; r.font.size=Pt(10); r.bold=bold
    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Nirmala UI' if hindi else 'Calibri')
    return r

def set_margins(cell):
    tcPr=cell._tc.get_or_add_tcPr(); mar=OxmlElement('w:tcMar')
    for side,val in [('top',100),('bottom',100),('start',130),('end',130)]:
        x=OxmlElement('w:'+side); x.set(qn('w:w'),str(val)); x.set(qn('w:type'),'dxa'); mar.append(x)
    tcPr.append(mar)

def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)

def page_for_question(q, pages):
    first=q.splitlines()[0]
    for pg in pages:
        if first[:35] and first[:35] in pg['text']: return pg
    return None

def render_question_visuals(data, q, page):
    if not page or not is_formula(q): return None
    pdf=fitz.open(stream=data,filetype='pdf'); p=pdf[page['page']]
    target=q.splitlines()[0][:40]; hits=p.search_for(target)
    if not hits: return None
    r=hits[0]; rect=fitz.Rect(max(0,r.x0-5),max(0,r.y0-3),min(p.rect.width,r.x1+5),min(p.rect.height,r.y1+150))
    pix=p.get_pixmap(matrix=fitz.Matrix(1.7,1.7),clip=rect,alpha=False)
    return pix.tobytes('png')

def build_docx(data, qs, progress):
    pages=analyze_pdf(data); doc=Document()
    sec=doc.sections[0]; sec.top_margin=Inches(.45); sec.bottom_margin=Inches(.45); sec.left_margin=Inches(.45); sec.right_margin=Inches(.45)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(p,'BILINGUAL QUESTION PAPER',False,True).font.size=Pt(14)
    table=doc.add_table(rows=1,cols=2); table.autofit=False; table.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,title in enumerate(['English','हिन्दी अनुवाद']):
        c=table.rows[0].cells[i]; c.width=Inches(3.55); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; set_margins(c); shade(c,'DCE6F1'); p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(p,title,i==1,True)
    result=[]
    for idx,q in enumerate(qs):
        row=table.add_row().cells; page=page_for_question(q,pages); visual=render_question_visuals(data,q,page)
        for c in row: c.width=Inches(3.55); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP; set_margins(c)
        enp=row[0].paragraphs[0]; hip=row[1].paragraphs[0]
        for j,line in enumerate(q.splitlines()):
            if j: enp=row[0].add_paragraph()
            add_run(enp,line)
        hi=translate_text(q)
        for j,line in enumerate(hi.splitlines()):
            if j: hip=row[1].add_paragraph()
            add_run(hip,line,True)
        if visual:
            for c in row:
                pp=c.add_paragraph(); pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.add_run().add_picture(io.BytesIO(visual),width=Inches(2.8))
        result.append((q,hi)); progress.progress((idx+1)/max(1,len(qs)))
    bio=io.BytesIO(); doc.save(bio); bio.seek(0); return bio,result

uploaded=st.file_uploader('📄 Upload English Question Paper (PDF)',type=['pdf'])
if uploaded:
    data=uploaded.getvalue(); state=pdf_state(data)
    if st.session_state.get('pdf_state')!=state:
        st.session_state.clear(); st.session_state['pdf_state']=state
    pages=analyze_pdf(data); qs=parse_questions(pages)
    c1,c2=st.columns(2)
    with c1:
        st.markdown('<div class="card">',unsafe_allow_html=True); st.metric('Pages',len(pages)); st.metric('Detected Questions',len(qs)); st.markdown('</div>',unsafe_allow_html=True)
    with c2:
        if not qs: st.error('No numbered questions detected. This PDF may be scanned/image-only; OCR support is required for this type.')
        else: st.success('PDF structure detected. Ready to generate.')
    if qs and st.button('🚀 Generate Bilingual Word File',use_container_width=True):
        bar=st.progress(0); buf,result=build_docx(data,qs,bar); st.session_state['docx']=buf.getvalue(); st.session_state['result']=result
    if 'docx' in st.session_state:
        st.download_button('📥 Download Word Document',st.session_state['docx'],'Bilingual_Question_Paper.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)
        st.divider(); st.subheader('👀 Preview')
        for en,hi in st.session_state['result'][:3]:
            a,b=st.columns(2); a.info(en); b.success(hi)
else:
    st.info('Upload a PDF to begin.')
