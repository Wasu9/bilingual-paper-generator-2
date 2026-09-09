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

st.set_page_config(page_title="PaperCraft AI — Bilingual Paper Generator", page_icon="🎓", layout="wide", initial_sidebar_state="collapsed")

st.markdown(r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
.stApp{background:#f8fafc;font-family:Inter,sans-serif;color:#0f172a}
.block-container{max-width:1180px;padding:1rem 2rem 4rem}
header[data-testid="stHeader"]{background:transparent}
.nav{display:flex;align-items:center;justify-content:space-between;padding:12px 0 24px}.brand{display:flex;align-items:center;gap:10px;font-weight:800;font-size:20px}.brandmark{width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,#2563eb,#7c3aed);display:grid;place-items:center;color:#fff}.navlinks{display:flex;gap:28px;color:#64748b;font-size:14px;font-weight:600}.pill{border-radius:999px;padding:7px 12px;background:#ecfdf5;color:#047857;font-size:12px;font-weight:700}
.hero{border-radius:30px;padding:62px 56px;background:radial-gradient(circle at 85% 15%,#60a5fa55 0,transparent 28%),radial-gradient(circle at 10% 90%,#a78bfa44 0,transparent 30%),linear-gradient(135deg,#0f172a,#1e3a8a 58%,#312e81);color:#fff;box-shadow:0 24px 60px #0f172a20;overflow:hidden}.eyebrow{display:inline-block;background:#ffffff14;border:1px solid #ffffff22;border-radius:999px;padding:8px 14px;font-size:12px;font-weight:700;margin-bottom:18px}.hero h1{font-size:50px;line-height:1.06;letter-spacing:-2px;max-width:820px;margin:0 0 18px}.hero h1 span{color:#93c5fd}.hero p{font-size:17px;line-height:1.7;color:#dbeafe;max-width:720px;margin:0 0 24px}.points{display:flex;gap:10px;flex-wrap:wrap}.points span{background:#ffffff10;border:1px solid #ffffff1c;padding:8px 11px;border-radius:10px;font-size:13px}
.section{text-align:center;padding:42px 0 16px}.kicker{color:#2563eb;font-size:12px;font-weight:800;letter-spacing:1.3px;text-transform:uppercase}.section h2{font-size:30px;letter-spacing:-.8px;margin:7px 0 9px}.section p{color:#64748b;max-width:650px;margin:auto;line-height:1.6}
.feature,.step{background:#fff;border:1px solid #e2e8f0;border-radius:18px;padding:23px;height:100%;box-shadow:0 8px 28px #0f172a08}.feature-icon{width:42px;height:42px;border-radius:12px;background:#eff6ff;display:grid;place-items:center;font-size:20px;margin-bottom:14px}.feature h3,.step h3{margin:0 0 7px;font-size:16px}.feature p,.step p{margin:0;color:#64748b;font-size:13px;line-height:1.55}.step{text-align:center}.stepno{display:inline-block;color:#2563eb;background:#eff6ff;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:800;margin-bottom:10px}
.upload{background:#fff;border:1px solid #dbe4f0;border-radius:24px;padding:30px;box-shadow:0 16px 45px #0f172a0b;margin-top:22px}.upload h2{text-align:center;margin:0 0 7px;font-size:24px}.upload .sub{text-align:center;color:#64748b;font-size:14px;margin-bottom:20px}.metric{background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;padding:14px;text-align:center}.metric small{display:block;color:#64748b;font-size:10px;text-transform:uppercase;font-weight:800;letter-spacing:.7px}.metric b{font-size:23px}.stButton>button{border-radius:12px!important;height:46px!important;font-weight:700!important;background:linear-gradient(135deg,#2563eb,#4f46e5)!important;border:0!important;box-shadow:0 8px 20px #4f46e533!important}.stDownloadButton>button{border-radius:12px!important;height:46px!important;font-weight:700!important;background:linear-gradient(135deg,#059669,#047857)!important;border:0!important;color:#fff!important}.footer{text-align:center;color:#94a3b8;font-size:12px;padding:40px 0 5px}[data-testid="stFileUploader"]{background:#f8fafc;border:1.5px dashed #93c5fd;border-radius:18px;padding:14px}
@media(max-width:800px){.navlinks{display:none}.hero{padding:42px 25px}.hero h1{font-size:37px}.block-container{padding:1rem}}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="nav"><div class="brand"><div class="brandmark">✦</div>PaperCraft AI</div><div class="navlinks"><span>Features</span><span>How it works</span><span>Bilingual Papers</span></div><span class="pill">● AI Paper Generator</span></div>', unsafe_allow_html=True)
st.markdown('<div class="hero"><div class="eyebrow">SMARTER PAPER CREATION • ENGLISH → ENGLISH + HINDI</div><h1>Turn any question paper into a <span>professional bilingual paper.</span></h1><p>Upload your English PDF and PaperCraft AI extracts questions, protects scientific notation, translates the content into Hindi, and prepares a clean two-column Word document.</p><div class="points"><span>✓ Question extraction</span><span>✓ Hindi translation</span><span>✓ Math & science protection</span><span>✓ Word-ready formatting</span></div></div>', unsafe_allow_html=True)

st.markdown('<div class="section"><div class="kicker">Built for educators</div><h2>Everything you need to create bilingual papers</h2><p>A focused workflow for teachers, coaching institutes and academic teams.</p></div>', unsafe_allow_html=True)
f1,f2,f3=st.columns(3)
with f1: st.markdown('<div class="feature"><div class="feature-icon">📄</div><h3>Smart PDF Extraction</h3><p>Detects numbered questions and removes common headers, page numbers and repeated metadata.</p></div>',unsafe_allow_html=True)
with f2: st.markdown('<div class="feature"><div class="feature-icon">🧠</div><h3>Science-Aware Translation</h3><p>Protects formulas, units, symbols and scientific terms before Hindi translation.</p></div>',unsafe_allow_html=True)
with f3: st.markdown('<div class="feature"><div class="feature-icon">📝</div><h3>Professional Word Output</h3><p>Creates a polished two-column English + Hindi document ready for editing and printing.</p></div>',unsafe_allow_html=True)

st.markdown('<div class="section"><div class="kicker">Simple workflow</div><h2>From PDF to paper in three steps</h2></div>', unsafe_allow_html=True)
s1,s2,s3=st.columns(3)
with s1: st.markdown('<div class="step"><div class="stepno">01</div><h3>Upload</h3><p>Choose your English question-paper PDF.</p></div>',unsafe_allow_html=True)
with s2: st.markdown('<div class="step"><div class="stepno">02</div><h3>Generate</h3><p>Questions are extracted, translated and formatted.</p></div>',unsafe_allow_html=True)
with s3: st.markdown('<div class="step"><div class="stepno">03</div><h3>Download</h3><p>Get the finished bilingual Word document.</p></div>',unsafe_allow_html=True)

FORMULA_CHARS=set('∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω')
CHEM=re.compile(r'\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|ATP|DNA|RNA)\b',re.I)

@st.cache_data(show_spinner=False)
def pdf_state(data): return hashlib.sha256(data).hexdigest()

@st.cache_data(show_spinner=False)
def analyze_pdf(data):
    pdf=fitz.open(stream=data,filetype='pdf'); pages=[]
    for pno,page in enumerate(pdf):
        blocks=page.get_text('blocks',sort=True); words=page.get_text('words',sort=True)
        text='\n'.join(b[4].strip() for b in blocks if b[4].strip())
        pages.append({'page':pno,'rect':page.rect,'blocks':blocks,'words':words,'text':text})
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
    for pg in pages: lines+=clean_lines(pg['text'])
    pat=re.compile(r'^(?:Q(?:uestion)?\s*)?(\d{1,3})\s*[\.)\:]',re.I)
    starts=[]
    for i,line in enumerate(lines):
        if pat.match(line): starts.append(i)
    if not starts: return []
    qs=[]
    for n,start in enumerate(starts):
        end=starts[n+1] if n+1<len(starts) else len(lines)
        block='\n'.join(lines[start:end]).strip()
        if len(block)>=5: qs.append(block)
    return qs

def protect(text):
    saved=[]
    def repl(m): saved.append(m.group(0)); return f'__PC_{len(saved)-1}__'
    pattern=re.compile(r'\\(?:[A-Za-z]+\{[^}]*\}|[A-Za-z]+)|\$[^$]+\$|(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|vec|frac)\b|[A-Za-z]+\d+[A-Za-z\d]*|\d+(?:\.\d+)?\s*(?:N|J|W|Pa|V|A|Hz|kg|g|m|cm|mm|s|min|mol|K)\b|[∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω]|https?://\S+',re.I)
    return pattern.sub(repl,text),saved

@st.cache_data(show_spinner=False)
def translate_text(text):
    protected,saved=protect(text)
    try: translated=GoogleTranslator(source='en',target='hi').translate(protected)
    except Exception: translated=protected
    for i,value in enumerate(saved): translated=translated.replace(f'__PC_{i}__',value)
    return translated

def set_cell_margins(cell):
    tcPr=cell._tc.get_or_add_tcPr(); tcMar=OxmlElement('w:tcMar')
    for side,val in [('top',100),('bottom',100),('start',130),('end',130)]:
        x=OxmlElement('w:'+side); x.set(qn('w:w'),str(val)); x.set(qn('w:type'),'dxa'); tcMar.append(x)
    tcPr.append(tcMar)

def shade(cell,fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)

def add_run(par,text,hindi=False,bold=False):
    r=par.add_run(text); r.font.name='Nirmala UI' if hindi else 'Calibri'; r.font.size=Pt(10); r.bold=bold
    r._element.rPr.rFonts.set(qn('w:eastAsia'),'Nirmala UI' if hindi else 'Calibri'); return r

def build_docx(qs,progress):
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Inches(.45); sec.bottom_margin=Inches(.45); sec.left_margin=Inches(.45); sec.right_margin=Inches(.45)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(p,'BILINGUAL QUESTION PAPER',False,True).font.size=Pt(14)
    table=doc.add_table(rows=1,cols=2); table.autofit=False; table.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,title in enumerate(['English','हिन्दी अनुवाद']):
        c=table.rows[0].cells[i]; c.width=Inches(3.55); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; set_cell_margins(c); shade(c,'DCE6F1'); p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(p,title,i==1,True)
    result=[]
    for idx,q in enumerate(qs):
        cells=table.add_row().cells
        for c in cells: c.width=Inches(3.55); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP; set_cell_margins(c)
        en=cells[0].paragraphs[0]
        for j,line in enumerate(q.splitlines()):
            if j: en=cells[0].add_paragraph()
            add_run(en,line)
        hi=translate_text(q); hp=cells[1].paragraphs[0]
        for j,line in enumerate(hi.splitlines()):
            if j: hp=cells[1].add_paragraph()
            add_run(hp,line,True)
        result.append((q,hi)); progress.progress((idx+1)/max(1,len(qs)))
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf,result

st.markdown('<div class="upload"><h2>Create your bilingual question paper</h2><div class="sub">Upload an English PDF and generate a clean English + Hindi Word document.</div>',unsafe_allow_html=True)
uploaded=st.file_uploader('Upload English Question Paper (PDF)',type=['pdf'])
st.markdown('</div>',unsafe_allow_html=True)

if uploaded:
    data=uploaded.getvalue(); state=pdf_state(data)
    if st.session_state.get('pdf_state')!=state:
        st.session_state.clear(); st.session_state['pdf_state']=state
    pages=analyze_pdf(data); qs=parse_questions(pages)
    m1,m2,m3=st.columns(3)
    with m1: st.markdown(f'<div class="metric"><small>Pages</small><b>{len(pages)}</b></div>',unsafe_allow_html=True)
    with m2: st.markdown(f'<div class="metric"><small>Detected questions</small><b>{len(qs)}</b></div>',unsafe_allow_html=True)
    with m3: st.markdown('<div class="metric"><small>Output format</small><b>DOCX</b></div>',unsafe_allow_html=True)
    st.write('')
    if not qs:
        st.error('No numbered questions detected. This PDF may be scanned/image-only and needs OCR support.')
    else:
        st.success(f'PDF ready — {len(qs)} questions detected.')
        if st.button('🚀 Generate Bilingual Word File',use_container_width=True):
            bar=st.progress(0); buf,result=build_docx(qs,bar); st.session_state['docx']=buf.getvalue(); st.session_state['result']=result
        if 'docx' in st.session_state:
            st.download_button('📥 Download Bilingual Word Document',st.session_state['docx'],'Bilingual_Question_Paper.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)
            st.markdown('<div class="section" style="padding-top:25px"><div class="kicker">Preview</div><h2>First questions</h2></div>',unsafe_allow_html=True)
            for en,hi in st.session_state['result'][:3]:
                a,b=st.columns(2); a.info(en); b.success(hi)

st.markdown('<div class="footer">PaperCraft AI • Bilingual Question Paper Generator</div>',unsafe_allow_html=True)
