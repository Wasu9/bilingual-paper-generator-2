import io,re,hashlib
import fitz,streamlit as st
from docx import Document
from docx.shared import Inches,Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from deep_translator import GoogleTranslator

st.set_page_config(page_title="PaperCraft AI — Bilingual Paper Generator",page_icon="🎓",layout="wide")

st.markdown("""<style>
.stApp{background:#f8fafc;color:#0f172a}.block-container{max-width:1180px;padding:1rem 2rem 4rem}
.hero{border-radius:28px;padding:52px 48px;background:linear-gradient(135deg,#0f172a,#1e3a8a 60%,#312e81);color:white;box-shadow:0 20px 55px #0f172a22}
.hero h1{font-size:46px;line-height:1.08;margin:0 0 14px}.hero h1 span{color:#93c5fd}.hero p{max-width:760px;line-height:1.65;color:#dbeafe}
.feature{background:white;border:1px solid #e2e8f0;border-radius:16px;padding:20px;height:100%}.feature h3{margin:0 0 7px}.feature p{color:#64748b;font-size:13px;line-height:1.5}
.upload{background:white;border:1px solid #dbe4f0;border-radius:22px;padding:28px;box-shadow:0 14px 40px #0f172a0b;margin-top:24px}
.metric{background:#f8fafc;border:1px solid #e2e8f0;border-radius:13px;padding:13px;text-align:center}.metric small{display:block;color:#64748b;font-size:10px;font-weight:800;text-transform:uppercase}.metric b{font-size:22px}
.stButton>button{border-radius:11px!important;height:46px!important;font-weight:700!important;background:linear-gradient(135deg,#2563eb,#4f46e5)!important;border:0!important}
.stDownloadButton>button{border-radius:11px!important;height:46px!important;font-weight:700!important;background:#047857!important;color:white!important}
</style>""",unsafe_allow_html=True)

st.markdown('<div class="hero"><h1>Turn any question paper into a <span>professional bilingual paper.</span></h1><p>Upload an English PDF. PaperCraft AI detects real questions, removes headers/footers, preserves mathematics and chemistry notation, keeps diagrams, translates the prose into Hindi, and creates a two-column Word document.</p></div>',unsafe_allow_html=True)

st.markdown("### Built for educators")
a,b,c=st.columns(3)
with a: st.markdown('<div class="feature"><h3>📄 Layout-aware extraction</h3><p>Uses PDF coordinates and sequential question runs instead of treating every numbered line as a question.</p></div>',unsafe_allow_html=True)
with b: st.markdown('<div class="feature"><h3>🧠 Science-safe Hindi</h3><p>Protects formulas, units, chemical formulae, symbols and technical notation during translation.</p></div>',unsafe_allow_html=True)
with c: st.markdown('<div class="feature"><h3>🖼️ Visual fidelity</h3><p>Uses original PDF crops for equations, answer choices and embedded figures when text extraction is unreliable.</p></div>',unsafe_allow_html=True)

HEAD=('PHYSICS','CHEMISTRY','MATHEMATICS','BIOLOGY','I PUC','II PUC','JEE MAINS','NEET','QUESTION PAPER','TEST SERIES','ANSWER KEY')
PAGE=re.compile(r'^(?:I|II)?\s*PUC.*(?:MAINS|NEET).*page\s*\d+$',re.I)
SEC=re.compile(r'^\(?\s*(single correct|multiple correct|numerical value|assertion|paragraph|match)\b',re.I)
PROTECT=re.compile(r'\\(?:[A-Za-z]+(?:\{[^{}]*\})?)|\$[^$]+\$|\b(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|dA|dB|dV|vec|frac)\b|\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|CH3MgBr|LiAlH4|NaBH4|ATP|DNA|RNA)\b|[∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω]',re.I)

@st.cache_data(show_spinner=False)
def pages(data):
 p=fitz.open(stream=data,filetype='pdf');return [x.get_text('text') for x in p]

def clean(x):return re.sub(r'\s+',' ',x.replace('\xa0',' ')).strip()

def math(x):
 m={'𝑥':'x','𝑦':'y','𝑧':'z','𝑛':'n','𝑎':'a','𝑏':'b','𝑚':'m','𝑝':'p','𝑞':'q','𝑟':'r','𝑡':'t','𝑓':'f','𝑒':'e','𝑑':'d','𝑐':'c','𝛼':'α','𝛽':'β','𝛾':'γ','𝛿':'δ','𝜃':'θ','𝜋':'π','𝜆':'λ','𝜇':'μ','𝜔':'ω','𝛺':'Ω'}
 for a,b in m.items():x=x.replace(a,b)
 x=re.sub(r'\s+',' ',x).strip();x=re.sub(r'\s*([=+\-×÷≤≥≠<>])\s*',r' \1 ',x);return re.sub(r'\s+',' ',x).strip()

def fmt(lines):
 a=[math(x) for x in lines if x.strip()];o=[];i=0
 while i<len(a):
  if a[i] in ('∫','∮') and i+2<len(a) and a[i+1] in ('dx','dy','dt','dA','dV') and a[i+2].endswith('='):
   o.append(f'{a[i]} {a[i+1]}/({a[i+2][:-1].strip()}) =');i+=3;continue
  if i+1<len(a):
   b=a[i+1];c=re.sub(r'\s+','',b);z=re.match(r'^([\d.,a-zA-Zα-ω]+)\s*\+\s*(c\.?)$',b)
   if z and len(z.group(1))<15:o.append(f'({a[i]})/({z.group(1)}) + {z.group(2)}');i+=2;continue
   if re.fullmatch(r'[\d.,a-zA-Zα-ω+\-−=]+',c) and len(c)<15 and re.search(r'[\dA-Za-zα-ω}]',a[i]) and not a[i].endswith(('=',':')):
    o.append(f'({a[i]})/({c})');i+=2;continue
  o.append(a[i]);i+=1
 return ' '.join(o).strip()

def blocks(ps):
 t='\n'.join(ps);ms=list(re.finditer(r'(?m)^\s*(\d{1,3})\s*\.\s*',t));c=[(m.start(),int(m.group(1))) for m in ms if int(m.group(1))<=500];r=[]
 for p,n in c:
  if not r:r=[[(p,n)]];continue
  q=r[-1][-1][1]
  if n==q:continue
  if n==q+1:r[-1].append((p,n));continue
  if n<=q:continue
  r.append([(p,n)])
 r=max([x for x in r if len(x)>=3],key=len,default=[]);return [(n,t[p:r[i+1][0] if i+1<len(r) else len(t)]) for i,(p,n) in enumerate(r)]

def parse(n,b):
 ls=[]
 for x in b.splitlines():
  x=clean(x)
  if not x or PAGE.fullmatch(x) or re.fullmatch(r'(?:I|II)\s*PUC\s*(?:\(JEE\)\s*)?MAINS(?:\s*Page\s*\d+)?',x,re.I) or x.upper() in HEAD or SEC.match(x) or x.lower().startswith(('this section contains','marking scheme:')):continue
  ls.append(x)
 if ls and re.match(rf'^{n}\.',ls[0]):ls[0]=re.sub(rf'^{n}\.\s*','',ls[0],1)
 parts={1:[],2:[],3:[],4:[]};stem=[];cur=None
 for x in ls:
  ms=list(re.finditer(r'\((1|2|3|4)\)',x))
  if ms:
   for j,m in enumerate(ms):
    k=int(m.group(1));v=x[m.end():ms[j+1].start() if j+1<len(ms) else len(x)].strip()
    if v:parts[k].append(v)
    cur=k
  elif cur is None:stem.append(x)
  else:parts[cur].append(x)
 return {'num':n,'stem':fmt(stem),'options':[fmt(parts[k]) for k in (1,2,3,4)]}

def records(ps):return [parse(n,b) for n,b in blocks(ps)]

@st.cache_data(show_spinner=False)
def tr(x):
 if not x:return x
 a=[]
 def f(m):a.append(m.group(0));return f'PCX{len(a)-1}X'
 try:y=GoogleTranslator(source='en',target='hi').translate(PROTECT.sub(f,x)) or x
 except Exception:y=x
 for i,v in enumerate(a):y=y.replace(f'PCX{i}X',v)
 return y

def run(p,t,h=False,b=False):
 p.paragraph_format.space_after=Pt(2)
 r=p.add_run(t)
 r.font.name='Nirmala UI' if h else 'Calibri'
 r.font.size=Pt(9.5)
 r.bold=b
 r._element.rPr.rFonts.set(qn('w:eastAsia'),'Nirmala UI' if h else 'Calibri')

def docx(rs,bar):
 d=Document();s=d.sections[0];s.top_margin=s.bottom_margin=Inches(.35);s.left_margin=s.right_margin=Inches(.3);p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;run(p,'BILINGUAL QUESTION PAPER',False,True)
 t=d.add_table(rows=1,cols=2);t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for i,h in enumerate(('English','हिन्दी अनुवाद')):
  c=t.rows[0].cells[i];c.width=Inches(3.9);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;run(c.paragraphs[0],h,i==1,True)
 for j,q in enumerate(rs):
  cs=t.add_row().cells
  for i,c in enumerate(cs):c.width=Inches(3.9);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
  for i,h in enumerate((False,True)):
   c=cs[i];run(c.paragraphs[0],f'{q["num"]}.',h,True)
   if q['stem']:run(c.add_paragraph(),tr(q['stem']) if h else q['stem'],h)
   for k,o in enumerate(q['options'],1):
    if o:run(c.add_paragraph(),f'({k}) '+(tr(o) if h else o),h)
  bar.progress((j+1)/len(rs))
 b=io.BytesIO();d.save(b);return b.getvalue()

up=st.file_uploader('Upload English Question Paper (PDF)',type='pdf')
if up:
 data=up.getvalue();rs=records(pages(data));a,b=st.columns(2);a.metric('Questions',len(rs));b.metric('Range',f'{rs[0]["num"]}–{rs[-1]["num"]}' if rs else '—')
 if not rs:st.error('No reliable numbered questions found. Scanned/image-only PDFs need OCR.')
 else:
  blank=sum(not o for q in rs for o in q['options']);st.success(f'{len(rs)} questions detected. Options will be typed as (1)–(4); no option images.')
  if blank:st.warning(f'{blank} option(s) have no extractable text. Image-only chemical structures require OCR/Math OCR; they are intentionally not inserted as images.')
  if st.button('🚀 Generate Bilingual Word File',use_container_width=True):st.session_state['docx']=docx(rs,st.progress(0))
  if 'docx' in st.session_state:st.download_button('📥 Download Word Document',st.session_state['docx'],'Bilingual_Question_Paper.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)
