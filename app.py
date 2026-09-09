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

st.set_page_config(page_title="PaperCraft AI — Bilingual Paper Generator", page_icon="🎓", layout="wide")

st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
.stApp{background:#f8fafc;color:#0f172a;font-family:Inter,sans-serif}.block-container{max-width:1180px;padding:1rem 2rem 4rem}
.nav{display:flex;align-items:center;justify-content:space-between;padding:10px 0 22px}.brand{display:flex;align-items:center;gap:10px;font-weight:800;font-size:20px}.brandmark{width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,#2563eb,#7c3aed);display:grid;place-items:center;color:white}.navlinks{display:flex;gap:28px;color:#64748b;font-size:14px;font-weight:600}.pill{border-radius:999px;padding:7px 12px;background:#ecfdf5;color:#047857;font-size:12px;font-weight:700}
.hero{border-radius:30px;padding:64px 56px;background:radial-gradient(circle at 85% 15%,#60a5fa55 0,transparent 28%),radial-gradient(circle at 10% 90%,#a78bfa44 0,transparent 30%),linear-gradient(135deg,#0f172a,#1e3a8a 58%,#312e81);color:#fff;box-shadow:0 24px 60px #0f172a20}.eyebrow{display:inline-block;background:#ffffff14;border:1px solid #ffffff22;border-radius:999px;padding:8px 14px;font-size:12px;font-weight:700;margin-bottom:18px}.hero h1{font-size:50px;line-height:1.06;letter-spacing:-2px;max-width:820px;margin:0 0 18px}.hero h1 span{color:#93c5fd}.hero p{font-size:17px;line-height:1.7;color:#dbeafe;max-width:720px}.points{display:flex;gap:10px;flex-wrap:wrap}.points span{background:#ffffff10;border:1px solid #ffffff1c;padding:8px 11px;border-radius:10px;font-size:13px}
.section{text-align:center;padding:42px 0 16px}.kicker{color:#2563eb;font-size:12px;font-weight:800;letter-spacing:1.3px;text-transform:uppercase}.section h2{font-size:30px;letter-spacing:-.8px;margin:7px 0 9px}.section p{color:#64748b;max-width:650px;margin:auto;line-height:1.6}.feature,.step{background:#fff;border:1px solid #e2e8f0;border-radius:18px;padding:23px;height:100%;box-shadow:0 8px 28px #0f172a08}.feature-icon{width:42px;height:42px;border-radius:12px;background:#eff6ff;display:grid;place-items:center;font-size:20px;margin-bottom:14px}.feature h3,.step h3{margin:0 0 7px;font-size:16px}.feature p,.step p{margin:0;color:#64748b;font-size:13px;line-height:1.55}.step{text-align:center}.stepno{display:inline-block;color:#2563eb;background:#eff6ff;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:800;margin-bottom:10px}
.upload{background:#fff;border:1px solid #dbe4f0;border-radius:24px;padding:30px;box-shadow:0 16px 45px #0f172a0b;margin-top:22px}.upload h2{text-align:center;margin:0 0 7px;font-size:24px}.upload .sub{text-align:center;color:#64748b;font-size:14px;margin-bottom:20px}.metric{background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;padding:14px;text-align:center}.metric small{display:block;color:#64748b;font-size:10px;text-transform:uppercase;font-weight:800;letter-spacing:.7px}.metric b{font-size:23px}
.stButton>button{border-radius:12px!important;height:46px!important;font-weight:700!important;background:linear-gradient(135deg,#2563eb,#4f46e5)!important;border:0!important;box-shadow:0 8px 20px #4f46e533!important}.stDownloadButton>button{border-radius:12px!important;height:46px!important;font-weight:700!important;background:linear-gradient(135deg,#059669,#047857)!important;border:0!important;color:#fff!important}[data-testid="stFileUploader"]{background:#f8fafc;border:1.5px dashed #93c5fd;border-radius:18px;padding:14px}#MainMenu,footer{visibility:hidden}@media(max-width:800px){.navlinks{display:none}.hero{padding:42px 25px}.hero h1{font-size:37px}.block-container{padding:1rem}}
</style>""",unsafe_allow_html=True)

st.markdown('<div class="nav"><div class="brand"><div class="brandmark">✦</div>PaperCraft AI</div><div class="navlinks"><span>Features</span><span>How it works</span><span>Bilingual Papers</span></div><span class="pill">● AI Paper Generator</span></div>',unsafe_allow_html=True)
st.markdown('<div class="hero"><div class="eyebrow">SMARTER PAPER CREATION • ENGLISH → ENGLISH + HINDI</div><h1>Turn any question paper into a <span>professional bilingual paper.</span></h1><p>Upload your English PDF and PaperCraft AI extracts real questions, preserves scientific notation and figures, translates the content into Hindi, and prepares a clean two-column Word document.</p><div class="points"><span>✓ Smart question detection</span><span>✓ Hindi translation</span><span>✓ Math & chemistry protection</span><span>✓ Figures preserved</span></div></div>',unsafe_allow_html=True)
st.markdown('<div class="section"><div class="kicker">Built for educators</div><h2>Everything you need to create bilingual papers</h2><p>A focused workflow for teachers, coaching institutes and academic teams.</p></div>',unsafe_allow_html=True)
a,b,c=st.columns(3)
with a: st.markdown('<div class="feature"><div class="feature-icon">📄</div><h3>Smart PDF Extraction</h3><p>Ignores instructions, headers, footers and option numbering while detecting the real question sequence.</p></div>',unsafe_allow_html=True)
with b: st.markdown('<div class="feature"><div class="feature-icon">🧠</div><h3>Science-Aware Translation</h3><p>Keeps equations, units, symbols and scientific notation out of ordinary translation.</p></div>',unsafe_allow_html=True)
with c: st.markdown('<div class="feature"><div class="feature-icon">📝</div><h3>Visual Fidelity</h3><p>Preserves embedded figures and reaction schemes from the original PDF.</p></div>',unsafe_allow_html=True)
st.markdown('<div class="section"><div class="kicker">Simple workflow</div><h2>From PDF to paper in three steps</h2></div>',unsafe_allow_html=True)
a,b,c=st.columns(3)
with a: st.markdown('<div class="step"><div class="stepno">01</div><h3>Upload</h3><p>Choose your English question-paper PDF.</p></div>',unsafe_allow_html=True)
with b: st.markdown('<div class="step"><div class="stepno">02</div><h3>Generate</h3><p>Questions are detected, translated and formatted.</p></div>',unsafe_allow_html=True)
with c: st.markdown('<div class="step"><div class="stepno">03</div><h3>Download</h3><p>Get the finished bilingual Word document.</p></div>',unsafe_allow_html=True)

QRE=re.compile(r'^\s*(\d{1,3})\.\s*(.*)$')
HEADERS=("PHYSICS","CHEMISTRY","MATHEMATICS","BIOLOGY")
SECTIONS=("SINGLE CORRECT","MULTIPLE CORRECT","NUMERICAL VALUE","ASSERTION","PARAGRAPH","MATCH")
FORMULA=set("∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω")
TECH=re.compile(r'\b(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|dA|dB|dV|vec)\b',re.I)
CHEM=re.compile(r'\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|CH3MgBr|LiAlH4|NaBH4|ATP|DNA|RNA)\b',re.I)

@st.cache_data(show_spinner=False)
def pdf_state(data): return hashlib.sha256(data).hexdigest()

@st.cache_data(show_spinner=False)
def pages_of(data):
    pdf=fitz.open(stream=data,filetype="pdf"); out=[]
    for pno,p in enumerate(pdf): out.append({"page":pno,"width":p.rect.width,"height":p.rect.height,"words":p.get_text("words",sort=True)})
    return out

def lines_of(page):
    groups=[]
    for w in page["words"]:
        if len(w)<5: continue
        x0,y0,x1,y1,txt=w[:5]
        g=next((g for g in groups if abs(g["y"]-y0)<=2.5),None)
        if g: g["ws"].append(w); g["y"]=(g["y"]+y0)/2
        else: groups.append({"y":y0,"ws":[w]})
    out=[]
    for g in sorted(groups,key=lambda z:z["y"]):
        ws=sorted(g["ws"],key=lambda w:w[0]); out.append({"y0":min(w[1] for w in ws),"y1":max(w[3] for w in ws),"x0":min(w[0] for w in ws),"x1":max(w[2] for w in ws),"text":" ".join(w[4] for w in ws)})
    return out

def clean(s): return re.sub(r'\s+([,.;:!?])',r'\1',re.sub(r'\s+',' ',s).strip())

def detect_starts(pages):
    candidates=[]
    for p in pages:
        for ln in lines_of(p):
            m=QRE.match(clean(ln["text"]))
            if m and ln["x0"]<p["width"]*.18: candidates.append({**ln,"page":p["page"],"num":int(m.group(1))})
    runs=[]
    for x in candidates:
        if not runs or x["num"]!=runs[-1][-1]["num"]+1: runs.append([x])
        else: runs[-1].append(x)
    runs=[r for r in runs if len(r)>=3]
    return max(runs,key=len) if runs else []

def build_records(pages):
    starts=detect_starts(pages); rec=[]
    for i,s in enumerate(starts):
        e=starts[i+1] if i+1<len(starts) else None; seg=[]; p=s["page"]
        while p<len(pages):
            pg=pages[p]; y0=s["y0"]-1 if p==s["page"] else 18; y1=e["y0"]-4 if e and p==e["page"] else pg["height"]-24
            for ln in lines_of(pg):
                u=ln["text"].upper()
                if y0<ln["y0"]<y1 and ln["x0"]<pg["width"]*.45 and (u.startswith(HEADERS) or any(u.startswith("("+x) for x in SECTIONS)): y1=min(y1,ln["y0"]-3)
            if y1>y0+4: seg.append((p,y0,y1))
            if e and p>=e["page"]: break
            p+=1
        texts=[]
        for pno,ya,yb in seg:
            for ln in lines_of(pages[pno]):
                if ln["y0"]>=ya-1 and ln["y0"]<=yb+1:
                    t=clean(ln["text"])
                    if t and not re.fullmatch(r'(?:I|II)\s*PUC.*MAINS(?:\s*Page\s*\d+)?',t,re.I) and not re.fullmatch(r'Page\s*\d+',t,re.I): texts.append(t)
        rec.append({"num":s["num"],"segments":seg,"text":"\n".join(texts)})
    return rec

def formulaish(t): return bool(FORMULA.intersection(t) or TECH.search(t) or CHEM.search(t) or re.search(r'\d\s*[/^]\s*\d|[A-Za-z]\s*[_^]\s*\d|[A-Za-z]\s*[=<>]',t))

@st.cache_data(show_spinner=False)
def tr(line):
    if not line.strip() or formulaish(line): return line
    try: return GoogleTranslator(source="en",target="hi").translate(line) or line
    except Exception: return line

def crop(data,pno,rect,scale=2):
    pdf=fitz.open(stream=data,filetype="pdf"); pg=pdf[pno]
    r=fitz.Rect(max(0,rect.x0-3),max(0,rect.y0-3),min(pg.rect.width,rect.x1+3),min(pg.rect.height,rect.y1+3))
    return pg.get_pixmap(matrix=fitz.Matrix(scale,scale),clip=r,alpha=False).tobytes("png")

def images_for(data,pages,seg):
    pdf=fitz.open(stream=data,filetype="pdf"); out=[]
    for pno,y0,y1 in seg:
        pg=pdf[pno]
        for im in pg.get_images(full=True):
            for r in pg.get_image_rects(im[0]):
                if r.width*r.height>500 and r.y1>=y0 and r.y0<=y1 and not (r.width>pg.rect.width*.85 and r.height<30): out.append((pno,r))
    return out

def addimg(cell,img,width=3.55):
    p=cell.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run().add_picture(io.BytesIO(img),width=Inches(width))

def add_content(cell,data,pages,r,hindi=False):
    if hindi:
        p=cell.paragraphs[0]; add_run(p,f'{r["num"]}.',True,True)
    for pno,y0,y1 in r["segments"]:
        lines=[ln for ln in lines_of(pages[pno]) if ln["y0"]>=y0-1 and ln["y0"]<=y1+1]; i=0
        while i<len(lines):
            ln=lines[i]; text=clean(ln["text"])
            if not text: i+=1; continue
            if text.startswith(f'{r["num"]}.'): text=text[len(str(r["num"]))+1:].strip()
            if formulaish(text):
                x0,x1=ln["x0"],ln["x1"]; top=ln["y0"]; bottom=ln["y1"]; j=i+1
                while j<len(lines) and lines[j]["y0"]<=bottom+10 and not QRE.match(clean(lines[j]["text"])):
                    x0=min(x0,lines[j]["x0"]); x1=max(x1,lines[j]["x1"]); bottom=min(y1,max(bottom,lines[j]["y1"])); j+=1
                rect=fitz.Rect(max(0,x0-5),max(0,top-4),min(pages[pno]["width"],x1+5),min(pages[pno]["height"],bottom+4))
                addimg(cell,crop(data,pno,rect),3.55); i=j; continue
            p=cell.add_paragraph(); p.paragraph_format.space_after=Pt(2); add_run(p,tr(text) if hindi else text,hindi); i+=1
    for pno,rr in images_for(data,pages,r["segments"]): addimg(cell,crop(data,pno,rr),3.55)

def add_run(par,text,hindi=False,bold=False):
    r=par.add_run(text); r.font.name="Nirmala UI" if hindi else "Calibri"; r.font.size=Pt(9.5); r.bold=bold; r._element.rPr.rFonts.set(qn("w:eastAsia"),"Nirmala UI" if hindi else "Calibri")

def margins(cell):
    tc=cell._tc.get_or_add_tcPr(); m=OxmlElement("w:tcMar")
    for s,v in [("top",90),("bottom",90),("start",110),("end",110)]:
        x=OxmlElement("w:"+s); x.set(qn("w:w"),str(v)); x.set(qn("w:type"),"dxa"); m.append(x)
    tc.append(m)

def docx_make(data,records,pages,progress):
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.3); sec.bottom_margin=Inches(.3); sec.left_margin=Inches(.25); sec.right_margin=Inches(.25)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(p,"BILINGUAL QUESTION PAPER",False,True)
    t=doc.add_table(rows=1,cols=2); t.autofit=False; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(["English","हिन्दी अनुवाद"]):
        c=t.rows[0].cells[i]; c.width=Inches(4); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; margins(c); sh=OxmlElement("w:shd"); sh.set(qn("w:fill"),"E8EEF8"); c._tc.get_or_add_tcPr().append(sh); pp=c.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(pp,h,i==1,True)
    for i,r in enumerate(records):
        cells=t.add_row().cells
        for c in cells: c.width=Inches(4); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP; margins(c)
        add_content(cells[0],data,pages,r,False); add_content(cells[1],data,pages,r,True); progress.progress((i+1)/len(records))
    b=io.BytesIO(); doc.save(b); return b.getvalue()

st.markdown('<div class="upload"><h2>Create your bilingual question paper</h2><div class="sub">Upload an English PDF and generate a clean English + Hindi Word document.</div>',unsafe_allow_html=True)
up=st.file_uploader("Upload English Question Paper (PDF)",type=["pdf"])
st.markdown('</div>',unsafe_allow_html=True)
if up:
    data=up.getvalue(); state=pdf_state(data)
    if st.session_state.get("state")!=state: st.session_state.clear(); st.session_state["state"]=state
    pages=pages_of(data); records=build_records(pages)
    a,b,c=st.columns(3)
    with a: st.markdown(f'<div class="metric"><small>Pages</small><b>{len(pages)}</b></div>',unsafe_allow_html=True)
    with b: st.markdown(f'<div class="metric"><small>Detected questions</small><b>{len(records)}</b></div>',unsafe_allow_html=True)
    with c: st.markdown(f'<div class="metric"><small>Range</small><b>{records[0]["num"]}–{records[-1]["num"]}</b></div>' if records else '<div class="metric"><small>Range</small><b>—</b></div>',unsafe_allow_html=True)
    if not records: st.error("No reliable numbered-question sequence found. This PDF may be scanned/image-only and needs OCR.")
    else:
        st.success(f"{len(records)} questions detected. Instructions and option numbering are excluded.")
        if st.button("🚀 Generate Bilingual Word File",use_container_width=True):
            bar=st.progress(0); st.session_state["docx"]=docx_make(data,records,pages,bar)
        if "docx" in st.session_state: st.download_button("📥 Download Word Document",st.session_state["docx"],"Bilingual_Question_Paper.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
else: st.markdown('<div class="section"><div class="kicker">Ready when you are</div><h2>Upload a paper to get started</h2><p>Your PDF is processed in the browser session and converted into a bilingual Word document.</p></div>',unsafe_allow_html=True)
