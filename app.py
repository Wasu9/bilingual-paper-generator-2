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
.stApp{background:#f8fafc;color:#0f172a}.block-container{max-width:1180px;padding:1rem 2rem 4rem}
.hero{border-radius:28px;padding:52px 48px;background:linear-gradient(135deg,#0f172a,#1e3a8a 60%,#312e81);color:white;box-shadow:0 20px 55px #0f172a22}
.hero h1{font-size:46px;line-height:1.08;margin:0 0 14px}.hero h1 span{color:#93c5fd}.hero p{max-width:760px;line-height:1.65;color:#dbeafe}
.feature{background:white;border:1px solid #e2e8f0;border-radius:16px;padding:20px;height:100%}.feature h3{margin:0 0 7px}.feature p{color:#64748b;font-size:13px;line-height:1.5}
.upload{background:white;border:1px solid #dbe4f0;border-radius:22px;padding:28px;box-shadow:0 14px 40px #0f172a0b;margin-top:24px}
.metric{background:#f8fafc;border:1px solid #e2e8f0;border-radius:13px;padding:13px;text-align:center}.metric small{display:block;color:#64748b;font-size:10px;font-weight:800;text-transform:uppercase}.metric b{font-size:22px}
.stButton>button{border-radius:11px!important;height:46px!important;font-weight:700!important;background:linear-gradient(135deg,#2563eb,#4f46e5)!important;border:0!important}
.stDownloadButton>button{border-radius:11px!important;height:46px!important;font-weight:700!important;background:#047857!important;color:white!important}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="hero"><h1>Turn any question paper into a <span>professional bilingual paper.</span></h1><p>Upload an English PDF. PaperCraft AI detects real questions, removes headers/footers, preserves mathematics and chemistry notation, keeps diagrams, translates the prose into Hindi, and creates a two-column Word document.</p></div>', unsafe_allow_html=True)

st.markdown("### Built for educators")
a,b,c=st.columns(3)
with a: st.markdown('<div class="feature"><h3>📄 Layout-aware extraction</h3><p>Uses PDF coordinates and sequential question runs instead of treating every numbered line as a question.</p></div>',unsafe_allow_html=True)
with b: st.markdown('<div class="feature"><h3>🧠 Science-safe Hindi</h3><p>Protects formulas, units, chemical formulae, symbols and technical notation during translation.</p></div>',unsafe_allow_html=True)
with c: st.markdown('<div class="feature"><h3>🖼️ Visual fidelity</h3><p>Uses original PDF crops for equations, answer choices and embedded figures when text extraction is unreliable.</p></div>',unsafe_allow_html=True)

QRE = re.compile(r'^\s*(\d{1,3})\s*[\.\):]\s*(.*)$')
HEADER_WORDS = ("PHYSICS","CHEMISTRY","MATHEMATICS","BIOLOGY","I PUC","II PUC","JEE MAINS","NEET","QUESTION PAPER","TEST SERIES","ANSWER KEY")
FOOTER_RE = re.compile(r'^(?:page\s*)?\d+(?:\s*of\s*\d+)?$', re.I)
PAGE_RE = re.compile(r'^(?:I|II)?\s*PUC.*(?:MAINS|NEET).*page\s*\d+$', re.I)
SECTION_RE = re.compile(r'^\(?\s*(single correct|multiple correct|numerical value|assertion|paragraph|match)\b', re.I)
FORMULA_CHARS = set("∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω")
TECH_RE = re.compile(r'\b(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|dA|dB|dV|vec|frac)\b', re.I)
CHEM_RE = re.compile(r'\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|CH3MgBr|LiAlH4|NaBH4|ATP|DNA|RNA)\b', re.I)
MATH_RE = re.compile(r'(?:\d\s*[/^]\s*\d|[A-Za-z]\s*[_^]\s*\d|[A-Za-z]\s*[=<>]|[∫∑√])')
PROTECT_RE = re.compile(r'\\(?:[A-Za-z]+(?:\{[^{}]*\})?)|\$[^$]+\$|\b(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|dA|dB|dV|vec|frac)\b|\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|CH3MgBr|LiAlH4|NaBH4|ATP|DNA|RNA)\b|\b\d+(?:\.\d+)?\s*(?:N|J|W|Pa|V|A|Hz|kg|g|mg|m|cm|mm|s|min|mol|K|°C|°)\b|[∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω]|https?://\S+', re.I)

@st.cache_data(show_spinner=False)
def pdf_state(data): return hashlib.sha256(data).hexdigest()

@st.cache_data(show_spinner=False)
def pages_of(data):
    pdf=fitz.open(stream=data,filetype="pdf"); out=[]
    for pno,p in enumerate(pdf):
        out.append({"page":pno,"width":p.rect.width,"height":p.rect.height,"words":p.get_text("words",sort=True),"blocks":p.get_text("blocks",sort=True)})
    return out

def line_groups(page):
    groups=[]
    for w in page["words"]:
        x0,y0,x1,y1,txt=w[:5]
        g=None
        for cand in groups:
            tol=max(2.0,min(4.0,(y1-y0)*0.45))
            if abs(cand["y"]-y0)<=tol: g=cand; break
        if g is None: groups.append({"y":y0,"ws":[w]})
        else: g["ws"].append(w); g["y"]=(g["y"]+y0)/2
    out=[]
    for g in sorted(groups,key=lambda z:z["y"]):
        ws=sorted(g["ws"],key=lambda w:(w[0],w[1]))
        out.append({"x0":min(w[0] for w in ws),"x1":max(w[2] for w in ws),"y0":min(w[1] for w in ws),"y1":max(w[3] for w in ws),"text":" ".join(w[4] for w in ws)})
    return out

def clean(s):
    return re.sub(r'\s+',' ',re.sub(r'\s+([,.;:!?])',r'\1',s).strip())

def is_noise(line):
    s=clean(line)
    if not s or FOOTER_RE.fullmatch(s) or PAGE_RE.fullmatch(s): return True
    u=s.upper()
    if any(u==h or u.startswith(h+" ") for h in HEADER_WORDS): return True
    if re.fullmatch(r'www\.\S+',s,re.I): return True
    return False

def detect_starts(pages):
    candidates=[]
    for p in pages:
        for ln in line_groups(p):
            s=clean(ln["text"]); m=QRE.match(s)
            if m and ln["x0"]<=p["width"]*.22:
                n=int(m.group(1))
                if n<=300: candidates.append({**ln,"page":p["page"],"num":n})
    runs=[]
    for x in candidates:
        if not runs or x["num"]!=runs[-1][-1]["num"]+1: runs.append([x])
        else: runs[-1].append(x)
    long=[r for r in runs if len(r)>=3]
    return max(long,key=len) if long else []

def build_records(pages):
    starts=detect_starts(pages); records=[]
    for i,s in enumerate(starts):
        e=starts[i+1] if i+1<len(starts) else None; seg=[]; p=s["page"]
        while p<len(pages):
            pg=pages[p]; ya=s["y0"]-2 if p==s["page"] else 14; yb=e["y0"]-4 if e and p==e["page"] else pg["height"]-18
            for ln in line_groups(pg):
                t=clean(ln["text"])
                if ya<ln["y0"]<yb and ln["x0"]<pg["width"]*.55 and (t.upper().startswith(HEADER_WORDS) or SECTION_RE.match(t)): yb=min(yb,ln["y0"]-3)
            if yb>ya+4: seg.append((p,ya,yb))
            if e and p>=e["page"]: break
            p+=1
        records.append({"num":s["num"],"segments":seg})
    return records

def formulaish(t): return bool(FORMULA_CHARS.intersection(t) or TECH_RE.search(t) or CHEM_RE.search(t) or MATH_RE.search(t))

@st.cache_data(show_spinner=False)
def translate_line(text):
    if not text.strip(): return text
    saved=[]
    def repl(m):
        key=f"PCXFORM{len(saved)}X"; saved.append(m.group(0)); return key
    protected=PROTECT_RE.sub(repl,text)
    try: out=GoogleTranslator(source="en",target="hi").translate(protected) or protected
    except Exception: out=protected
    for i,v in enumerate(saved): out=out.replace(f"PCXFORM{i}X",v)
    return out

def crop(data,pno,rect,scale=2.2):
    pdf=fitz.open(stream=data,filetype="pdf"); pg=pdf[pno]
    r=fitz.Rect(max(0,rect.x0-4),max(0,rect.y0-4),min(pg.rect.width,rect.x1+4),min(pg.rect.height,rect.y1+4))
    return pg.get_pixmap(matrix=fitz.Matrix(scale,scale),clip=r,alpha=False).tobytes("png")

def question_lines(pages,r):
    out=[]
    for pno,ya,yb in r["segments"]:
        for ln in line_groups(pages[pno]):
            if ya-1<=ln["y0"]<=yb+1:
                t=clean(ln["text"])
                if not is_noise(t): out.append((pno,ln))
    return out

def add_picture(cell,img,width=3.65):
    p=cell.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(2); p.add_run().add_picture(io.BytesIO(img),width=Inches(width))

def add_run(par,text,hindi=False,bold=False):
    r=par.add_run(text); r.font.name="Nirmala UI" if hindi else "Calibri"; r.font.size=Pt(9.5); r.bold=bold; r._element.rPr.rFonts.set(qn("w:eastAsia"),"Nirmala UI" if hindi else "Calibri")

def margins(cell):
    tc=cell._tc.get_or_add_tcPr(); m=OxmlElement("w:tcMar")
    for s,v in [("top",90),("bottom",90),("start",110),("end",110)]:
        x=OxmlElement("w:"+s); x.set(qn("w:w"),str(v)); x.set(qn("w:type"),"dxa"); m.append(x)
    tc.append(m)

def shade(cell,fill="E8EEF8"):
    tc=cell._tc.get_or_add_tcPr(); x=OxmlElement("w:shd"); x.set(qn("w:fill"),fill); tc.append(x)

def add_question_content(cell,data,pages,r,hindi=False):
    lines=question_lines(pages,r)
    p=cell.paragraphs[0]; add_run(p,f'{r["num"]}.',hindi,True)
    if not lines: return
    i=0
    while i<len(lines):
        pno,ln=lines[i]; text=clean(ln["text"])
        if text.startswith(f'{r["num"]}.'): text=text[len(str(r["num"]))+1:].strip()
        if formulaish(text):
            j=i+1; bottom=ln["y1"]; x0=ln["x0"]; x1=ln["x1"]
            while j<len(lines) and lines[j][0]==pno and lines[j][1]["y0"]<=bottom+9 and not QRE.match(clean(lines[j][1]["text"])):
                x0=min(x0,lines[j][1]["x0"]); x1=max(x1,lines[j][1]["x1"]); bottom=max(bottom,lines[j][1]["y1"]); j+=1
            rect=fitz.Rect(max(0,x0-5),max(0,ln["y0"]-4),min(pages[pno]["width"],x1+5),min(pages[pno]["height"],bottom+5))
            add_picture(cell,crop(data,pno,rect),3.65); i=j; continue
        p=cell.add_paragraph(); p.paragraph_format.space_after=Pt(2); add_run(p,translate_line(text) if hindi else text,hindi); i+=1
    seen=set(); pdf=fitz.open(stream=data,filetype="pdf")
    for pno,ya,yb in r["segments"]:
        pg=pdf[pno]
        for im in pg.get_images(full=True):
            for rr in pg.get_image_rects(im[0]):
                if rr.width*rr.height<500 or rr.y1<ya or rr.y0>yb or (rr.width>pg.rect.width*.85 and rr.height<30): continue
                key=(pno,round(rr.x0),round(rr.y0),round(rr.x1),round(rr.y1))
                if key in seen: continue
                seen.add(key); add_picture(cell,crop(data,pno,rr),3.65)

def docx_make(data,records,pages,progress):
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.30); sec.bottom_margin=Inches(.30); sec.left_margin=Inches(.28); sec.right_margin=Inches(.28)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(p,"BILINGUAL QUESTION PAPER",False,True)
    t=doc.add_table(rows=1,cols=2); t.autofit=False; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(["English","हिन्दी अनुवाद"]):
        c=t.rows[0].cells[i]; c.width=Inches(3.85); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; margins(c); shade(c); pp=c.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_run(pp,h,i==1,True)
    for i,r in enumerate(records):
        cells=t.add_row().cells
        for c in cells: c.width=Inches(3.85); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP; margins(c)
        add_question_content(cells[0],data,pages,r,False); add_question_content(cells[1],data,pages,r,True); progress.progress((i+1)/len(records))
    b=io.BytesIO(); doc.save(b); return b.getvalue()

st.markdown('<div class="upload"><h3 style="text-align:center">Create your bilingual question paper</h3><p style="text-align:center;color:#64748b">Upload an English PDF and generate a clean English + Hindi Word document.</p>',unsafe_allow_html=True)
up=st.file_uploader("Upload English Question Paper (PDF)",type=["pdf"])
st.markdown("</div>",unsafe_allow_html=True)
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
else: st.markdown('<div style="text-align:center;padding:35px;color:#64748b">Ready when you are — upload a PDF to begin.</div>',unsafe_allow_html=True)
