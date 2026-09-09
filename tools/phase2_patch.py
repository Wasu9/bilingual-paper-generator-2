from pathlib import Path
import re

p = Path('app.py')
s = p.read_text()

protect = r'''PROTECT=re.compile(r''' + "'''" + r'''\\(?:[A-Za-z]+(?:\{[^{}]*\})?)|\$[^$]+\$|\b(?:sin|cos|tan|cot|sec|cosec|log|ln|lim|exp|dx|dy|dt|dA|dB|dV|vec|frac)\b|\b(?:H2O|CO2|O2|N2|H2|NaCl|HCl|H2SO4|HNO3|NH3|CH4|C2H5OH|KMnO4|K2Cr2O7|NaOH|CaCO3|CaO|CH3MgBr|LiAlH4|NaBH4|ATP|DNA|mRNA|tRNA|eV|keV|MeV|GeV|Hz|kHz|MHz|GHz|Pa|kPa|MPa|mol|kg|mg|μg|cm|mm|nm|km|ms|μs)\b|\b[A-Za-zα-ωΑ-Ω]+(?:\^[-+]?\d+|[₀₁₂₃₄₅₆₇₈₉])+\b|\b[A-Za-zα-ωΑ-Ω0-9]+\s*[=≈≠≤≥<>±∝→←↔⇌]\s*[A-Za-zα-ωΑ-Ω0-9²³⁴⁵⁻⁺⁰¹²³⁴⁵⁶⁷⁸⁹./()+×÷√-]+(?:\s+[A-Za-zα-ωΑ-Ω0-9²³⁴⁵⁻⁺⁰¹²³⁴⁵⁶⁷⁸⁹./()+×÷√-]+)*|[∫∮∑√∞±×÷≈≠≤≥→←↔⇌∆∂∇θπλμΩαβγδφψω²³⁴⁵⁻⁺⁰¹²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉]',re.I)''' + "'''"

s, n = re.subn(r'^PROTECT=.*$', lambda _m: protect, s, count=1, flags=re.M)
if n != 1:
    raise SystemExit('PROTECT definition not found')

new_math = r'''def math(x):
    m={'𝑥':'x','𝑦':'y','𝑧':'z','𝑛':'n','𝑎':'a','𝑏':'b','𝑚':'m','𝑝':'p','𝑞':'q','𝑟':'r','𝑡':'t','𝑓':'f','𝑒':'e','𝑑':'d','𝑐':'c','𝛼':'α','𝛽':'β','𝛾':'γ','𝛿':'δ','𝜃':'θ','𝜋':'π','𝜆':'λ','𝜇':'μ','𝜔':'ω','𝛺':'Ω'}
    for a,b in m.items(): x=x.replace(a,b)
    x=x.replace('∗','×').replace('−','-').replace('–','-').replace('—','-')
    x=re.sub(r'\s+',' ',x).strip()
    x=re.sub(r'\s*([=+\-×÷≤≥≠<>∝→←↔⇌])\s*',r' \1 ',x)
    return re.sub(r'\s+',' ',x).strip()
'''
s = re.sub(r'def math\(x\):.*?\n(?=def fmt\()', new_math + '\n', s, count=1, flags=re.S)

new_run = r'''def _add_typed_run(p,text,h=False,b=False):
    sub='₀₁₂₃₄₅₆₇₈₉'; sup='⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾'
    pattern=re.compile(r'(\^[+-]?\d+|[₀₁₂₃₄₅₆₇₈₉]+|[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾]+)')
    def style(r):
        r.font.name='Nirmala UI' if h else 'Calibri'; r.font.size=Pt(9.5); r.bold=b
        r._element.rPr.rFonts.set(qn('w:eastAsia'),'Nirmala UI' if h else 'Calibri')
    pos=0
    for m in pattern.finditer(text):
        if m.start()>pos:
            r=p.add_run(text[pos:m.start()]); style(r)
        token=m.group(0); r=p.add_run(token[1:] if token.startswith('^') else token); style(r)
        if token.startswith('^') or all(c in sup for c in token): r.font.superscript=True
        else: r.font.subscript=True
        pos=m.end()
    if pos<len(text):
        r=p.add_run(text[pos:]); style(r)

def run(p,t,h=False,b=False):
    p.paragraph_format.space_after=Pt(2)
    _add_typed_run(p,t,h,b)
'''
s = re.sub(r'def run\(p,t,h=False,b=False\):.*?\n(?=def subject_map\()', new_run + '\n', s, count=1, flags=re.S)

p.write_text(s)
print('Phase 2 patch applied')
