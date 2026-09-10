from pathlib import Path
p=Path('app.py')
s=p.read_text(encoding='utf-8')
s2=s.replace('rect.x0-8<=x<=rect.x1+8 and rect.y0-4<=y<=rect.y1+4','rect.x0-30<=x<=rect.x1+30 and rect.y0-4<=y<=rect.y1+4')
s2=s2.replace('rect.x0-12<=x<=rect.x1+12 and rect.y0-3<=y<=rect.y1+3','rect.x0-30<=x<=rect.x1+30 and rect.y0-3<=y<=rect.y1+3')
if s2==s: raise SystemExit('tolerance targets not found')
p.write_text(s2,encoding='utf-8')
print('Phase 3 tolerance adjusted')
