from pathlib import Path

p = Path('app.py')
s = p.read_text(encoding='utf-8')
start = s.index('def figure_assets(data):')
end = s.index('@st.cache_data(show_spinner=False)\ndef assets_cached', start)

new_func = r'''def _marker_centers(markers):
    return {k:((v[1]+v[3])/2,(v[2]+v[4])/2) for k,v in markers.items()}

def _crop_image(page,rect,pad=2,scale=2):
    r=fitz.Rect(max(page.rect.x0,rect.x0-pad),max(page.rect.y0,rect.y0-pad),min(page.rect.x1,rect.x1+pad),min(page.rect.y1,rect.y1+pad))
    return page.get_pixmap(matrix=fitz.Matrix(scale,scale),clip=r,alpha=False).tobytes('png')

def _split_composite_option(page,rect,markers):
    centers=_marker_centers(markers)
    if len(centers)<3 or rect.width<page.rect.width*.45:return {}
    inside=[k for k,(x,y) in centers.items() if rect.x0-12<=x<=rect.x1+12 and rect.y0-3<=y<=rect.y1+3]
    if len(inside)<3:return {}
    keys=sorted(inside,key=lambda k:centers[k][0])
    if keys != sorted(centers):return {}
    xs=[centers[k][0] for k in keys]
    bounds=[rect.x0]+[(xs[i]+xs[i+1])/2 for i in range(len(xs)-1)]+[rect.x1]
    out={}
    for k,x0,x1 in zip(keys,bounds,bounds[1:]):
        if x1-x0<18:continue
        out[k]=_crop_image(page,fitz.Rect(x0,rect.y0,x1,rect.y1),pad=1)
    return out if len(out)>=3 else {}

def _assign_option_images(candidates,markers):
    centers=_marker_centers(markers);out={};used=set()
    for rect,img in sorted(candidates,key=lambda z:(z[0].y0,z[0].x0)):
        possible=[]
        for k,(x,y) in centers.items():
            if k in used:continue
            inside=(rect.x0-8<=x<=rect.x1+8 and rect.y0-4<=y<=rect.y1+4)
            if inside:
                dx=max(0,rect.x0-x,x-rect.x1);dy=max(0,rect.y0-y,y-rect.y1)
                possible.append((dx+dy,k))
        if possible:
            _,k=min(possible);out[k]=img;used.add(k)
    return out

def figure_assets(data):
    p=fitz.open(stream=data,filetype='pdf');assets={}
    for page in p:
        qs=_q_ranges(page);imgs=_small_images(page)
        for idx,(q,y0,y1) in enumerate(qs):
            next_y=qs[idx+1][1] if idx+1<len(qs) else page.rect.y1
            markers=_option_markers(page,y0,next_y)
            candidates=[]
            for r,x in imgs:
                if r.y1<=y1 or r.y0>=next_y or r.width>=page.rect.width*.95:continue
                candidates.append((r,_crop_image(page,r,pad=2,scale=2)))
            option_imgs={};figure_imgs=[]
            # Composite option panels are common in graph/chemical-structure questions.
            for r,img in candidates:
                split=_split_composite_option(page,r,markers)
                if split:
                    option_imgs.update({k:(r,v) for k,v in split.items()})
                    continue
                # A normal option image overlaps its option marker. Body diagrams usually
                # finish before the option-marker row, so they remain question figures.
                assigned=_assign_option_images([(r,img)],markers)
                if assigned:
                    for k,v in assigned.items():option_imgs[k]=(r,v)
                else:
                    figure_imgs.append((r,img))
            if option_imgs:
                assets.setdefault(q,{})['options']={k:v[1] for k,v in sorted(option_imgs.items()) if 1<=k<=4}
            if figure_imgs:
                r,img=max(figure_imgs,key=lambda z:z[0].width*z[0].height)
                assets.setdefault(q,{})['figure']=img
    return assets

'''
s = s[:start] + new_func + s[end:]

old = "Transcribe this exam-paper figure/option into precise editable plain text. Do not omit symbols. For a chemical structure, give a compact unambiguous description or SMILES plus visible substituents. For a graph, state axes, curve direction and labels. For a circuit, state components and connections. Never invent unreadable values; use [unclear] instead."
new = "Transcribe this exam-paper visual into precise editable text for a question paper. Preserve every visible symbol, charge, subscript, superscript, bond, arrow, label, axis, scale and numerical value. If it is a chemical structure, describe the structure left-to-right with bond types and substituents; include a SMILES form when unambiguous, otherwise use [unclear] for unreadable parts. If it is a graph, state x/y axes, labels, intercepts, slope/curve shape, marked points and values. If it is a circuit, state every component and its connections. If it is a mathematical diagram, transcribe equations and labels exactly. Do not guess missing information. Return only the transcription, with no commentary."
if old not in s:
    raise SystemExit('vision prompt target not found')
s = s.replace(old,new,1)
p.write_text(s,encoding='utf-8')
print('Phase 3 patch applied')
