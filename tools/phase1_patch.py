from pathlib import Path
p=Path('app.py'); s=p.read_text()
old="PAGE=re.compile(r'^(?:I|II)?\\s*PUC.*(?:MAINS|NEET).*page\\s*\\d+$',re.I)"
new="""PAGE=re.compile(r'^(?:VERSION\\s*[- ]?\\s*\\d+\\s+)?I\\s*[- ]*PUC.*?Page\\s*\\d+\\s+of\\s+\\d+$',re.I)
QUESTION_START=re.compile(r'^\\s*(\\d{1,3})(?:\\.(?!\\d)\\s*|\\s+(?=(?:What|Which|When|Where|Who|Why|How|A|An|The|If|During|Identify|Choose|Select|Calculate|Find|Determine|Consider|Given|For|In|On|According|Assertion|Reason|Match|Name|Our|Cells|Cell|Gas|Resistance|Eukaryotic|Mechanisms|Movement|Additional|Tidal|Normal|Inspiratory|Total|Prophase|Chromosomes|Meiosis|Does|Is|Are|Can|Let)\\b))',re.I|re.M)
PAGE_INLINE=re.compile(r'(?i)\\bPage\\s*\\d+\\s+of\\s+\\d+\\b|\\bVERSION\\s*[- ]?\\s*\\d+\\s*[/\\-]?\\s*I\\s*[- ]*PUC\\s*(?:[- ]*(?:NEET|JEE\\s*MAINS))?\\s*(?:[/\\-]?\\s*of\\s*\\d+)?')"""
if old not in s: raise SystemExit('parser anchor not found')
s=s.replace(old,new,1)
start=s.index('def blocks(ps):'); end=s.index('def parse(n,b):',start)
s=s[:start]+'''def _question_candidates(text):
    return [(m.start(),int(m.group(1))) for m in QUESTION_START.finditer(text) if 1 <= int(m.group(1)) <= 500]

def blocks(ps):
    t='\\n'.join(ps); candidates=_question_candidates(t); runs=[]
    for i,(pos,num) in enumerate(candidates):
        if num != 1: continue
        run=[(pos,num)]; last=num
        for pos2,num2 in candidates[i+1:]:
            if num2 == last+1:
                run.append((pos2,num2)); last=num2; continue
            if last >= 10 and 1 <= num2 <= 4: continue
            break
        if len(run) >= 3:
            first_end=run[1][0] if len(run)>1 else len(t)
            first_block=t[pos:first_end]
            option_count=len(re.findall(r'\\((?:1|2|3|4)\\)',first_block))
            question_words=len(re.findall(r'\\b(?:what|which|when|where|who|why|how|find|calculate|given|consider|identify|determine|select|choose|the|during|assertion|reason)\\b',first_block,re.I))
            runs.append(((len(run),option_count,question_words,len(first_block)),run))
    if not runs: return []
    best=max(runs,key=lambda z:z[0])[1]
    return [(num,t[pos:best[i+1][0] if i+1<len(best) else len(t)]) for i,(pos,num) in enumerate(best)]

'''+s[end:]
start=s.index('def parse(n,b):'); end=s.index('def records(ps):',start)
s=s[:start]+'''def parse(n,b):
    ls=[]
    for x in b.splitlines():
        x=norm(x); x=PAGE_INLINE.sub(' ',x)
        x=re.sub(r'\\s*/\\s*$', '', x)
        x=re.sub(r'\\s*\\(?/?\\s*correct\\.\\s*\\)?\\s*$', '', x, flags=re.I)
        x=norm(x)
        if not x or PAGE.fullmatch(x) or x.upper() in HEAD or SEC.match(x) or x.lower().startswith(('this section contains','marking scheme:')): continue
        ls.append(x)
    if ls:
        qprefix=re.compile(rf'^\\s*{n}(?:\\.(?!\\d))?\\s*',re.I)
        ls[0]=qprefix.sub('',ls[0],1).strip()
    parts={1:[],2:[],3:[],4:[]}; stem=[]; cur=None
    for x in ls:
        markers=list(re.finditer(r'\\((1|2|3|4)\\)',x))
        if markers:
            prefix=x[:markers[0].start()].strip()
            if prefix: (parts[cur] if cur is not None else stem).append(prefix)
            for j,m in enumerate(markers):
                k=int(m.group(1)); v=x[m.end():markers[j+1].start() if j+1<len(markers) else len(x)].strip()
                if v: parts[k].append(v)
                cur=k
        elif cur is None: stem.append(x)
        else: parts[cur].append(x)
    return {'num':n,'stem':fmt(stem),'options':[fmt(parts[k]) for k in (1,2,3,4)]}

def validate_records(rs):
    nums=[q['num'] for q in rs]; issues=[]
    if not nums: return ['No numbered questions detected.']
    if nums[0] != 1: issues.append(f'Question sequence starts at {nums[0]}, not 1.')
    missing=[n for n in range(1,max(nums)+1) if n not in nums]
    duplicates=sorted({n for n in nums if nums.count(n)>1})
    if missing: issues.append('Missing question number(s): '+', '.join(map(str,missing[:20])))
    if duplicates: issues.append('Duplicate question number(s): '+', '.join(map(str,duplicates[:20])))
    if nums != sorted(nums): issues.append('Question numbers are out of order.')
    return issues

'''+s[end:]
old_ui="""    if not rs:st.error('No reliable numbered questions found. Scanned/image-only PDFs need OCR.')
    else:
        blank=sum(not o for q in rs for o in q['options']);visual_blank=sum(1 for q in rs for k in range(1,5) if not q['options'][k-1] and assets.get(q['num'],{}).get('options',{}).get(k))
        st.success(f'{len(rs)} questions detected. Options will be typed as (1)–(4); no option images.')"""
new_ui="""    structure_issues=validate_records(rs)
    if not rs:st.error('No reliable numbered questions found. Scanned/image-only PDFs need OCR.')
    elif structure_issues:
        st.error('Question extraction needs review before Word generation:')
        for issue in structure_issues: st.write('• '+issue)
        st.warning('Generation is stopped to prevent a corrupted paper.')
    else:
        blank=sum(not o for q in rs for o in q['options']);visual_blank=sum(1 for q in rs for k in range(1,5) if not q['options'][k-1] and assets.get(q['num'],{}).get('options',{}).get(k))
        st.success(f'{len(rs)} questions detected. Sequence verified: 1–{rs[-1]["num"]}. Options will be typed as (1)–(4); no option images.')"""
if old_ui not in s: raise SystemExit('UI anchor not found')
s=s.replace(old_ui,new_ui,1)
p.write_text(s)
