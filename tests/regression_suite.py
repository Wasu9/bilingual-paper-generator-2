"""Parser regression suite for PaperCraft AI.

Real PDF paths may be supplied on the command line. The suite imports only
parser functions/constants from app.py, so Streamlit UI is never started.
"""
from __future__ import annotations
import ast, re, sys
from pathlib import Path

def _load_parser_namespace():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    tree = ast.parse(app_path.read_text(encoding="utf-8"))
    wanted = {"HEAD","PAGE","QUESTION_START","PAGE_INLINE","SEC","LIG","PROTECT","norm","math","fmt","_question_candidates","blocks","parse","validate_records"}
    nodes=[]
    for node in tree.body:
        if isinstance(node,ast.Assign):
            names={t.id for t in node.targets if isinstance(t,ast.Name)}
            if names & wanted:nodes.append(node)
        elif isinstance(node,ast.FunctionDef) and node.name in wanted:nodes.append(node)
    ns={"re":re}
    exec(compile(ast.Module(body=nodes,type_ignores=[]),str(app_path),"exec"),ns)
    return ns

def _assert(condition,message):
    if not condition:raise AssertionError(message)

def _synthetic_cases(ns):
    cases=[]
    text="\n".join([
        "IMPORTANT INSTRUCTIONS","1. Use of calculator is prohibited.","2. Candidates must follow all rules.","3. Do not write on the question paper.","4. Report to the examination hall.",
        "1. A particle of mass m is at rest.","(1) F²t²/2m (2) F²t²/3m (3) Ft²/2m (4) Ft/2m",
        "2. A body has momentum p.","(1) p/m (2) p²/m (3) mp (4) 2p/m",
        "3. Which quantity is conserved?","(1) Energy (2) Mass (3) Momentum (4) All",
        "4. Calculate the value of x.","(1) 1 (2) 2 (3) 3 (4) 4"])
    rs=ns["blocks"]([text])
    _assert([n for n,_ in rs]==[1,2,3,4],"instruction-numbering selection failed")
    parsed=ns["parse"](1,rs[0][1]);_assert(len(parsed["options"])==4,"inline four-option parsing failed");_assert(parsed["options"][0].startswith("F"),"option (1) lost")
    cases.append("instruction-prefix protection")

    text="\n".join(["1. Match the following columns.","(1) A-I (2) B-II (3) C-III (4) D-IV","2. Which statement is correct?","(1) A (2) B (3) C (4) D","3. Find the value.","(1) 10 (2) 20 (3) 30 (4) 40","4. Choose the correct answer.","(1) P (2) Q (3) R (4) S","5. The next question.","(1) X (2) Y (3) Z (4) W"])
    rs=ns["blocks"]([text]);_assert([n for n,_ in rs]==[1,2,3,4,5],"internal option labels broke sequential parsing");cases.append("internal-option-label protection")

    malformed="154 During inspiration, the diaphragm contracts."
    candidates=ns["_question_candidates"](malformed)
    _assert(candidates and candidates[0][1]==154,"malformed Q154 candidate was not detected")
    lines=[f"{n}. Synthetic question {n}.\n(1) A (2) B (3) C (4) D" for n in range(1,154)]
    lines.append("154 During inspiration, the diaphragm contracts.\n(1) True (2) False (3) Both (4) None")
    lines.append("155. The next question.\n(1) A (2) B (3) C (4) D")
    rs=ns["blocks"](["\n".join(lines)])
    _assert([n for n,_ in rs]==list(range(1,156)),"malformed Q154 broke full sequential run")
    cases.append("malformed-question-number")
    return cases

def _real_pdf(ns,path:Path,expected_count=None):
    import fitz
    doc=fitz.open(path);pages=[p.get_text("text") for p in doc];rs=ns["blocks"](pages)
    issues=ns["validate_records"]([ns["parse"](n,b) for n,b in rs])
    if expected_count is not None:_assert(len(rs)==expected_count,f"{path.name}: expected {expected_count}, got {len(rs)}")
    _assert(not issues,f"{path.name}: parser validation failed: {issues}")
    _assert([n for n,_ in rs]==list(range(1,len(rs)+1)),f"{path.name}: sequence is not contiguous")
    return len(rs)

def main(argv):
    ns=_load_parser_namespace();synthetic=_synthetic_cases(ns);print(f"PASS synthetic cases: {len(synthetic)}")
    expected={
        "07-09-2026  I-PUC NEET Q.P V-1.pdf":180,
        "07-09-26 II PUC JEE MAINS BRANCH.pdf":75,
        "Test 05_12th_Eng_12-9-2026(1).pdf":180,
        "Test 05_12th_Eng_12-9-2026.pdf":180,
        "parser_12q.pdf":12,
        "parser_155q_malformed154.pdf":155,
    }
    for raw in argv:
        path=Path(raw);count=_real_pdf(ns,path,expected.get(path.name));print(f"PASS PDF fixture: {path.name} -> {count} questions")
    print("Phase 9/11 parser regression suite: PASS")

if __name__=="__main__":main(sys.argv[1:])
