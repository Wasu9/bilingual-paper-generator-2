"""Generate compact PDF fixtures that exercise production parser edge cases.

These are synthetic regression fixtures, not copies of copyrighted exam papers.
They intentionally mimic common PDF extraction layouts: instruction numbering,
inline options, internal option labels, page breaks, and malformed Q154 numbering.
"""
from __future__ import annotations
import sys
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def _draw(c, lines, width, height):
    y = height - 42
    for line in lines:
        c.drawString(42, y, line)
        y -= 15
        if y < 42:
            c.showPage()
            y = height - 42
    return y


def make_fixture(path: Path, count: int, malformed_154: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    c.setFont("Helvetica", 9)
    c.drawString(42, height - 28, "PAPERCRAFT AI REGRESSION FIXTURE")
    c.drawString(42, height - 40, "IMPORTANT INSTRUCTIONS")
    instructions = [
        "1. Use of calculator is prohibited.",
        "2. Candidates must follow all rules.",
        "3. Do not write on the question paper.",
        "4. Report to the examination hall.",
    ]
    y = _draw(c, instructions, width, height)
    for n in range(1, count + 1):
        if y < 105:
            c.showPage(); c.setFont("Helvetica", 9); y = height - 42
        if malformed_154 and n == 154:
            c.drawString(42, y, "154 During inspiration, the diaphragm contracts.")
        elif n == 1:
            c.drawString(42, y, "1. Match the following columns.")
        else:
            c.drawString(42, y, f"{n}. Which statement is correct for the given condition?")
        y -= 14
        if n == 1:
            opts = "(1) A-I    (2) B-II    (3) C-III    (4) D-IV"
        else:
            opts = "(1) Option one    (2) Option two    (3) Option three    (4) Option four"
        c.drawString(58, y, opts)
        y -= 22
    c.save()


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("tests/fixtures")
    make_fixture(out / "parser_12q.pdf", 12)
    make_fixture(out / "parser_155q_malformed154.pdf", 155, malformed_154=True)
    print(f"Generated PDF fixtures in {out}")


if __name__ == "__main__":
    main()
