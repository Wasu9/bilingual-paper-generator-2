"""PaperCraft AI Phase 7: Master Page + stable PageMaker-style Word flow.

This post-processor keeps the existing PaperCraft UI, parser, translation and visual
pipelines unchanged. It hardens the generated DOCX for print/edit workflows by adding
repeatable master-page header/footer geometry, stable two-column bilingual geometry,
non-splitting logical question rows, and pagination-safe paragraph rules.
"""
from docx import document as _document_module
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_ORIGINAL_SAVE = _document_module.Document.save


def _set_cell_margins(cell, top=55, start=65, bottom=55, end=65):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f'w:{side}'))
        if node is None:
            node = OxmlElement(f'w:{side}')
            tcMar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')


def _cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None:
        trPr.append(OxmlElement('w:cantSplit'))


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:tblHeader')) is None:
        trPr.append(OxmlElement('w:tblHeader'))


def _fixed_table(table, section):
    """Give the bilingual table deterministic print geometry.

    python-docx Length values are EMU while OOXML table widths are twips.
    Keeping this conversion explicit avoids the subtle 635x width bug that can
    otherwise make a generated paper wider than the printable page.
    """
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')
    table.autofit = False

    available_emu = int(section.page_width - section.left_margin - section.right_margin - section.gutter)
    available_twips = max(1, int(round(available_emu / 635.0)))
    half_emu = max(1, available_emu // 2)

    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(available_twips))
    tblW.set(qn('w:type'), 'dxa')

    # Explicitly mark this as a fixed-layout table so Word does not reflow the
    # English/Hindi columns differently after opening/editing the document.
    tblLook = tblPr.find(qn('w:tblLook'))
    if tblLook is None:
        tblLook = OxmlElement('w:tblLook')
        tblPr.append(tblLook)
    tblLook.set(qn('w:firstRow'), '1')
    tblLook.set(qn('w:noVBand'), '1')
    tblLook.set(qn('w:noHBand'), '1')

    for row in table.rows:
        for cell in row.cells:
            cell.width = half_emu
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(max(1, int(round(available_twips / 2)))))
            tcW.set(qn('w:type'), 'dxa')


def _paragraph_quality(paragraph, keep_with_next=False):
    pf = paragraph.paragraph_format
    pf.widow_control = True
    if keep_with_next:
        pf.keep_with_next = True
    if pf.space_after is None:
        pf.space_after = Pt(2)
    for r in paragraph.runs:
        if r.font.size is None:
            r.font.size = Pt(9.5)


def _page_field(paragraph):
    run = paragraph.add_run()
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    run._r.append(fld)


def _master_page(section):
    """Apply a repeatable master-page header/footer without changing body content."""
    section.header_distance = Pt(8)
    section.footer_distance = Pt(8)

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.text = ''
    r = hp.add_run('PaperCraft AI  •  Bilingual Question Paper')
    r.font.name = 'Calibri'
    r.font.size = Pt(8)
    r.bold = True

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.text = ''
    r = fp.add_run('PaperCraft AI  •  Page ')
    r.font.name = 'Calibri'
    r.font.size = Pt(8)
    _page_field(fp)
    for r in fp.runs:
        r.font.name = 'Calibri'
        r.font.size = Pt(8)


def _harden_question_cell(cell):
    """Keep the question stem attached to its following option paragraphs.

    The table row remains the primary logical unit, while keep-with-next prevents
    Word from creating avoidable orphaned stems when it repaginates an edited file.
    The final paragraph is intentionally left free so Word still has a legal place
    to break between questions.
    """
    paragraphs = list(cell.paragraphs)
    nonempty = [p for p in paragraphs if (p.text or '').strip()]
    for idx, paragraph in enumerate(nonempty):
        is_first = idx == 0
        _paragraph_quality(paragraph, keep_with_next=(is_first or idx < len(nonempty) - 1))


def _harden_document(doc):
    for section in doc.sections:
        _master_page(section)

    # PaperCraft currently uses a two-column bilingual table because English and
    # Hindi must remain aligned question-by-question. Phase 7 therefore hardens
    # the table as a deterministic threaded *logical* flow rather than converting
    # it to independent newspaper columns (which would destroy bilingual alignment).
    section = doc.sections[0]
    for table_index, table in enumerate(doc.tables):
        _fixed_table(table, section)
        if table.rows:
            _repeat_header(table.rows[0])
        for row_index, row in enumerate(table.rows):
            _cant_split(row)
            for cell in row.cells:
                _set_cell_margins(cell)
                _harden_question_cell(cell)
                for paragraph in cell.paragraphs:
                    if not (paragraph.text or '').strip():
                        paragraph.paragraph_format.space_after = Pt(0)

    for paragraph in doc.paragraphs:
        _paragraph_quality(paragraph)


def _patched_save(self, path_or_stream):
    _harden_document(self)
    return _ORIGINAL_SAVE(self, path_or_stream)


if _document_module.Document.save is not _patched_save:
    _document_module.Document.save = _patched_save
