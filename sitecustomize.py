"""PaperCraft AI Phase 7: Master Page + stable PageMaker-style Word flow.

This post-processor keeps the existing PaperCraft UI, parser, translation and visual
pipelines unchanged. It hardens the generated DOCX for print/edit workflows by adding
repeatable master-page header/footer geometry, stable two-column table geometry,
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
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')
    table.autofit = False

    available = section.page_width - section.left_margin - section.right_margin - section.gutter
    half_twips = max(1, int(available / 2))
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(half_twips * 2))
    tblW.set(qn('w:type'), 'dxa')
    for row in table.rows:
        for cell in row.cells:
            cell.width = available / 2


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


def _harden_document(doc):
    for section in doc.sections:
        section.header_distance = Pt(8)
        section.footer_distance = Pt(8)
        _master_page(section)

    for table in doc.tables:
        section = doc.sections[0]
        _fixed_table(table, section)
        if table.rows:
            _repeat_header(table.rows[0])
        for row in table.rows:
            _cant_split(row)
            for cell in row.cells:
                _set_cell_margins(cell)
                for paragraph in cell.paragraphs:
                    text = (paragraph.text or '').strip()
                    is_question = bool(text[:4].split('.')[0].isdigit() and '.' in text[:5])
                    is_subject = bool(text) and len(text) < 40 and not text.startswith(('(', '['))
                    _paragraph_quality(paragraph, keep_with_next=(is_question or is_subject))
                    if not text:
                        paragraph.paragraph_format.space_after = Pt(0)

    for paragraph in doc.paragraphs:
        _paragraph_quality(paragraph)


def _patched_save(self, path_or_stream):
    _harden_document(self)
    return _ORIGINAL_SAVE(self, path_or_stream)


if _document_module.Document.save is not _patched_save:
    _document_module.Document.save = _patched_save
