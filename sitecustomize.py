"""PaperCraft AI Phase 7/8: stable Word flow and final DOCX QA gate.

Phase 7 hardens the existing bilingual table layout for print/edit workflows.
Phase 8 validates the actual generated DOCX at the final download boundary, so
layout/generation regressions cannot silently become downloadable output.
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
    """Give the bilingual table deterministic print geometry."""
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
    paragraphs = list(cell.paragraphs)
    nonempty = [p for p in paragraphs if (p.text or '').strip()]
    for idx, paragraph in enumerate(nonempty):
        _paragraph_quality(paragraph, keep_with_next=(idx < len(nonempty) - 1))


def _harden_document(doc):
    for section in doc.sections:
        _master_page(section)
    section = doc.sections[0]
    for table in doc.tables:
        _fixed_table(table, section)
        if table.rows:
            _repeat_header(table.rows[0])
        for row in table.rows:
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


# Phase 8: validate the actual generated DOCX at the existing download boundary.
# This preserves the premium UI and requires no changes to the app's download UI.
try:
    import streamlit as _st
    _ORIGINAL_DOWNLOAD = _st.download_button

    def _qa_download_button(*args, **kwargs):
        data = kwargs.get('data')
        if data is None and len(args) >= 2:
            data = args[1]
        filename = kwargs.get('file_name')
        if filename is None and len(args) >= 3:
            filename = args[2]
        is_docx = isinstance(filename, str) and filename.lower().endswith('.docx')
        if is_docx and isinstance(data, (bytes, bytearray)):
            try:
                from final_qa import final_docx_qa
                errors, warnings = final_docx_qa(bytes(data))
                if errors:
                    _st.error(f'Final Word QA failed: {len(errors)} issue(s). Download is blocked to prevent a corrupted paper.')
                    for issue in errors[:30]:
                        _st.write('• ' + issue)
                    if len(errors) > 30:
                        _st.write(f'• …and {len(errors)-30} more issue(s).')
                    return False
                if warnings:
                    _st.warning(f'Final Word QA passed with {len(warnings)} warning(s). Review recommended before printing.')
                    for warning in warnings[:10]:
                        _st.write('• ' + warning)
                    if len(warnings) > 10:
                        _st.write(f'• …and {len(warnings)-10} more warning(s).')
                else:
                    _st.success('Final Word QA passed: structure, bilingual cells, options and protected values are intact.')
            except Exception as exc:
                _st.error(f'Final Word QA could not be completed: {exc}. Download is blocked for safety.')
                return False
        return _ORIGINAL_DOWNLOAD(*args, **kwargs)

    if _st.download_button is not _qa_download_button:
        _st.download_button = _qa_download_button
except Exception:
    # Never prevent the application from starting if Streamlit internals change.
    pass
