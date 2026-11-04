"""Build editable technical bids using the user's approved main formatting."""
import argparse
import json
from pathlib import Path
from docx import Document
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips, Mm, RGBColor

ROOT = Path(__file__).resolve().parents[1]
PROFILE = json.loads((ROOT / 'assets/hospital-style.json').read_text(encoding='utf-8'))


def element(tag, **attrs):
    e = OxmlElement('w:' + tag)
    for k, v in attrs.items():
        e.set(qn('w:' + k), str(v))
    return e


def font(style, east_asia, latin, size, bold=False):
    style.font.name = latin
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    style.font.underline = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    rp = style.element.get_or_add_rPr()
    rf = rp.find(qn('w:rFonts'))
    if rf is None:
        rf = element('rFonts'); rp.insert(0, rf)
    for attr in list(rf.attrib):
        if attr.split('}')[-1].lower().endswith('theme'):
            del rf.attrib[attr]
    rf.set(qn('w:eastAsia'), east_asia)
    for name, value in [('bCs', '1' if bold else '0'), ('iCs', '0')]:
        old = rp.find(qn('w:' + name))
        if old is not None: rp.remove(old)
        rp.append(element(name, val=value))


def no_grid(style):
    pp = style.element.get_or_add_pPr()
    old = pp.find(qn('w:snapToGrid'))
    if old is not None:
        pp.remove(old)
    pp.append(element('snapToGrid', val='0'))


def layout(section, landscape=False):
    p = PROFILE['landscape' if landscape else 'portrait']
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width = Twips(p['width_twips'])
    section.page_height = Twips(p['height_twips'])
    for k in ['top', 'bottom', 'left', 'right']:
        setattr(section, k + '_margin', Twips(p[k + '_twips']))
    section.header_distance = Twips(PROFILE['header_twips'])
    section.footer_distance = Twips(PROFILE['footer_twips'])


def field(paragraph, code):
    r = paragraph.add_run()
    r._r.append(element('fldChar', fldCharType='begin'))
    instr = element('instrText'); instr.text = ' ' + code + ' '
    instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r._r.append(instr)
    r._r.append(element('fldChar', fldCharType='separate'))
    t = element('t'); t.text = '1'; r._r.append(t)
    r._r.append(element('fldChar', fldCharType='end'))


def new_document():
    doc = Document()
    for style in doc.styles:
        for border in list(style.element.iter(qn('w:pBdr'))):
            border.getparent().remove(border)
    b = PROFILE['body']
    s = doc.styles['Normal']
    font(s, b['east_asia'], b['latin'], b['size_pt'])
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Twips(b['first_line_twips'])
    pf.line_spacing = b['line_spacing']
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    no_grid(s)
    nr = doc.part.numbering_part.element
    aids = [int(e.get(qn('w:abstractNumId'))) for e in nr.findall(qn('w:abstractNum'))]
    nids = [int(e.get(qn('w:numId'))) for e in nr.findall(qn('w:num'))]
    aid = max(aids + [0]) + 1; nid = max(nids + [0]) + 1
    abstract = element('abstractNum', abstractNumId=aid)
    abstract.append(element('multiLevelType', val='multilevel'))
    for h in PROFILE['headings']:
        level = h['level']; st = doc.styles['Heading ' + str(level)]
        font(st, h['font'], h['font'], h['size_pt'], True)
        pf = st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = b['line_spacing']
        pf.space_before = Pt(h['before_pt']); pf.space_after = Pt(h['after_pt'])
        pf.left_indent = Twips(h['indent_twips'])
        pf.first_line_indent = Twips(-h['indent_twips'])
        pf.keep_with_next = True; pf.keep_together = True
        pf.page_break_before = False
        no_grid(st)
        np = st.element.get_or_add_pPr().get_or_add_numPr()
        np.get_or_add_ilvl().val = level - 1
        np.get_or_add_numId().val = nid
        lv = element('lvl', ilvl=level - 1)
        lv.append(element('start', val=1)); lv.append(element('numFmt', val='decimal'))
        lv.append(element('pStyle', val=st.style_id))
        lv.append(element('lvlText', val='.'.join('%' + str(i) for i in range(1, level + 1)) + '.'))
        lv.append(element('suff', val='nothing'))
        lv.append(element('lvlJc', val='left'))
        pp = element('pPr'); pp.append(element('ind', left=h['indent_twips'], hanging=h['indent_twips']))
        lv.append(pp); abstract.append(lv)
    nr.append(abstract)
    num = element('num', numId=nid); num.append(element('abstractNumId', val=aid)); nr.append(num)
    for name, size in [('Caption', 10.5), ('Footer', 9), ('Title', 22)]:
        st = doc.styles[name]; font(st, '宋体', '宋体', size, name == 'Title')
        st.paragraph_format.first_line_indent = Pt(0)
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        no_grid(st)
    for section in doc.sections:
        layout(section)
    footer = doc.sections[0].footer.paragraphs[0]
    footer.style = doc.styles['Footer']
    footer.add_run('第'); field(footer, 'PAGE'); footer.add_run('页共'); field(footer, 'NUMPAGES'); footer.add_run('页')
    doc.core_properties.title = ''
    doc.core_properties.author = ''
    doc.core_properties.last_modified_by = ''
    doc.settings.element.append(element('updateFields', val='true'))
    return doc


def add_heading(doc, text, level=1):
    if not 1 <= level <= 7:
        raise ValueError('Heading level must be 1–7')
    return doc.add_paragraph(text, style='Heading ' + str(level))


def add_body(doc, text):
    return doc.add_paragraph(text, style='Normal')


def set_orientation(doc, landscape=False):
    section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    layout(section, landscape)
    return section


def add_table(doc, headers, rows, widths_mm=None):
    count = len(headers)
    if count == 0 or any(len(row) != count for row in rows):
        raise ValueError('Every row must match the header column count')
    section = doc.sections[-1]
    usable = section.page_width - section.left_margin - section.right_margin
    if widths_mm is None:
        widths_mm = [usable / 36000 / count] * count
    if len(widths_mm) != count or sum(Mm(w) for w in widths_mm) > usable + Mm(0.1):
        raise ValueError('Table columns exceed the section text width')
    t = doc.add_table(rows=1, cols=count); t.autofit = False
    for col, w in zip(t.columns, widths_mm): col.width = Mm(w)
    data = [headers] + list(rows)
    for index, values in enumerate(data):
        row = t.rows[0] if index == 0 else t.add_row()
        if index == 0:
            row._tr.get_or_add_trPr().append(element('tblHeader'))
        for cell, value, width in zip(row.cells, values, widths_mm):
            cell.width = Mm(width); cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.text = str(value)
            tcpr = cell._tc.get_or_add_tcPr()
            borders = element('tcBorders')
            for side in ['top', 'left', 'bottom', 'right']:
                borders.append(element(side, val='single', sz=4, color='000000'))
            tcpr.append(borders)
            margins = element('tcMar')
            for side in ['top', 'left', 'bottom', 'right']:
                margins.append(element(side, w=80, type='dxa'))
            tcpr.append(margins)
            for p in cell.paragraphs:
                pf = p.paragraph_format
                pf.first_line_indent = Pt(0)
                pf.space_before = Pt(0); pf.space_after = Pt(0)
                pf.line_spacing = 1.5
                pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if index == 0 or len(str(value)) < 12 else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.bold = index == 0
                    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    return t


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--out', required=True)
    p.add_argument('--demo', action='store_true')
    args = p.parse_args()
    doc = new_document()
    if args.demo:
        doc.add_paragraph('技术服务部分', style='Title')
        add_heading(doc, '项目理解及需求分析')
        add_heading(doc, '需求分析', 2)
        add_heading(doc, '工作范围与成果用途', 3)
        add_body(doc, '本示例仅用于检查技术文件格式。编制正式方案时，应依据当前采购需求确认服务范围、工作期限和成果要求，并建立需求与设计内容的对应关系。')
        for lv, text in [(4,'需求核验'),(5,'资料登记'),(6,'记录检查'),(7,'问题闭环')]: add_heading(doc, text, lv)
        add_body(doc, '核验记录应明确问题来源、处理责任、计划节点和复核结论，形成可追溯的工作记录。')
        set_orientation(doc, True)
        add_heading(doc, '任务与成果对应', 2)
        add_table(doc, ['工作项','责任角色','形成成果','检查方式'], [['资料收集','项目负责人','输入资料清单','完整性核对'],['需求分析','专业负责人','需求分析报告','与采购需求逐项对应']], [35,35,90,70])
        set_orientation(doc, False)
        add_heading(doc, '质量保障措施')
        add_body(doc, '正式交付前完成内容、格式及版本检查，保留审核意见和整改记录。')
    else:
        doc.add_paragraph()
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.out)


if __name__ == '__main__': main()
