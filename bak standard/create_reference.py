"""
生成 Pandoc reference.docx 模板
对齐 BDAE ESA 报告现有样式
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── 品牌色 ──────────────────────────────────────────────────────────────────
BRAND_GREEN = RGBColor(0x1C, 0x7E, 0x4F)   # #1C7E4F BDAE 绿
DARK_TEXT   = RGBColor(0x26, 0x26, 0x26)    # #262626 正文近黑
GRAY_TEXT   = RGBColor(0x80, 0x80, 0x80)    # #808080 次要文字


def make_reference_docx(output_path='/app/reference.docx'):
    doc = Document()

    # ── 页面设置 A4，对齐原报告边距 ────────────────────────────────────────
    for section in doc.sections:
        section.page_height   = Cm(29.7)
        section.page_width    = Cm(21.0)
        section.top_margin    = Cm(2.3)
        section.bottom_margin = Cm(2.3)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    # ── Normal 正文 ─────────────────────────────────────────────────────────
    normal = doc.styles['Normal']
    normal.font.name  = 'Arial'
    normal.font.size  = Pt(11)
    normal.font.color.rgb = DARK_TEXT
    _set_east_asia_font(normal, '宋体')
    normal.paragraph_format.space_before = Pt(2.5)
    normal.paragraph_format.space_after  = Pt(2.5)
    normal.paragraph_format.line_spacing = 1.15

    # ── Heading 1：品牌绿，16pt，粗体 ──────────────────────────────────────
    h1 = _get_style(doc, 'Heading 1')
    h1.font.name  = 'Arial'
    h1.font.size  = Pt(16)
    h1.font.bold  = True
    h1.font.color.rgb = BRAND_GREEN
    _set_east_asia_font(h1, '黑体')
    h1.paragraph_format.space_before    = Pt(18)
    h1.paragraph_format.space_after     = Pt(6)
    h1.paragraph_format.keep_with_next  = True

    # ── Heading 2：深色，13pt，粗体 ─────────────────────────────────────────
    h2 = _get_style(doc, 'Heading 2')
    h2.font.name  = 'Arial'
    h2.font.size  = Pt(13)
    h2.font.bold  = True
    h2.font.color.rgb = DARK_TEXT
    _set_east_asia_font(h2, '黑体')
    h2.paragraph_format.space_before    = Pt(12)
    h2.paragraph_format.space_after     = Pt(4)
    h2.paragraph_format.keep_with_next  = True

    # ── Heading 3：深色，11pt，粗体 ─────────────────────────────────────────
    h3 = _get_style(doc, 'Heading 3')
    h3.font.name  = 'Arial'
    h3.font.size  = Pt(11)
    h3.font.bold  = True
    h3.font.color.rgb = DARK_TEXT
    _set_east_asia_font(h3, '黑体')
    h3.paragraph_format.space_before    = Pt(8)
    h3.paragraph_format.space_after     = Pt(3)
    h3.paragraph_format.keep_with_next  = True

    # 占位段落
    doc.add_paragraph('BDAE ESA Report Template', style='Normal')

    doc.save(output_path)
    print(f'reference.docx saved → {output_path}')


def _get_style(doc, name):
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, 1)


def _set_east_asia_font(style, font_name):
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'),       font_name)


if __name__ == '__main__':
    make_reference_docx()
