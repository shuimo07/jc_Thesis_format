# -*- coding: utf-8 -*-
"""
jc_thesis_format — 成都锦城学院 本科毕业论文（设计）排版工具库
==================================================================

本库把「成都锦城学院本科毕业论文」的版式规范固化成可复用的 python-docx 原语，
参数全部来自对参考模板 / 样本论文页面截图的像素级测量（见 docs/format-spec.md）。

规范摘要
--------
页面      A4 210×297mm，上下 2.5cm / 左 3.0cm / 右 2.5cm（正文宽 15.48cm）
页眉      「成都锦城学院毕业论文（设计）」，宋体五号居中，下加双线，距顶 2.0cm
页脚      页码居中，宋体五号
页码格式  封面不编页码；前置部分（声明/授权书/摘要/Abstract/目录）大写罗马数字 I,II,III…
          正文起阿拉伯数字 1,2,3…
一级标题  黑体小二(18pt)加粗左对齐，段前 30 磅，固定行距 28 磅，每章另起页
二级标题  黑体四号(14pt)加粗左对齐，段前 5 磅
三级标题  黑体小四(12pt)加粗左对齐，段前 7 磅
正文      宋体小四(12pt)，首行缩进 2 字符，两端对齐，固定行距 22 磅
题注      黑体五号(10.5pt)加粗居中；表题在表上方、图题在图下方
表格      三线表：顶/底 1.5pt 单线，表头下 0.75pt 单线，无竖线；表内宋体五号
参考文献  宋体小四，顶格（无首行缩进），固定行距 22 磅
目录      标题黑体小二居中；条目小四，固定行距 20 磅，缩进 0 / 2 / 4 字符，点线前导
封面      校名+「本科生毕业论文（设计）」黑体小一(24pt)；字段四号(14pt)带下划线

用法
----
    from jc_thesis_format import *
    doc = Document()
    apply_page(doc.sections[0])
    apply_base_style(doc)
    add_chapter_body_styles(doc)
    add_toc_styles(doc)
    build_cover(doc, title='…', college='…', major='…', name='…', sid='…')
    sec = new_section(doc, 'upperRoman', 1)
    build_integrity_statement(doc)
    ...
"""

import copy

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

__all__ = [
    'SPEC', 'IDEO_SPACE', 'TEXT_WIDTH_TW', 'CM2TW',
    'ALIGN_LEFT', 'ALIGN_CENTER', 'ALIGN_RIGHT', 'ALIGN_BOTH',
    'ppr_set', 'rf', 'pf', 'new_p',
    'apply_page', 'set_page_numbering', 'add_header', 'add_footer_page_number',
    'apply_base_style', 'add_chapter_body_styles', 'add_toc_styles',
    'new_section', 'enable_update_fields', 'field_paragraph',
    'make_three_line_table', 'caption_paragraph', 'toc_field',
    'build_cover', 'build_integrity_statement', 'build_authorization',
    'build_cn_abstract', 'build_en_abstract',
]

# ------------------------------------------------------------------ 规范常量

SPEC = {
    'page_w_cm': 21.0,
    'page_h_cm': 29.7,
    'margin_top_cm': 2.5,
    'margin_bottom_cm': 2.5,
    'margin_left_cm': 3.0,
    'margin_right_cm': 2.5,
    'header_distance_cm': 2.0,
    'footer_distance_cm': 2.2,
    'header_text': '成都锦城学院毕业论文（设计）',

    # 标题
    'h1_size': 18.0, 'h1_before': 30.0, 'h1_line': 28.0,
    'h2_size': 14.0, 'h2_before': 5.0,
    'h3_size': 12.0, 'h3_before': 7.0,

    # 正文
    'body_size': 12.0, 'body_line': 22.0, 'body_first_indent_chars': 2.0,

    # 题注 / 表格 / 目录 / 封面 / 前置部分
    'caption_size': 10.5, 'caption_line': 16.0,
    'table_text_size': 10.5, 'table_line': 16.0,
    'toc_line': 20.0,
    'cover_title_size': 24.0, 'cover_field_size': 14.0,
    'front_size': 16.0, 'front_line': 40.0,

    # 三线表线宽（八分之一磅）
    'tbl_top_bottom_sz': 12,   # 1.5pt
    'tbl_header_sz': 6,        # 0.75pt
}

IDEO_SPACE = '\u3000'          # 全角空格，用于封面标签排版
CM2TW = 567                    # 1cm = 567 twips
TEXT_WIDTH_TW = int((SPEC['page_w_cm'] - SPEC['margin_left_cm']
                     - SPEC['margin_right_cm']) * CM2TW)   # 8777 tw

ALIGN_LEFT = WD_ALIGN_PARAGRAPH.LEFT
ALIGN_CENTER = WD_ALIGN_PARAGRAPH.CENTER
ALIGN_RIGHT = WD_ALIGN_PARAGRAPH.RIGHT
ALIGN_BOTH = WD_ALIGN_PARAGRAPH.JUSTIFY

# w:pPr 子元素的法定顺序——顺序错了 Word 会拒开文件
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
             'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
             'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
             'sectPr', 'pPrChange']


# ------------------------------------------------------------------ 底层原语

def ppr_set(pPr, tag, elem):
    """把 elem 作为 w:<tag> 插入 pPr，并维持 CT_PPr 的法定子元素顺序。"""
    old = pPr.find(qn('w:' + tag))
    if old is not None:
        pPr.remove(old)
    i = PPR_ORDER.index(tag)
    for ch in pPr:
        nm = ch.tag.split('}')[1]
        if nm in PPR_ORDER and PPR_ORDER.index(nm) > i:
            ch.addprevious(elem)
            return elem
    pPr.append(elem)
    return elem


def rf(run, cjk='宋体', latin='Times New Roman', size=12.0, bold=False,
       ul=False, italic=False, superscript=False):
    """统一设置 run 的中西文字体/字号/字形（必须显式写 w:eastAsia，否则中文回落字体）。"""
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = ul
    if superscript:
        run.font.superscript = True
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    rFonts.set(qn('w:eastAsia'), cjk)
    rFonts.set(qn('w:cs'), latin)
    return run


def pf(p, align=None, before=None, after=None, line=None, first=None, left=None,
       right=None, hanging=None, page_break=False):
    """段落格式。line 为磅值表示「固定行距」；first/left/right/hanging 以「字符」为单位。"""
    f = p.paragraph_format
    if align is not None:
        f.alignment = align
    if before is not None:
        f.space_before = Pt(before)
    if after is not None:
        f.space_after = Pt(after)
    if line is not None:
        f.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        f.line_spacing = Pt(line)
    if page_break:
        f.page_break_before = True
    pPr = p._element.get_or_add_pPr()
    if any(v is not None for v in (first, left, right, hanging)):
        ind = OxmlElement('w:ind')
        # 同时写 Chars 与绝对值：Chars 让缩进随字号自适应，绝对值兜底
        if first is not None:
            ind.set(qn('w:firstLineChars'), str(int(round(first * 100))))
            ind.set(qn('w:firstLine'), str(int(round(first * 12 * 20))))
        if hanging is not None:
            ind.set(qn('w:hangingChars'), str(int(round(hanging * 100))))
            ind.set(qn('w:hanging'), str(int(round(hanging * 12 * 20))))
        if left is not None:
            ind.set(qn('w:leftChars'), str(int(round(left * 100))))
            ind.set(qn('w:left'), str(int(round(left * 12 * 20))))
        if right is not None:
            ind.set(qn('w:rightChars'), str(int(round(right * 100))))
            ind.set(qn('w:right'), str(int(round(right * 12 * 20))))
        ppr_set(pPr, 'ind', ind)
    return p


def new_p(container, text='', **kw):
    """建段落的快捷函数：排版参数走 pf()，字体参数走 rf()。"""
    p = container.add_paragraph()
    text_kw = {k: kw.pop(k) for k in ('cjk', 'latin', 'size', 'bold', 'ul') if k in kw}
    pf(p, **kw)
    if text:
        rf(p.add_run(text), **text_kw)
    return p


# ------------------------------------------------------------------ 页面 / 节

def apply_page(sect, spec=SPEC):
    """A4 + 学院规定的页边距与页眉页脚距离。"""
    sect.page_width = Cm(spec['page_w_cm'])
    sect.page_height = Cm(spec['page_h_cm'])
    sect.top_margin = Cm(spec['margin_top_cm'])
    sect.bottom_margin = Cm(spec['margin_bottom_cm'])
    sect.left_margin = Cm(spec['margin_left_cm'])
    sect.right_margin = Cm(spec['margin_right_cm'])
    sect.header_distance = Cm(spec['header_distance_cm'])
    sect.footer_distance = Cm(spec['footer_distance_cm'])
    return sect


def set_page_numbering(sect, fmt='decimal', start=1):
    """设置本节页码格式。fmt: 'decimal' | 'upperRoman' | 'lowerRoman'。"""
    sectPr = sect._sectPr
    el = OxmlElement('w:pgNumType')
    el.set(qn('w:fmt'), fmt)
    el.set(qn('w:start'), str(start))
    old = sectPr.find(qn('w:pgNumType'))
    if old is not None:
        sectPr.remove(old)
    # pgNumType 必须排在 pgMar/lnNumType 之后、cols/docGrid 之前
    ref = None
    for tag in ('w:cols', 'w:formProt', 'w:vAlign', 'w:titlePg', 'w:textDirection',
                'w:bidi', 'w:rtlGutter', 'w:docGrid', 'w:printerSettings'):
        ref = sectPr.find(qn(tag))
        if ref is not None:
            break
    if ref is not None:
        ref.addprevious(el)
    else:
        sectPr.append(el)


def add_header(sect, text=None):
    """页眉：宋体五号居中 + 下双线。"""
    text = SPEC['header_text'] if text is None else text
    sect.header.is_linked_to_previous = False
    p = sect.header.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    pf(p, align=ALIGN_CENTER, before=0, after=0)
    rf(p.add_run(text), cjk='宋体', size=10.5)
    bdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'double')
    bottom.set(qn('w:sz'), '8')      # 每线 1pt
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    bdr.append(bottom)
    ppr_set(p._element.get_or_add_pPr(), 'pBdr', bdr)


def add_footer_page_number(sect):
    """页脚：居中的 PAGE 域。

    注意必须用完整的 begin/instrText/separate/占位/end 五件套。只写 instrText
    不写 fldChar 时，Word 会把它当成普通文本直接把 ' PAGE ' 打印出来。
    """
    sect.footer.is_linked_to_previous = False
    p = sect.footer.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    pf(p, align=ALIGN_CENTER, before=0, after=0)

    def fld(t):
        r = p.add_run()
        rf(r, cjk='宋体', size=10.5)
        e = OxmlElement('w:fldChar')
        e.set(qn('w:fldCharType'), t)
        r._element.append(e)
        return r

    fld('begin')
    r = p.add_run()
    rf(r, cjk='宋体', size=10.5)
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = ' PAGE '
    r._element.append(it)
    fld('separate')
    rf(p.add_run('1'), cjk='宋体', size=10.5)   # 缓存值，Word 打开后会重算
    fld('end')


def new_section(doc, num_fmt='decimal', start=1, header=True, footer=True):
    """另起一节（新页），设置页码格式、页眉、页脚，返回该节。"""
    sect = apply_page(doc.add_section(WD_SECTION.NEW_PAGE))
    set_page_numbering(sect, num_fmt, start)
    if header:
        add_header(sect)
    if footer:
        add_footer_page_number(sect)
    return sect


def enable_update_fields(doc):
    """让 Word 打开文档时自动更新所有域（目录页码）。"""
    stg = doc.settings.element
    uf = stg.find(qn('w:updateFields'))
    if uf is None:
        uf = OxmlElement('w:updateFields')
        stg.append(uf)
    uf.set(qn('w:val'), 'true')


# ------------------------------------------------------------------ 样式

def apply_base_style(doc, spec=SPEC):
    """Normal 样式：西文 Times New Roman + 中文宋体，小四。"""
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(spec['body_size'])
    st.element.rPr.get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)
    return st


def _style(doc, name, size, cjk, latin='Times New Roman', bold=False, align='both',
           first=None, line=22.0, before=0.0, after=0.0):
    s = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    s.base_style = doc.styles['Normal']
    s.font.name = latin
    s.font.size = Pt(size)
    s.font.bold = bold
    s.element.rPr.get_or_add_rFonts().set(qn('w:eastAsia'), cjk)
    pPr = s.element.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(int(before * 20)))
    sp.set(qn('w:after'), str(int(after * 20)))
    if line is None:
        sp.set(qn('w:line'), '240')
        sp.set(qn('w:lineRule'), 'auto')
    else:
        sp.set(qn('w:line'), str(int(line * 20)))
        sp.set(qn('w:lineRule'), 'exact')
    ppr_set(pPr, 'spacing', sp)
    ind = OxmlElement('w:ind')
    if first is not None:
        ind.set(qn('w:firstLineChars'), str(int(first * 100)))
        ind.set(qn('w:firstLine'), str(int(first * size * 20)))
    else:
        ind.set(qn('w:firstLine'), '0')
    ind.set(qn('w:left'), '0')
    ppr_set(pPr, 'ind', ind)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), align)
    ppr_set(pPr, 'jc', jc)
    return s


def add_chapter_body_styles(doc, spec=SPEC):
    """建立正文/图表题注/参考文献等自定义样式，并配置 Heading 1-3。"""
    apply_base_style(doc, spec)

    for name, size, before, line in (('Heading 1', spec['h1_size'], spec['h1_before'], spec['h1_line']),
                                     ('Heading 2', spec['h2_size'], spec['h2_before'], spec['body_line']),
                                     ('Heading 3', spec['h3_size'], spec['h3_before'], spec['body_line'])):
        s = doc.styles[name]
        s.font.name = 'Times New Roman'
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.italic = False
        s.font.color.rgb = None
        s.element.rPr.get_or_add_rFonts().set(qn('w:eastAsia'), '黑体')
        s.element.rPr.get_or_add_rFonts().set(qn('w:cs'), 'Times New Roman')
        pPr = s.element.get_or_add_pPr()
        ppr_set(pPr, 'spacing', OxmlElement('w:spacing'))
        sp = pPr.find(qn('w:spacing'))
        sp.set(qn('w:beforeLines'), '0')
        sp.set(qn('w:before'), str(int(before * 20)))
        sp.set(qn('w:after'), '0')
        sp.set(qn('w:line'), str(int(line * 20)))
        sp.set(qn('w:lineRule'), 'exact')
        ppr_set(pPr, 'ind', OxmlElement('w:ind'))
        ind = pPr.find(qn('w:ind'))
        ind.set(qn('w:firstLine'), '0')
        ind.set(qn('w:left'), '0')
        ppr_set(pPr, 'jc', OxmlElement('w:jc'))
        pPr.find(qn('w:jc')).set(qn('w:val'), 'left')

    _style(doc, '论文正文', spec['body_size'], '宋体', first=spec['body_first_indent_chars'],
           line=spec['body_line'], align='both')
    _style(doc, '图题', spec['caption_size'], '黑体', bold=True, align='center',
           line=spec['caption_line'], before=3, after=6)
    _style(doc, '表题', spec['caption_size'], '黑体', bold=True, align='center',
           line=spec['caption_line'], before=6, after=3)
    _style(doc, '表格文字', spec['table_text_size'], '宋体', align='center',
           line=spec['table_line'])
    _style(doc, '参考文献', spec['body_size'], '宋体', first=None,
           line=spec['body_line'], align='both')
    _style(doc, '图片段', spec['body_size'], '宋体', align='center', line=None)


def add_toc_styles(doc):
    """目录条目样式 TOC1/TOC2/TOC3：小四、固定行距 20 磅、缩进 0/2/4 字符 + 右对齐点线制表位。

    styleId 必须是 TOC1/TOC2/TOC3，Word 的 TOC 域才能把条目套上去。
    """
    def one(style_id, name, indent_chars):
        s = OxmlElement('w:style')
        s.set(qn('w:type'), 'paragraph')
        s.set(qn('w:styleId'), style_id)
        nm = OxmlElement('w:name')
        nm.set(qn('w:val'), name)
        s.append(nm)
        bo = OxmlElement('w:basedOn')
        bo.set(qn('w:val'), 'Normal')
        s.append(bo)

        pPr = OxmlElement('w:pPr')
        tabs = OxmlElement('w:tabs')
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'right')
        tab.set(qn('w:leader'), 'dot')
        tab.set(qn('w:pos'), str(TEXT_WIDTH_TW))
        tabs.append(tab)
        pPr.append(tabs)
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '0')
        sp.set(qn('w:line'), str(int(SPEC['toc_line'] * 20)))
        sp.set(qn('w:lineRule'), 'exact')
        pPr.append(sp)
        ind = OxmlElement('w:ind')
        if indent_chars:
            ind.set(qn('w:leftChars'), str(indent_chars * 100))
            ind.set(qn('w:left'), str(int(indent_chars * 12 * 20)))
        ind.set(qn('w:firstLine'), '0')
        pPr.append(ind)
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'left')
        pPr.append(jc)
        s.append(pPr)

        rPr = OxmlElement('w:rPr')
        f = OxmlElement('w:rFonts')
        f.set(qn('w:ascii'), 'Times New Roman')
        f.set(qn('w:hAnsi'), 'Times New Roman')
        f.set(qn('w:eastAsia'), '宋体')
        rPr.append(f)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), str(int(SPEC['body_size'] * 2)))
        rPr.append(sz)
        szcs = OxmlElement('w:szCs')
        szcs.set(qn('w:val'), str(int(SPEC['body_size'] * 2)))
        rPr.append(szcs)
        s.append(rPr)
        doc.styles.element.append(s)

    one('TOC1', 'toc 1', 0)
    one('TOC2', 'toc 2', 2)
    one('TOC3', 'toc 3', 4)


# ------------------------------------------------------------------ 内容构件

def field_paragraph(doc, label, value, label_pos_tw=720, spec=SPEC):
    """封面字段行：标签 + 居中值 + 连续下划线填充。

    坑：Word 不会给「行尾空白」画下划线，所以右侧填充线不能用下划线空格，
    必须用右制表位 + underscore 前导符；左侧填充线用下划线空格才不留缝。
    """
    size = spec['cover_field_size']
    label_w = 4 * size * 20                       # 标签 4 字宽
    fill_s = label_pos_tw + label_w
    fill_e = TEXT_WIDTH_TW
    fill_em = (fill_e - fill_s) / CM2TW / (size * 0.03528)

    p = doc.add_paragraph()
    pf(p, align=ALIGN_LEFT, before=8, after=16)
    pPr = p._element.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(label_pos_tw))
    ind.set(qn('w:firstLine'), '0')
    ppr_set(pPr, 'ind', ind)
    tabs = OxmlElement('w:tabs')
    t = OxmlElement('w:tab')
    t.set(qn('w:val'), 'right')
    t.set(qn('w:pos'), str(fill_e))
    t.set(qn('w:leader'), 'underscore')
    tabs.append(t)
    ppr_set(pPr, 'tabs', tabs)

    v_em = sum(1.0 if ord(c) > 0x2000 else 0.5 for c in value)
    n_half = int((fill_em - v_em) / 2.0 / 0.5)     # 向下取整，避免折行
    rf(p.add_run(label), cjk='宋体', size=size)
    if n_half > 0:
        rf(p.add_run(' ' * n_half), cjk='宋体', latin='宋体', size=size, ul=True)
    if value:
        rf(p.add_run(value), cjk='宋体', size=size)
    r = p.add_run()
    rf(r, cjk='宋体', size=size)
    r._element.append(OxmlElement('w:tab'))
    return p


def build_cover(doc, title='', college='', major='', name='', sid='', advisor='',
                date_text='20XX 年 X 月', school='成都锦城学院',
                degree_line='本科生毕业论文（设计）', spec=SPEC):
    """封面页。缺省字段留空，下划线照画，学生手填或后续替换。"""
    size = spec['cover_title_size']
    p = new_p(doc, align=ALIGN_CENTER, before=72, after=0)
    rf(p.add_run(school), cjk='黑体', size=size, bold=True)
    p = new_p(doc, align=ALIGN_CENTER, before=6, after=0)
    rf(p.add_run(degree_line), cjk='黑体', size=size, bold=True)
    new_p(doc, before=0, after=0)
    new_p(doc, before=0, after=0)

    # 题目可能超一行：用悬挂缩进让折行后的文字与填充起始位置对齐
    fsize = spec['cover_field_size']
    label_w = 4 * fsize * 20
    fill_s = 720 + label_w
    p = doc.add_paragraph()
    pf(p, align=ALIGN_LEFT, before=8, after=16)
    pPr = p._element.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(fill_s))
    ind.set(qn('w:hanging'), str(label_w))
    ppr_set(pPr, 'ind', ind)
    tabs = OxmlElement('w:tabs')
    t = OxmlElement('w:tab')
    t.set(qn('w:val'), 'left')
    t.set(qn('w:pos'), str(fill_s))
    tabs.append(t)
    ppr_set(pPr, 'tabs', tabs)
    rf(p.add_run('题' + IDEO_SPACE * 2 + '目'), cjk='宋体', size=fsize)
    r = p.add_run()
    rf(r, cjk='宋体', size=fsize)
    r._element.append(OxmlElement('w:tab'))
    rf(p.add_run(title), cjk='宋体', size=fsize, ul=True)

    field_paragraph(doc, '二级学院', college, spec=spec)
    field_paragraph(doc, '专' + IDEO_SPACE * 2 + '业', major, spec=spec)
    field_paragraph(doc, '学生姓名', name, spec=spec)
    field_paragraph(doc, '学' + IDEO_SPACE * 2 + '号', sid, spec=spec)
    field_paragraph(doc, '指导教师', advisor, spec=spec)

    new_p(doc, before=0, after=0)
    p = new_p(doc, align=ALIGN_RIGHT, before=0, after=0, right=5.0)
    rf(p.add_run('教务处' + IDEO_SPACE + '制表'), cjk='宋体', size=fsize)
    p = new_p(doc, align=ALIGN_RIGHT, before=0, after=0, right=5.0)
    rf(p.add_run(date_text), cjk='宋体', size=fsize)


def _front_title(doc, t1, t2, page_break=False):
    """前置部分大标题：两行黑体，小二偏大（22pt）。"""
    p = new_p(doc, align=ALIGN_CENTER, before=60, after=0, page_break=page_break)
    rf(p.add_run(t1), cjk='黑体', size=22, bold=True)
    p = new_p(doc, align=ALIGN_CENTER, before=12, after=36)
    rf(p.add_run(t2), cjk='黑体', size=22, bold=True)


INTEGRITY_TEXT = ('本人郑重声明：所呈交的毕业论文（设计），是本人在导师的指导下，独立进行实践及研究'
                  '工作所取得的成果。除文中已经注明引用的内容外，本论文（设计）不包含任何其他个人或'
                  '集体已经发表或撰写过的作品成果。对本文的研究做出重要贡献的个人和集体，均已在文中'
                  '以明确方式标明。本人完全意识到本声明的法律结果由本人承担。')

AUTHORIZE_TEXT = ('本毕业论文（设计）作者同意学校保留并向国家有关部门或机构送交论文（设计）的复印件'
                  '和电子版，允许论文（设计）被查阅和借阅。本人授权成都锦城学院可以将本毕业论文'
                  '（设计）的全部或部分内容编入有关数据库进行检索，可以采用影印、缩印或扫描等复制'
                  '手段保存和汇编本毕业论文（设计）。')


def build_integrity_statement(doc, text=INTEGRITY_TEXT, year=2026, month=6, spec=SPEC):
    """学术诚信声明页。"""
    fs, ln = spec['front_size'], spec['front_line']
    _front_title(doc, '成都锦城学院', '毕业论文（设计）' + IDEO_SPACE + '学术诚信声明')
    p = new_p(doc, align=ALIGN_BOTH, before=0, after=0, line=ln, first=2)
    rf(p.add_run(text), cjk='宋体', size=fs)
    new_p(doc, before=0, after=0, line=ln)
    p = new_p(doc, align=ALIGN_RIGHT, before=0, after=0, line=ln, right=5.0)
    rf(p.add_run('作者签名'), cjk='宋体', size=fs)
    p = new_p(doc, align=ALIGN_LEFT, before=0, after=0, line=ln, left=5.09 / 0.494)
    rf(p.add_run('日期：' + IDEO_SPACE + f'{year} 年' + IDEO_SPACE * 2 + f'{month} 月'
                 + IDEO_SPACE * 4 + '日'), cjk='宋体', size=fs)


def build_authorization(doc, text=AUTHORIZE_TEXT, year=2026, month=6, spec=SPEC):
    """版权使用授权书页（作者 / 指导教师双签署栏并排）。"""
    fs, ln = spec['front_size'], spec['front_line']
    _front_title(doc, '成都锦城学院', '毕业论文（设计）' + IDEO_SPACE + '版权使用授权书',
                 page_break=True)
    p = new_p(doc, align=ALIGN_BOTH, before=0, after=0, line=ln, first=2)
    rf(p.add_run(text), cjk='宋体', size=fs)
    new_p(doc, before=0, after=0, line=ln)
    new_p(doc, before=0, after=0, line=ln)
    p = new_p(doc, align=ALIGN_LEFT, before=0, after=0, line=ln, left=1.03 / 0.494)
    rf(p.add_run('作者签名：' + IDEO_SPACE * 9 + '指导教师签名：'), cjk='宋体', size=14)
    p = new_p(doc, align=ALIGN_LEFT, before=0, after=0, line=ln, left=1.09 / 0.494)
    rf(p.add_run(f'日期：{year} 年 {month} 月' + IDEO_SPACE + '日' + IDEO_SPACE * 6
                 + f'日期：{year} 年 {month} 月' + IDEO_SPACE + '日'), cjk='宋体', size=14)


def build_cn_abstract(doc, text='', keywords=(), spec=SPEC):
    """中文摘要页：摘要/关键词标签加粗小四，正文小四首行缩进 2 字符。"""
    p = new_p(doc, align=ALIGN_BOTH, before=0, after=0, line=SPEC['body_line'],
              first=2, page_break=True)
    rf(p.add_run('摘要：'), cjk='宋体', size=12, bold=True)
    rf(p.add_run(text), cjk='宋体', size=12)
    p = new_p(doc, align=ALIGN_BOTH, before=12, after=0, line=SPEC['body_line'], first=2)
    rf(p.add_run('关键词：'), cjk='宋体', size=12, bold=True)
    rf(p.add_run('；'.join(keywords) if keywords else ''), cjk='宋体', size=12)


def build_en_abstract(doc, text='', keywords=(), spec=SPEC):
    """英文摘要页。"""
    p = new_p(doc, align=ALIGN_BOTH, before=0, after=0, line=SPEC['body_line'],
              first=2, page_break=True)
    rf(p.add_run('Abstract: '), cjk='宋体', size=12, bold=True)
    rf(p.add_run(text), cjk='宋体', size=12)
    p = new_p(doc, align=ALIGN_BOTH, before=12, after=0, line=SPEC['body_line'], first=2)
    rf(p.add_run('Key Words: '), cjk='宋体', size=12, bold=True)
    rf(p.add_run('; '.join(keywords) if keywords else ''), cjk='宋体', size=12)


def make_three_line_table(doc, grid, header_rows=1, spec=SPEC):
    """三线表：顶线/底线 1.5pt，表头下 0.75pt，无竖线与内横线。"""
    rows = len(grid)
    cols = max(len(r) for r in grid)
    tb = doc.add_table(rows=rows, cols=cols)
    tb.autofit = True
    for ri, row in enumerate(grid):
        for ci in range(cols):
            val = row[ci] if ci < len(row) else ''
            cp = tb.cell(ri, ci).paragraphs[0]
            cp.style = doc.styles['表格文字']
            pf(cp, align=ALIGN_CENTER, before=1, after=1, line=spec['table_line'])
            rf(cp.add_run(val), cjk='宋体', size=spec['table_text_size'],
               bold=(ri < header_rows))

    tblPr = tb._tbl.tblPr
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    bd = OxmlElement('w:tblBorders')
    want = {}
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        if edge in ('top', 'bottom'):
            e.set(qn('w:val'), 'single')
            e.set(qn('w:sz'), str(spec['tbl_top_bottom_sz']))
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), 'auto')
        else:
            e.set(qn('w:val'), 'none')
            e.set(qn('w:sz'), '0')
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), 'auto')
        want[edge] = e
    # tblBorders 的子元素顺序固定为 top,left,bottom,right,insideH,insideV
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        bd.append(want[edge])
    ref = None
    for tag in ('w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption',
                'w:tblDescription'):
        ref = tblPr.find(qn(tag))
        if ref is not None:
            break
    if ref is not None:
        ref.addprevious(bd)
    else:
        tblPr.append(bd)

    # 表头行下边线 0.75pt
    for ci in range(cols):
        tcPr = tb.cell(header_rows - 1, ci)._tc.get_or_add_tcPr()
        tb_borders = OxmlElement('w:tcBorders')
        e = OxmlElement('w:bottom')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), str(spec['tbl_header_sz']))
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), 'auto')
        tb_borders.append(e)
        tcPr.append(tb_borders)
    return tb


def caption_paragraph(doc, text, kind='fig', spec=SPEC):
    """题注：kind='fig' → 图题（放图下方）；kind='tab' → 表题（放表上方）。"""
    if kind == 'fig':
        p = doc.add_paragraph(style='图题')
        pf(p, align=ALIGN_CENTER, before=3, after=6, line=spec['caption_line'])
    else:
        p = doc.add_paragraph(style='表题')
        pf(p, align=ALIGN_CENTER, before=6, after=3, line=spec['caption_line'])
    rf(p.add_run(text), cjk='黑体', size=spec['caption_size'], bold=True)
    return p


def toc_field(doc, cached_entries=(), levels='1-3'):
    """插入多段式 TOC 域。

    cached_entries 是「打开文档时还没刷新」的占位条目，用 (level, text) 给出；
    Word 打开后（文档已开 updateFields）会用真实标题和页码覆盖它们。
    """
    paras = []
    n = len(cached_entries)
    for idx, (lvl, txt) in enumerate(cached_entries):
        p = doc.add_paragraph()
        p.style = doc.styles['TOC%d' % lvl]
        pf(p, line=SPEC['toc_line'])
        if idx == 0:
            r = p.add_run()
            rf(r, size=SPEC['body_size'])
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), 'begin')
            fc.set(qn('w:dirty'), 'true')
            r._element.append(fc)
            r = p.add_run()
            rf(r, size=SPEC['body_size'])
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = ' TOC \\o "%s" \\h \\z \\u ' % levels
            r._element.append(it)
            r = p.add_run()
            rf(r, size=SPEC['body_size'])
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), 'separate')
            r._element.append(fc)
        rf(p.add_run(txt), cjk='宋体', size=SPEC['body_size'])
        r = p.add_run()
        rf(r, cjk='宋体', size=SPEC['body_size'])
        r._element.append(OxmlElement('w:tab'))
        if idx == n - 1:
            r = p.add_run()
            rf(r, size=SPEC['body_size'])
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), 'end')
            r._element.append(fc)
        paras.append(p)
    return paras


def toc_heading(doc, text='目录', page_break=True, spec=SPEC):
    """目录页标题：黑体小二居中。"""
    p = new_p(doc, align=ALIGN_CENTER, before=0, after=18, page_break=page_break)
    rf(p.add_run(IDEO_SPACE.join(text)), cjk='黑体', size=spec['h1_size'], bold=True)
    return p
