# -*- coding: utf-8 -*-
"""
refresh_and_export.py — 用本机 Word 刷新域（目录页码）并导出 PDF
================================================================

python-docx 只能往文档里写「TOC 域」这个空壳，页码要真正的排版引擎才算得出来。
这个脚本调用本机 Word（COM 自动化）完成三件事：

    1. 打开文档
    2. 更新所有域（TOC / PAGE / NUMPAGES）并保存
    3. 另存为 PDF，供逐页栅格化后肉眼/程序校验

依赖：pip install pywin32
本机需安装 Microsoft Word。WPS 的 COM 接口（KWPS.Application）域更新不完整，
不建议用它做这一步。

用法：
    python refresh_and_export.py 论文.docx            # 刷新域 + 导出同名 pdf
    python refresh_and_export.py 论文.docx --no-pdf   # 只刷新域
    python refresh_and_export.py 论文.docx -o out/    # 指定 pdf 输出目录
"""

import argparse
import os
import sys

WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_PDF = 17


def refresh(path, pdf=True, outdir=None):
    import win32com.client as win32

    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise SystemExit('找不到文件: %s' % path)

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(path)
        # 目录 / 页码域全部重算
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        doc.Save()

        if pdf:
            base = os.path.splitext(os.path.basename(path))[0]
            outdir = outdir or os.path.dirname(path)
            os.makedirs(outdir, exist_ok=True)
            pdf_path = os.path.join(outdir, base + '.pdf')
            doc.ExportAsFixedFormat(pdf_path, WD_FORMAT_PDF)
            print('PDF:', pdf_path)
        print('domain refreshed:', path)
    finally:
        try:
            doc.Close(WD_DO_NOT_SAVE_CHANGES if False else -1)
        except Exception:
            pass
        word.Quit()


def raster(pdf_path, outdir=None, dpi=110):
    """把 PDF 逐页转成 PNG，便于逐页比对（需要 pymupdf）。"""
    import pymupdf
    outdir = outdir or os.path.splitext(pdf_path)[0] + '_pages'
    os.makedirs(outdir, exist_ok=True)
    d = pymupdf.open(pdf_path)
    for i in range(d.page_count):
        d[i].get_pixmap(dpi=dpi).save(os.path.join(outdir, 'p%03d.png' % (i + 1)))
    print('pages:', d.page_count, '->', outdir)
    return d.page_count


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--no-pdf', action='store_true')
    ap.add_argument('-o', '--outdir', default=None)
    ap.add_argument('--raster', action='store_true', help='导出 PDF 后顺便栅格化')
    a = ap.parse_args()
    try:
        refresh(a.docx, pdf=not a.no_pdf, outdir=a.outdir)
    except ImportError:
        print('缺少 pywin32：pip install pywin32', file=sys.stderr)
        raise SystemExit(1)
    if a.raster and not a.no_pdf:
        base = os.path.splitext(os.path.basename(a.docx))[0]
        raster(os.path.join(a.outdir or os.path.dirname(os.path.abspath(a.docx)),
                            base + '.pdf'))
