# -*- coding: utf-8 -*-
"""
measure_reference.py — 从参考论文的页面截图里反推版面参数
==========================================================

为什么要这样做
--------------
拿到的是 PDF/图片形式的模板和样张，不是 .docx，没法直接读出字号行距。
但截图里每条文字的墨迹范围是可测的，于是可以从「几何」反推「排版参数」：

  行距     相邻两行墨迹顶部的间距
  字号     单行汉字墨迹高度 → pt（汉字墨迹高 ≈ 0.92 em，1pt = 0.3528mm
           ⇒ pt = h_mm / 0.3528 / 0.92 = h_mm / 0.3246）
  缩进     一个文字块左右墨迹相对版心边缘的偏移量
  页边距   版心边框（或最长行的起止）相对图像边缘的距离
  表格线   整行/整列墨迹覆盖率 > 阈值的长直线，其厚度即线宽

前提假设
--------
截图是「整页宽度」截取的，即图像宽度 = A4 纸宽 210mm。若不是，用
--page-width-mm 指定；若图像有裁边，用 --ink 先看墨迹范围再校准。

用法
----
    python measure_reference.py page   参考图.png
    python measure_reference.py lines  参考图.png [--block 0] [--gap 4]
    python measure_reference.py table  参考图.png
    python measure_reference.py all    参考图.png
    python measure_reference.py lines  D:/ref/*.png      # 支持通配
"""

import argparse
import glob
import os
import sys

import numpy as np
from PIL import Image

MM_PER_PT = 0.3528
CJK_INK_RATIO = 0.92        # 汉字墨迹高度占 em 的比例（经验值）
DARK = 170                  # 灰度阈值：小于它算墨迹


def load(path):
    im = Image.open(path).convert('L')
    return np.array(im)


def mm_per_px(arr, page_w_mm):
    return page_w_mm / arr.shape[1]


def blocks_of(arr, gap=4, axis=1):
    """把墨迹按连续行（或列）切成块，返回 [(start, end), ...]。"""
    dark = arr < DARK
    if axis == 1:                      # 按行切
        prof = dark.sum(axis=1)
    else:                              # 按列切
        prof = dark.sum(axis=0)
    idx = np.where(prof > 0)[0]
    if len(idx) == 0:
        return []
    out = []
    s = p = idx[0]
    for r in idx[1:]:
        if r - p > gap:
            out.append((s, p))
            s = r
        p = r
    out.append((s, p))
    return out


def ink_box(arr):
    dark = arr < DARK
    rows = np.where(dark.sum(axis=1) > 1)[0]
    cols = np.where(dark.sum(axis=0) > 1)[0]
    if len(rows) == 0 or len(cols) == 0:
        return None
    return int(rows.min()), int(rows.max()), int(cols.min()), int(cols.max())


def cmd_page(path, page_w_mm):
    arr = load(path)
    h, w = arr.shape
    cm = mm_per_px(arr, page_w_mm)
    box = ink_box(arr)
    print(f'=== {os.path.basename(path)}  image {w}x{h}px  {cm:.4f} mm/px')
    if not box:
        print('   (空白页 / 无墨迹)')
        return
    r0, r1, c0, c1 = box
    print(f'   墨迹范围: rows {r0}..{r1}  cols {c0}..{c1}')
    print(f'   距图边   : top {r0*cm:6.1f}mm  bottom {(h-1-r1)*cm:6.1f}mm'
          f'  left {c0*cm:6.1f}mm  right {(w-1-c1)*cm:6.1f}mm')
    print('   若这是正文页，上面的 top/bottom/left/right 即版心边距'
          '（页眉页脚会额外占位，需按页判断）')
    return box


def cmd_lines(path, page_w_mm, gap=4, which=None, max_blocks=40):
    arr = load(path)
    cm = mm_per_px(arr, page_w_mm)
    box = ink_box(arr)
    if not box:
        print('   空白页')
        return
    r0, r1, c0, c1 = box
    sub = arr[:, c0:c1 + 1]
    blocks = blocks_of(sub, gap=gap, axis=1)
    if which is not None:
        blocks = [blocks[which]] if which < len(blocks) else []
    print(f'=== {os.path.basename(path)}  text block cols {c0}..{c1}')
    prev_top = None
    for i, (s, e) in enumerate(blocks[:max_blocks]):
        seg = (sub[s:e + 1] < DARK)
        cols = np.where(seg.sum(axis=0) > 0)[0]
        lx, rx = c0 + int(cols.min()), c0 + int(cols.max())
        gh = (e - s + 1) * cm
        pt = gh / (MM_PER_PT * CJK_INK_RATIO)
        pitch = '' if prev_top is None else f'  pitch={(s-prev_top)*cm:5.2f}mm'
        # 若行距固定 22 磅 ≈ 7.76mm，可用 pitch 直接确认
        print(f'   blk{i:2d} rows {s:4d}-{e:4d}  h={gh:4.2f}mm →≈{pt:5.1f}pt'
              f'  L={(lx-c0)*cm:5.2f}mm  R={(c1-rx)*cm:5.2f}mm{pitch}')
        prev_top = s
    if len(blocks) > max_blocks:
        print(f'   ... 共 {len(blocks)} 块')
    print('   提示：pitch 稳定且 ≈7.76mm 即固定行距 22 磅；'
          '≈8.47mm 对应 24 磅（1.5 倍），≈11.29mm 对应 32 磅。')
    return blocks


def cmd_table(path, page_w_mm, cover=0.8):
    """找横贯整版的长横线，报告位置与厚度（三线表的判别依据）。"""
    arr = load(path)
    cm = mm_per_px(arr, page_w_mm)
    dark = arr < DARK
    box = ink_box(arr)
    if not box:
        print('   空白页')
        return
    r0, r1, c0, c1 = box
    width = c1 - c0 + 1
    rows = dark[:, c0:c1 + 1].sum(axis=1)
    hits = np.where(rows >= cover * width)[0]
    print(f'=== {os.path.basename(path)}  版心宽 {width*cm:.1f}mm')
    if len(hits) == 0:
        print(f'   未发现覆盖率 ≥{cover:.0%} 的横线')
        return
    groups = []
    s = p = hits[0]
    for r in hits[1:]:
        if r - p > 2:
            groups.append((s, p))
            s = r
        p = r
    groups.append((s, p))
    print('   横线（行号 / 厚度 / 距版心顶）:')
    for (s, e) in groups:
        thick_mm = (e - s + 1) * cm
        print(f'      rows {s:4d}-{e:4d}  {thick_mm:4.2f}mm ≈ {thick_mm/MM_PER_PT:4.2f}pt'
              f'   @ {(s-r0)*cm:6.1f}mm')
    if len(groups) == 3:
        print('   → 3 条 = 标准三线表：第 1 条顶线(1.5pt)、'
              '第 2 条表头下(0.75pt)、第 3 条底线(1.5pt)')
    return groups


def expand(paths):
    out = []
    for p in paths:
        if any(ch in p for ch in '*?['):
            out.extend(sorted(glob.glob(p)))
        else:
            out.append(p)
    return out


def main():
    ap = argparse.ArgumentParser(description='从参考论文截图反推版面参数')
    ap.add_argument('mode', choices=['page', 'lines', 'table', 'all'])
    ap.add_argument('images', nargs='+')
    ap.add_argument('--page-width-mm', type=float, default=210.0,
                    help='图像宽度对应的实际纸宽（默认 A4 210mm）')
    ap.add_argument('--gap', type=int, default=4,
                    help='行切割的空白容差像素（文字行间距大时要调大）')
    ap.add_argument('--block', type=int, default=None, help='只看第 N 个文字块')
    args = ap.parse_args()

    for path in expand(args.images):
        if not os.path.exists(path):
            print(f'!! 找不到 {path}', file=sys.stderr)
            continue
        if args.mode in ('page', 'all'):
            cmd_page(path, args.page_width_mm)
        if args.mode in ('lines', 'all'):
            cmd_lines(path, args.page_width_mm, gap=args.gap, which=args.block)
        if args.mode in ('table', 'all'):
            cmd_table(path, args.page_width_mm)
        if args.mode != 'all':
            print()


if __name__ == '__main__':
    main()
