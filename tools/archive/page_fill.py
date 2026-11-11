#!/usr/bin/env python3
"""page_fill.py —— 版面利用率统计（2026-10-21 跨页表格方案评估用；出版物标准：少留白、不乱拆）。
用法：python3 page_fill.py <成品.pdf> [--list]
口径：每页正文区（页眉下到页脚上）里，最后一块内容（文字 / 插图 / 表格线）以下的空白占正文区的比例。
  · 下一页以章 / 节 / 速查主题标题开头的页（规则要求另起一页）单列「节末页」，不算浪费；
  · 其余页底空白 ≥ 30% 记「页底留白页」，≥ 60% 记「半空页」。"""
import sys, re
import pymupdf

d = pymupdf.open(sys.argv[1]); H = d[0].rect.height
TOP, BOT = 36, H - 50   # 版心大致上下界（横版 A4：页边 0.5in、页脚约 36pt）


def bottom(p):
    ys = [b[3] for b in p.get_text('blocks') if b[4].strip() and not re.search(r'第\s*\d+\s*页\s*$', b[4].strip())]
    ys += [x['bbox'][3] for x in p.get_image_info() if (x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 4000]
    ys += [dr['rect'].y1 for dr in p.get_drawings() if dr['rect'].y1 < BOT and dr['rect'].height < 400]
    return max(ys) if ys else TOP


def starts_section(p):
    """下一页第一行是章 / 节 / 速查主题标题（字号 ≥ 13pt，规则要求另起一页）"""
    for b in p.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            sp = [s for s in l['spans'] if s['text'].strip()]
            if not sp or re.search(r'第\s*\d+\s*页', ''.join(s['text'] for s in sp)): continue
            return max(s['size'] for s in sp) >= 13
    return True


tail, waste, half, lst = 0, 0.0, 0, []
for i in range(1, len(d) - 1):
    f = max(0.0, (BOT - bottom(d[i])) / (BOT - TOP))
    if starts_section(d[i + 1]): tail += 1; continue
    waste += f
    if f >= 0.3: lst.append((i + 1, round(f * 100)))
    if f >= 0.6: half += 1
print('%s：%d 页；节末页 %d；页底留白页（≥30%%）%d，其中半空（≥60%%）%d；非节末页留白合计 %.1f 页' % (
    sys.argv[1].split('/')[-1], len(d), tail, len(lst), half, waste))
if '--list' in sys.argv: print(' '.join('%d(%d%%)' % x for x in lst))
