#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_gaps.py —— 留白页检查（2026-10-22，用户截图速查第 23 页：「下面这部分上面的空白区域放不下吗？为什么要再下一页重新开始？」）
用法：python3 check_gaps.py <成品.pdf> [--max 0]
找出：本页页底留白 ≥ 15%，下一页开头那一块（条目标题 / 要点 + 第一张表 + 表后紧跟的注 / 公司差异）按实测高度放得进本页留白——
本不该移到下一页。下一页以主题 / 章标题（≥ 11.5pt）开头的另起页属正常，不报。有问题时退出码 1（--max 允许的处数）。"""
import sys, re, pymupdf

PDF = sys.argv[1]
MAXOK = int(sys.argv[sys.argv.index('--max') + 1]) if '--max' in sys.argv else 0
d = pymupdf.open(PDF); H = d[0].rect.height; TOP, BOT = 36, H - 50
FOOT = re.compile(r'第\s*\d+\s*页\s*$')


def blocks(p):
    return [b for b in p.get_text('blocks') if b[4].strip() and not FOOT.search(b[4].strip()) and TOP <= b[1] < BOT]


out = []
for i in range(len(d) - 1):
    bs = blocks(d[i])
    ims = [x['bbox'] for x in d[i].get_image_info() if (x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 4000]
    last = max([b[3] for b in bs] + [r[3] for r in ims] + [TOP])
    gap = BOT - 8 - last
    if gap < 0.15 * (BOT - TOP): continue
    p2 = d[i + 1]
    try: ts = [t for t in p2.find_tables().tables if t.bbox[3] - t.bbox[1] > 10]
    except Exception: ts = []
    if not ts: continue
    t = ts[0]; b2 = blocks(p2)
    first = min([b[1] for b in b2] + [t.bbox[1]])
    if t.bbox[1] - first > 90: continue                       # 下一页开头不是这张表（前面还有正文）
    big = [s for b in p2.get_text('dict')['blocks'] for l in b.get('lines', []) for s in l['spans']
           if TOP <= l['bbox'][1] < t.bbox[1] and s['size'] >= 11.5 and s['text'].strip()]
    if big: continue                                          # 主题 / 章另起页
    end = t.bbox[3]
    for r in [x['bbox'] for x in p2.get_image_info() if (x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 4000]:
        if r[1] < t.bbox[3]: end = max(end, r[3] + 18); first = min(first, r[1])   # 图表并排：图（连图注）也算进这一块
    for b in sorted(b2, key=lambda b: b[1]):
        if b[1] >= end - 1 and b[1] - end < 14 and re.match(r'^\s*(注|公司差异)', b[4]): end = b[3]
    need = end - first + 6
    if need <= gap * 0.95:   # 留 5% 余量（实测误差；fit_fix 拉回要 3%）
        out.append((i + 1, round(gap), i + 2, round(need), re.sub(r'\s+', '', p2.get_text(clip=t.bbox))[:24]))
for o in out: print('第 %d 页留白 %dpt，第 %d 页开头那块 %dpt 放得下：%s' % o)
print('留白页检查：%d 处（允许 %d）' % (len(out), MAXOK))
sys.exit(1 if len(out) > MAXOK else 0)
