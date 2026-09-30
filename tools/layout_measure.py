#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""layout_measure.py —— 在成品 PDF 上实测「该加宽的格子」（SD-85，2026-09-30）
供 fit_fix.py 调用，也可单独运行：python3 layout_measure.py <全书.pdf>
找两类格子，给出所在表格的文字（用于匹配源表签名）、列号、需要加宽多少（pt）：
  · 末行孤字：3 行以内的短格，折行后末行只剩 1～2 个字（如「警 / 戒」），上一行接近撑满（自然折行）；
  · 短格折行：约 20 字以内的表头或短格自然折成 2 行，一行放得下。
原文 <br> 主动换行（上一行明显短于格宽）、括注行、分条（•）格不管。只读。"""
import sys, re, json, collections
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

PAD = 12   # 单元格左右内边距合计（pt）
norm = lambda x: ''.join(ch for ch in str(x or '') if ch.isalnum())

def lines_of(pg):
    out = []
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans'])
            if t.strip(): out.append((*l['bbox'], t))
    return out

def cell_lines(lines, bb):
    got = [l for l in lines if bb[0] - 1 <= (l[0] + l[2]) / 2 <= bb[2] + 1 and bb[1] - 1 <= (l[1] + l[3]) / 2 <= bb[3] + 1]
    rows = collections.OrderedDict()
    for l in sorted(got, key=lambda z: (round(z[1] / 3), z[0])):
        k = round(((l[1] + l[3]) / 2) / 3)
        r = rows.setdefault(k, [l[0], l[2], ''])
        r[0] = min(r[0], l[0]); r[1] = max(r[1], l[2]); r[2] += l[4]
    return [tuple(v) for v in rows.values()]

def measure(pdf):
    d = pymupdf.open(pdf); H = d[0].rect.height; found = []
    for i, p in enumerate(d):
        try: tabs = p.find_tables().tables
        except Exception: continue
        lines = lines_of(p)
        for t in tabs:
            if t.bbox[3] - t.bbox[1] < 8 or t.bbox[1] < 0.08 * H: continue
            ttext = norm(p.get_text(clip=t.bbox))
            for r in t.rows:
                for k, c in enumerate(r.cells):
                    if not c or (k + 1 < len(r.cells) and r.cells[k + 1] is None): continue   # 跨列格不处理
                    ls = cell_lines(lines, c)
                    if len(ls) < 2 or len(ls) > 3 or ls[0][2].lstrip().startswith('•'): continue
                    cw = c[2] - c[0]
                    prev = ls[-2]
                    if (prev[1] - prev[0]) < cw - PAD - 18 or re.match(r'\s*[（(]', ls[-1][2]): continue   # 原文主动换行
                    last = re.sub(r'[\s，。；：、（）()「」.,;:]', '', ls[-1][2])
                    txt = ''.join(l[2] for l in ls).strip()
                    extra = 0
                    if 0 < len(last) <= 2:                     # 末行孤字：把末行摊到前面各行
                        extra = (ls[-1][1] - ls[-1][0]) / (len(ls) - 1) + 3
                    elif len(ls) == 2 and len(re.sub(r'\s', '', txt)) <= 20:   # 短格折两行：给够一行
                        extra = sum(l[1] - l[0] for l in ls) + PAD + 3 - cw
                    if extra > 0:
                        found.append({'page': i + 1, 'col': k, 'extra_pt': round(extra, 1), 'cell': txt[:24], 'first': norm(ls[0][2]), 'cx': round((c[0] + c[2]) / 2 - t.bbox[0], 1), 'tw': round(t.bbox[2] - t.bbox[0], 1), 'table': ttext})
    return found

if __name__ == '__main__':
    res = measure(sys.argv[1])
    print('共 %d 格需要加宽' % len(res))
    for x in res: print('  第 %d 页 第 %d 列 +%.0fpt「%s」' % (x['page'], x['col'] + 1, x['extra_pt'], x['cell']))
