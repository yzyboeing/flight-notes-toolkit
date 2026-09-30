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

def lonely(pdf):
    """近乎空白的页（正文止于版面 25% 以内，SD-87）：
         lead：本页只有标题 / 导语，下一页一开头就是表（整表同页把表推走了）→ 返回下一页那张表；
         tail：本页只有上一页表格的表后说明（表格占满上一页）→ 返回上一页最后那张表。
       over＝（标题 / 说明高度 + 表高）÷ 版心高度，fit_fix 据此决定压缩级别；块索引表改为允许按块分页。"""
    d = pymupdf.open(pdf); H = d[0].rect.height
    info = []
    for p in d:
        bl = [b for b in p.get_text('blocks') if b[4].strip() and not re.match(r'\s*第\s*\d+\s*页', b[4]) and b[1] > 0.06 * H]
        try: tabs = [t for t in p.find_tables().tables if t.bbox[1] > 0.07 * H and t.bbox[3] - t.bbox[1] > 10]
        except Exception: tabs = []
        info.append((bl, tabs))
    spans = [max(b[3] for b in bl) - min(b[1] for b in bl) for bl, _ in info if bl]
    AREA = max(spans) if spans else H
    out = []
    for i in range(1, len(d) - 1):
        bl, tabs = info[i]
        if not bl or tabs: continue
        top, bot = min(b[1] for b in bl), max(b[3] for b in bl)
        if bot > 0.25 * H: continue
        nbl, ntabs = info[i + 1]
        HEAD = re.compile(r'^(\d\.\d+\u3000|[A-Z]-\d+\u3000|\d+\.\s|块索引\s*$)')
        if any(HEAD.match(b[4].strip()) for b in bl) and nbl:
            ntop = min(b[1] for b in nbl)
            first = min(nbl, key=lambda b: b[1])[4].strip()
            if re.match(r'^\d\.\d+\u3000', first): continue   # 下一页另起新节：本页只是上一节的末尾，正常
            t = min(ntabs, key=lambda t: t.bbox[1]) if ntabs else None
            if t is not None and t.bbox[1] <= ntop + 20:
                th, clip = t.bbox[3] - t.bbox[1], t.bbox
            else:   # 块索引等无竖线的表 PDF 认不出：取下一页顶部一段文字来匹配，高度按一整页估
                th, clip = AREA, pymupdf.Rect(0, ntop, d[i + 1].rect.width, ntop + 0.25 * H)
            out.append({'kind': 'lead', 'page': i + 1, 'over': round(((bot - top) + th + 12) / AREA, 3),
                        'table': norm(d[i + 1].get_text(clip=clip))})
            continue
        pbl, ptabs = info[i - 1]
        if ptabs and pbl:
            t = max(ptabs, key=lambda t: t.bbox[3])
            if t.bbox[3] >= max(b[3] for b in pbl) - 30:
                out.append({'kind': 'tail', 'page': i + 1, 'over': round(((t.bbox[3] - t.bbox[1]) + (bot - top) + 12) / AREA, 3),
                            'table': norm(d[i - 1].get_text(clip=t.bbox))})
    return out

def sparse(pdf):
    """只有一两行的页（2026-09-30 用户：避免一页只有一两行，SD-89）：返回上一页最后一张表的文字，
       fit_fix 把那张表收紧一级，腾出空间把这一两行拉回上一页。"""
    d = pymupdf.open(pdf); H = d[0].rect.height; out = []
    for i in range(1, len(d) - 1):
        ln = [l for l in lines_of(d[i]) if l[1] > 0.06 * H and not re.match(r'\s*第\s*\d+\s*页\s*$', l[4])]
        if not (0 < len(ln) <= 2): continue
        try: tabs = [t for t in d[i - 1].find_tables().tables if t.bbox[1] > 0.07 * H and t.bbox[3] - t.bbox[1] > 10]
        except Exception: tabs = []
        if not tabs: continue
        t = max(tabs, key=lambda t: t.bbox[3])
        out.append({'page': i + 1, 'table': norm(d[i - 1].get_text(clip=t.bbox))})
    return out
