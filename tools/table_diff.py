#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""table_diff.py —— 两版成品逐表逐列比对（2026-10-22，1022R1 回退教训：对齐 / 圆点 / 列宽被新规则改坏，靠用户截图才发现）
用法：python3 table_diff.py <旧.docx> <新.docx> [--md 输出.md] [--width 5]
按「表头＋首行」文字配对两版的表，逐列比较：样式（圆点 / 居中 / 靠左）、列宽占比（默认差 ≥ 5 个百分点才报）。
只列出变化，供出版前逐条核实「是不是用户要的」；对不上的表单独列出。"""
import sys, re, zipfile, collections

OLD, NEW = sys.argv[1], sys.argv[2]
OUT = sys.argv[sys.argv.index('--md') + 1] if '--md' in sys.argv else None
WTH = float(sys.argv[sys.argv.index('--width') + 1]) if '--width' in sys.argv else 5.0
N = lambda s: re.sub(r'[^\w]', '', s)


def tables(path):
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8')
    out = []
    for m in re.finditer(r'<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>', x, re.S):
        tb = m.group(0)
        grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tb)]
        if len(grid) < 2: continue
        rows = []
        for tr in re.findall(r'<w:tr\b.*?</w:tr>', tb, re.S):
            col, cells = 0, []
            for tc in re.findall(r'<w:tc>.*?</w:tc>', tr, re.S):
                sp = re.search(r'<w:gridSpan w:val="(\d+)"', tc); n = int(sp.group(1)) if sp else 1
                txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tc))
                jc = re.findall(r'<w:jc w:val="(\w+)"', tc)
                cells.append((col, n, txt, '●' in txt, bool(jc) and all(j == 'center' for j in jc)))
                col += n
            rows.append(cells)
        if len(rows) < 2: continue
        head = '|'.join(c[2] for c in rows[0])
        first = ''.join(c[2] for c in rows[1])
        st = {}
        for k in range(len(grid)):
            cs = [c for r in rows[1:] for c in r if c[0] == k and c[1] == 1 and N(c[2].replace('●', ''))]
            if not cs: continue
            b = sum(1 for c in cs if c[3]); ce = sum(1 for c in cs if c[4] and not c[3])
            st[k] = '圆点' if b >= max(1, len(cs) / 2) else ('居中' if ce >= len(cs) / 2 else '靠左')
        tot = sum(grid)
        out.append({'key': N(head)[:40] + '#' + N(first.replace('●', ''))[:16], 'head': head, 'grid': [round(100 * g / tot) for g in grid], 'st': st})
    return out


a, b = tables(OLD), tables(NEW)
ma = collections.defaultdict(list)
for t in a: ma[t['key']].append(t)
diffs, miss = [], 0
for t in b:
    if not ma.get(t['key']): miss += 1; continue
    o = ma[t['key']].pop(0)
    hs = t['head'].split('|')
    for k in sorted(set(o['st']) | set(t['st'])):
        if o['st'].get(k) != t['st'].get(k) and o['st'].get(k) and t['st'].get(k):
            diffs.append(('样式', t['head'][:40], hs[k] if k < len(hs) else k, o['st'][k], t['st'][k]))
    if len(o['grid']) == len(t['grid']):
        for k, (x, y) in enumerate(zip(o['grid'], t['grid'])):
            if abs(x - y) >= WTH: diffs.append(('列宽', t['head'][:40], hs[k] if k < len(hs) else k, '%d%%' % x, '%d%%' % y))
lines = ['# 两版逐表逐列比对', '', '旧：%s（%d 张表）　新：%s（%d 张表）；配对上 %d 张，新版里对不上（内容改过或新增）%d 张；变化 %d 处。' % (OLD.split('/')[-1], len(a), NEW.split('/')[-1], len(b), len(b) - miss, miss, len(diffs)), '',
         '| 类别 | 表头 | 列 | 旧 | 新 |', '|---|---|---|---|---|'] + ['| %s | %s | %s | %s | %s |' % d for d in diffs]
txt = '\n'.join(lines) + '\n'
if OUT: open(OUT, 'w', encoding='utf-8').write(txt)
print(lines[2]); print('样式变化 %d、列宽变化 %d' % (sum(d[0] == '样式' for d in diffs), sum(d[0] == '列宽' for d in diffs)))
