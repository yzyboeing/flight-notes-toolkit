#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""autow.py —— 按成品实测给个别表定列宽比例（w-NN），用于自动列宽规则处理不好的表（2026-10-06）
用法：python3 autow.py <成品.pdf> <页码> <表头第一格文字> [--apply <源文件.md>]
做法：量出每列各格的行宽；没折行的列给「最长一行 + 内边距 + 余量」，其余宽度按文字量分给折行的列；
输出各列百分比。--apply 时把 w-NN 写进源文件里表头文字相同的那张表的 <th>（表头有跨列格的不处理）。
口径同 layout_measure：折行＝某行排满（左右空白合计 < 内边距 + 14pt）且下一行不是新的一条。"""
import sys, re, pymupdf

PAD, CW = 13, 770.0
NEWIT = r'\s*([●•▪·–\-①-⑳【]|[A-H]-\d|\d{1,2}[.、)）]|注[：:])'

def lines_in(p, bb):
    out = []
    for b in p.get_text('dict', clip=bb)['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans'])
            if t.strip(): out.append((l['bbox'][0], l['bbox'][2], (l['bbox'][1] + l['bbox'][3]) / 2, t))
    rows = {}
    for x0, x1, y, t in sorted(out, key=lambda z: (round(z[2] / 3), z[0])):
        k = round(y / 3); r = rows.setdefault(k, [x0, x1, '']); r[0] = min(r[0], x0); r[1] = max(r[1], x1); r[2] += t
    return [tuple(v) for v in rows.values()]

def main():
    pdf, pg, head = sys.argv[1], int(sys.argv[2]), sys.argv[3]
    d = pymupdf.open(pdf); p = d[pg - 1]
    tb = next(t for t in p.find_tables().tables if (t.extract()[0][0] or '').replace('\n', '').startswith(head))
    hdr = [c for c in tb.rows[0].cells if c]
    xs = [c[0] for c in hdr] + [hdr[-1][2]]
    n = len(hdr); longest, text, wrapped = [0.0] * n, [0.0] * n, [False] * n
    for r in tb.rows[1:]:
        for c in r.cells:
            if not c: continue
            cov = [q for q in range(n) if xs[q] >= c[0] - 2 and xs[q + 1] <= c[2] + 2]
            if len(cov) != 1: continue
            k = cov[0]; ls = lines_in(p, c)
            for q, l in enumerate(ls):
                longest[k] = max(longest[k], l[1] - l[0]); text[k] += l[1] - l[0]
                if q + 1 < len(ls) and (c[2] - c[0]) - (l[1] - l[0]) < PAD + 14 and not re.match(NEWIT, ls[q + 1][2]): wrapped[k] = True
    for k in range(n):   # 表头也要放得下
        longest[k] = max(longest[k], max((l[1] - l[0] for l in lines_in(p, hdr[k])), default=0))
    want = [0.0] * n
    for k in range(n):
        if not wrapped[k]: want[k] = longest[k] + PAD + 8
    rest = CW - sum(want); wk = [k for k in range(n) if wrapped[k]]
    tot = sum(text[k] for k in wk) or 1
    for k in wk: want[k] = max(longest[k] * 0.45 + PAD, rest * text[k] / tot)
    if not wk: rest = 0
    total = sum(want)
    pct = [max(4, round(100 * w / total)) for w in want]
    print('列宽 pt', [round(w) for w in want], '折行', wrapped, '百分比', pct, '表宽', round(total))
    if '--apply' in sys.argv:
        src = sys.argv[sys.argv.index('--apply') + 1]; s = open(src, encoding='utf-8').read()
        heads = [re.sub(r'\s+', '', (h or '')) for h in tb.extract()[0]]
        for m in re.finditer(r'<tr class="hdr">(.*?)</tr>', s, re.S):
            ths = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1), re.S)
            if [re.sub(r'<[^>]+>|\s+', '', t[1]) for t in ths] != heads or any('colspan' in t[0] for t in ths): continue
            new = '<tr class="hdr">'
            for (attr, txt), v in zip(ths, pct):
                cm = re.search(r'class="([^"]*)"', attr)
                cls = re.sub(r'\bw-\d+(-\d+)?\b', '', cm.group(1) if cm else '').strip()
                new += '<th class="%s">%s</th>' % ((cls + ' w-%d' % v).strip(), txt)
            new += '</tr>'
            s = s[:m.start()] + new + s[m.end():]; open(src, 'w', encoding='utf-8').write(s); print('已写入', src); return
        print('源文件里找不到表头相同的表：', heads)

if __name__ == '__main__':
    main()
