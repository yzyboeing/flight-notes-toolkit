#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""autow.py —— 按成品实测给个别表定列宽比例（w-NN），用于自动列宽规则处理不好的表（2026-10-06）
用法：python3 autow.py <成品.pdf> <页码> <表头开头文字，可写「项目|左再循环风扇」> [--apply <源文件.md> [--anchor <表内文字>]]
      python3 autow.py --from-report [build/排版检查报告.md]   （2026-10-07）按报告里的 T13（列右侧空白收不掉）逐张定宽：
            自动在成品 PDF 找到那张表、量出列宽，按「表头 + 首行文字」在 notes_src / 速查源 里唯一定位后写入 w-NN；对不上或不唯一的只列出不改
做法：量出每列各格的行宽；没折行的列给「最长一行 + 内边距 + 余量」，其余宽度按文字量分给折行的列；
输出各列百分比。--apply 时把 w-NN 写进源文件里表头文字相同的那张表的 <th>（表头有跨列格的不处理）。
口径同 layout_measure：折行＝某行排满（左右空白合计 < 内边距 + 14pt）且下一行不是新的一条。"""
import sys, re, pymupdf
from layout_measure import forced_wrap   # 2026-10-07 被迫折行同口径

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

def measure_tb(p, tb):
    hdr = [c for c, x in zip(tb.rows[0].cells, tb.extract()[0]) if c and (x or '').strip()]   # 去掉表右侧虚构的空列（2026-10-06）
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
                if q + 1 < len(ls) and ((c[2] - c[0]) - (l[1] - l[0]) < PAD + 14 or forced_wrap(c[2] - c[0], l[1] - l[0], l[2], ls[q + 1][2], PAD)) and not re.match(NEWIT, ls[q + 1][2]): wrapped[k] = True
    for k in range(n):   # 表头也要放得下
        longest[k] = max(longest[k], max((l[1] - l[0] for l in lines_in(p, hdr[k])), default=0))
    want = [0.0] * n
    for k in range(n):
        if not wrapped[k]: want[k] = longest[k] + PAD + 8
    rest = CW - sum(want); wk = [k for k in range(n) if wrapped[k]]
    tot = sum(text[k] for k in wk) or 1
    for k in wk: want[k] = max(longest[k] * 0.45 + PAD, rest * text[k] / tot)
    # 按版心宽度的百分比写（合计不足 100 时生成器按版心百分比取宽，表格不会被拉满）；有折行列时合计拉到 100
    pct = [int(max(4, -(-100 * w // CW))) for w in want]
    if any(wrapped): pct[-1] += 100 - sum(pct)
    while sum(pct) > 100: pct[pct.index(max(pct))] -= 1
    return want, wrapped, pct

def apply_pct(s, m, pct):
    ths = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1), re.S)
    new = '<tr class="hdr">'
    for (attr, txt), v in zip(ths, pct):
        cm = re.search(r'class="([^"]*)"', attr)
        cls = re.sub(r'\bw-\d+(-\d+)?\b', '', cm.group(1) if cm else '').strip()
        new += '<th class="%s">%s</th>' % ((cls + ' w-%d' % v).strip(), txt)
    return s[:m.start()] + new + '</tr>' + s[m.end():]

N = lambda x: re.sub(r'<[^>]+>|\s+|&[a-z]+;', '', x or '')

def from_report():
    import glob, os
    rep = sys.argv[sys.argv.index('--from-report') + 1] if len(sys.argv) > sys.argv.index('--from-report') + 1 else 'build/排版检查报告.md'
    txt = open(rep, encoding='utf-8').read().split('## 新建议')[0]
    hits = re.findall(r'\[T13\] (全书|单册) 第 (\d+) 页：第 (\d+) 列（「([^」]*)」）[^\n]*?空约 (\d+)pt', txt)
    done = set()
    for book, pg, col, colh, spare in hits:
        pdf = 'build/B737机型理论知识笔记.pdf' if book == '全书' else 'build/B737机型理论知识速查.pdf'
        srcs = sorted(glob.glob('notes_src/[1-5]*/*.md')) if book == '全书' else ['速查/速查源.md']
        p = pymupdf.open(pdf)[int(pg) - 1]
        cand = [t for t in p.find_tables().tables if len(t.extract()[0]) >= int(col) and N((t.extract()[0][int(col) - 1] or '')).startswith(N(colh)[:6])]
        if not cand: print('  第 %s 页：找不到「%s」列所在的表' % (pg, colh)); continue
        def _spare(t):   # 这一列的实测右侧空白，与报告的「空约 N pt」最接近的那张就是
            c0 = [c for c in t.rows[0].cells if c][int(col) - 1]
            ls = [l for r in t.rows[1:] for c in r.cells if c and abs(c[0] - c0[0]) < 2 for l in lines_in(p, c)]
            return (c0[2] - c0[0]) - PAD - max((l[1] - l[0] for l in ls), default=0)
        tb = min(cand, key=lambda t: abs(_spare(t) - int(spare))); heads = [N(h) for h in tb.extract()[0] if N(h)]
        row1 = N(''.join(x or '' for x in tb.extract()[1])) if len(tb.extract()) > 1 else ''
        if (book, tuple(heads), row1[:20]) in done: continue
        done.add((book, tuple(heads), row1[:20]))
        want, wrapped, pct = measure_tb(p, tb)
        found = []
        for f in srcs:
            s = open(f, encoding='utf-8').read()
            for m in re.finditer(r'<tr class="hdr">(.*?)</tr>', s, re.S):
                ths = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1), re.S)
                if [N(t[1]) for t in ths] != heads or any('colspan' in t[0] for t in ths): continue
                first = re.search(r'(?s)<tr>(.*?)</tr>', s[m.end():s.find('</table>', m.end())])
                cells = [N(c) for c in re.findall(r'(?s)<td[^>]*>(.*?)</td>', first.group(1))] if first else []
                if cells and all(c[:6] in row1 for c in cells if c): found.append((f, m))
        if len(found) != 1: print('  第 %s 页「%s」：源文件里对上 %d 张表，未改' % (pg, '｜'.join(heads), len(found))); continue
        f, m = found[0]; s = open(f, encoding='utf-8').read()
        open(f, 'w', encoding='utf-8').write(apply_pct(s, m, pct))
        print('  第 %s 页「%s」→ %s：w-%s' % (pg, '｜'.join(heads), os.path.basename(f), '/'.join(map(str, pct))))

def main():
    if '--from-report' in sys.argv: return from_report()
    pdf, pg, head = sys.argv[1], int(sys.argv[2]), sys.argv[3]
    d = pymupdf.open(pdf); p = d[pg - 1]
    tb = next(t for t in p.find_tables().tables if '|'.join((x or '').replace('\n', '') for x in t.extract()[0]).startswith(head))   # head 可写「项目|左再循环风扇」区分同页同首列的表
    want, wrapped, pct = measure_tb(p, tb); total = sum(want)
    print('列宽 pt', [round(w) for w in want], '折行', wrapped, '百分比', pct, '表宽', round(total))
    if '--apply' in sys.argv:
        src = sys.argv[sys.argv.index('--apply') + 1]; s = open(src, encoding='utf-8').read()
        heads = [re.sub(r'\s+', '', (h or '')) for h in tb.extract()[0]]
        anchor = sys.argv[sys.argv.index('--anchor') + 1] if '--anchor' in sys.argv else None   # 表内一段文字，精确定位同表头的多张表中的那一张
        heads = [h for h in heads if h]
        for m in re.finditer(r'<tr class="hdr">(.*?)</tr>', s, re.S):
            ths = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1), re.S)
            if [re.sub(r'<[^>]+>|\s+', '', t[1]) for t in ths] != heads or any('colspan' in t[0] for t in ths): continue
            if anchor and anchor not in s[m.end():s.find('</table>', m.end())]: continue
            s = apply_pct(s, m, pct); open(src, 'w', encoding='utf-8').write(s); print('已写入', src); return
        print('源文件里找不到表头相同的表：', heads)

if __name__ == '__main__':
    main()
