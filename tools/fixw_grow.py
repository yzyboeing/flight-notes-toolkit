#!/usr/bin/env python3
"""fixw_grow.py —— 定宽表（表头写 w-NN）有空位却折行（2026-10-10 用户：速查第 23 条「N1 指示与游标定义第二列可以适当加宽以减少行数，为什么漏掉了」）
T13 跳过定宽表、fit_fix 不改定宽表、自动加宽不动定宽表——这里专查：表合计宽度 < 100%、某列真折行、表右侧空位够它省行（layout_measure 的 grow 项）。
用法：python3 fixw_grow.py --repo <gh-private> [--apply]
  不带 --apply：列出；带 --apply：把该列 w-NN 加上所需百分比（合计不超过 100），写回源文件（notes_src / 速查源）。
定位：表头文字 + 首行前几个字在源文件里唯一对上才改，对不上的只列出。"""
import argparse, glob, io, math, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from layout_measure import measure
CW = 770.0
N = lambda x: re.sub(r'<[^>]+>|[\s​⁠­]+|&[a-z]+;|[●•–]', '', x or '')

def fixed_tables(srcs):
    out = []
    for f in srcs:
        t = io.open(f, encoding='utf-8').read()
        for m in re.finditer(r'<table\b.*?</table>', t, re.S):
            tb = m.group(0); h = re.search(r'<tr class="hdr">(.*?)</tr>', tb, re.S)
            if not h or not re.search(r'\bw-\d+', h.group(1)): continue
            r1 = re.search(r'<tr>(.*?)</tr>', tb[h.end():], re.S)
            out.append({'f': f, 'pos': m.start() + h.start(), 'h': N(h.group(1)), 'r': N(r1.group(1) if r1 else '')[:4], 'hdr': h.group(0)})
    return out

def scan(repo, books=None):
    hits = {}
    books = books or (('全书', 'build/B737机型理论知识笔记.pdf', sorted(glob.glob(os.path.join(repo, 'notes_src/[1-5]*/*.md')))),
                      ('单册', 'build/B737机型理论知识速查.pdf', [os.path.join(repo, '速查/速查源.md')]))
    for name, pdf, srcs in books:
        pdf = os.path.join(repo, pdf)
        if not os.path.exists(pdf): continue
        ft = fixed_tables(srcs)
        for x in measure(pdf):
            if not x.get('grow') or x.get('fill'): continue
            tb = N(x.get('table', ''))
            c = [t for t in ft if tb.startswith(t['h']) and tb[len(t['h']):].startswith(t['r'])]
            if len(c) != 1: continue
            k = (c[0]['f'], c[0]['pos'])
            v = hits.setdefault(k, {'book': name, 'page': x['page'], 'hdr': c[0]['hdr'], 'cols': {}})
            v['cols'][x['col']] = max(v['cols'].get(x['col'], 0), x['extra_pt'])
    return hits

def apply(hits):
    done = 0
    for (f, pos), v in sorted(hits.items(), key=lambda kv: (kv[0][0], -kv[0][1])):
        t = io.open(f, encoding='utf-8').read()
        if t[pos:pos + len(v['hdr'])] != v['hdr']: print('  对不上，跳过', f, v['page']); continue
        ths = re.findall(r'<th\b[^>]*>', v['hdr']); ws = [int((re.search(r'\bw-(\d+)', a) or [0, 0])[1]) for a in ths]
        room = 100 - sum(ws)
        for c, e in sorted(v['cols'].items(), key=lambda kv: -kv[1]):
            if c >= len(ws) or not ws[c] or room <= 0: continue
            add = min(room, math.ceil(e / CW * 100) + 1); ws[c] += add; room -= add
        new = v['hdr']; k = 0
        def sub(m):
            nonlocal k
            a = m.group(0); r = re.sub(r'\bw-\d+\b', 'w-%d' % ws[k], a) if re.search(r'\bw-\d+', a) else a; k += 1; return r
        new = re.sub(r'<th\b[^>]*>', sub, new)
        if new != v['hdr']:
            io.open(f, 'w', encoding='utf-8').write(t[:pos] + new + t[pos + len(v['hdr']):]); done += 1
            print('  %s 第 %d 页 %s → %s' % (v['book'], v['page'], os.path.basename(f), re.findall(r'w-\d+', new)))
    return done

if __name__ == '__main__':
    ap = argparse.ArgumentParser(); ap.add_argument('--repo', default='.'); ap.add_argument('--apply', action='store_true'); a = ap.parse_args()
    h = scan(a.repo)
    for (f, pos), v in h.items():
        print('%s 第 %d 页 %s：%s' % (v['book'], v['page'], os.path.basename(f), '；'.join('第 %d 列可加宽约 %.0fpt' % (c + 1, e) for c, e in v['cols'].items())))
    if a.apply: print('已改 %d 张' % apply(h))
