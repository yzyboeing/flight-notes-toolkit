#!/usr/bin/env python3
"""verify_pages.py —— 写入后逐页读回核对（2026-10-06）：全文（含表格）逐字、每格对齐与圆点、表宽。
用法：python3 verify_pages.py --docx <基线docx> [--chapters 1-5]
对比对象：同一基线的转换结果（onenote_conv.convert）。插图图注只在读回里有，比对时去掉。"""
import argparse, html, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from onenote_api import content, list_pages, req
from onenote_build import split, locate, NB_ID, CHN, NUM
from onenote_conv import convert, strip_index, flat_lists
Z = r'[\s​⁠­ 　]+'
def plain(h): return re.sub(Z, '', html.unescape(re.sub(r'<[^>]+>', '', h)))
def cells(h):
    h = flat_lists(h)          # 2026-10-09 圆点改 OneNote 自带列表：列表还原成段落再比
    out = []
    for tb in re.findall(r'<table\b.*?</table>', h, re.S):
        for ta, td in re.findall(r'<td\b([^>]*)>(.*?)</td>', tb, re.S):
            ps = re.findall(r'<p\b([^>]*)>(.*?)</p>', td, re.S) or [(ta, td)]   # 读回时单段格子不带 <p>，对齐写在 td 上
            for a, p in ps:
                t = plain(p)
                if not t: continue
                al = (re.search(r'text-align:\s*(\w+)', a) or [0, 'left'])[1]
                out.append((t[:12], al if al in ('center', 'right') else 'left', '•' in p or '●' in p))
    return out
def widths(h):   # 每张表首行各格宽度之和（OneNote 把合并格拆成单格，格数会变，总宽不变；读回为 width:162，无 px）
    return [sum(int(x) for x in re.findall(r'<td\b[^>]*?width:(\d+)', re.search(r'<tr\b.*?</tr>', tb, re.S).group(0)))
            for tb in re.findall(r'<table\b.*?</table>', h, re.S) if '<tr' in tb]
ap = argparse.ArgumentParser(); ap.add_argument('--docx', required=True); ap.add_argument('--chapters', default='1-5'); a = ap.parse_args()
lo, hi = a.chapters.split('-'); want = [NUM[str(n)] for n in range(int(lo), int(hi) + 1)]
secs, _ = locate(NB_ID); bad = 0
for p in [p for p in split(a.docx) if p['chap'] in want]:
    exp = convert(strip_index(p['html'])[0])
    sid = secs[CHN[p['chap']]]
    got = sorted([q for q in req('GET', '/sections/%s/pages?$select=id,title,createdDateTime&$top=100' % sid)['value'] if q['title'] == p['title']],
                 key=lambda q: q['createdDateTime'], reverse=True)          # 同名取最新建的可读页
    h = next((x for x in (content(q['id']) for q in got) if x), None)
    if not h: print('!! 读不到', p['title']); bad += 1; continue
    body = re.sub(r'<img\b[^>]*>\s*<p\b[^>]*>.*?</p>', '', h, flags=re.S)
    body = body[body.find('<body'):]
    e, g = plain(exp), plain(body)
    g = g.replace(plain(CHN[p['chap']] + ' ｜ ' + p['title']), '', 1)
    probs = []
    if e not in g and g != e: probs.append('文字不同（期望 %d 字，读回 %d 字）' % (len(e), len(g)))
    ce, cg = cells(exp), cells(body)
    if ce != cg:
        d = next(((i, x, y) for i, (x, y) in enumerate(zip(ce, cg)) if x != y), ('长度', len(ce), len(cg)))
        probs.append('对齐/圆点不同 %s' % (d,))
    we, wg = widths(exp), widths(body)
    if len(we) != len(wg) or any(abs(x - y) > 3 for x, y in zip(we, wg)): probs.append('表宽不同 %s' % [(x, y) for x, y in zip(we, wg) if abs(x - y) > 3][:3] if len(we) == len(wg) else '表数不同（%d / %d）' % (len(we), len(wg)))
    print(('!! ' if probs else 'OK ') + p['title'], '；'.join(probs)); bad += bool(probs)
print('核对完：问题页', bad)
