#!/usr/bin/env python3
"""verify_pages.py —— 写入后逐页读回核对（2026-10-06）：全文（含表格）逐字、每格对齐与圆点、表宽及边框。
用法：python3 verify_pages.py --docx <基线docx> [--chapters 1-6]
对比对象：同一基线的转换结果（onenote_conv.convert）。插图图注只在读回里有，比对时去掉。"""
import argparse, html, re
from html.parser import HTMLParser
from pathlib import Path
Z = r'[\s​⁠­ 　]+'
def plain(h): return re.sub(Z, '', html.unescape(re.sub(r'<[^>]+>', '', h)))
def cells(h):
    from onenote_conv import flat_lists
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


_BORDER = re.compile(r'(?:^|;)\s*border:\s*1px\s+solid(?=\s|;|$)', re.I)
_NOTE = re.compile(r'(?:^|;)\s*background-color:\s*#eef4fb(?=\s|;|$)', re.I)


def border_signature(markup, *, converted=False):
    """Read table and cell borders in document order, including nested tables."""
    tables = []
    stack = []

    class Borders(HTMLParser):
        def handle_starttag(self, tag, attrs):
            attrs = dict(attrs)
            style = attrs.get('style', '') or ''
            if tag == 'table':
                solid = bool(_BORDER.search(style) or (converted and attrs.get('border') == '1'))
                current = (solid, [], [])
                tables.append(current)
                stack.append(current)
            elif tag in ('td', 'th') and stack:
                _, cells, blue = stack[-1]
                cells.append(bool(_BORDER.search(style)))
                blue.append(bool(_NOTE.search(style)))

        def handle_endtag(self, tag):
            if tag == 'table' and stack:
                stack.pop()

    parser = Borders()
    parser.feed(markup)
    parser.close()
    return [(solid, tuple(cells), len(cells) == 1 and blue == [True]) for solid, cells, blue in tables]


def border_counts(signature):
    return {
        'tables': len(signature),
        'borderedTables': sum(solid for solid, _, _ in signature),
        'borderedCells': sum(sum(cells) for _, cells, _ in signature),
        'borderlessNotes': sum(note and not solid and not any(cells) for solid, cells, note in signature),
    }


def border_differences(expected, actual):
    """Return 1-based positions so a border moved to another table cannot pass."""
    return [i for i in range(1, max(len(expected), len(actual)) + 1)
            if expected[i - 1:i] != actual[i - 1:i]]


def main():
    from onenote_api import content, req
    from onenote_build import split, locate, NB_ID, CHN, NUM, sec_name
    from onenote_conv import convert, strip_index
    ap = argparse.ArgumentParser(); ap.add_argument('--docx', required=True); ap.add_argument('--chapters', default='1-6'); ap.add_argument('--border-baseline', type=Path, help='写前逐页 HTML 目录，文件名如 1.8.html'); a = ap.parse_args()
    lo, hi = a.chapters.split('-'); want = [NUM[str(n)] for n in range(int(lo), int(hi) + 1)]
    secs, _ = locate(NB_ID); bad = 0
    for p in [p for p in split(a.docx) if p['chap'] in want]:
        exp = convert(strip_index(p['html'])[0])
        sid = secs[sec_name(CHN[p['chap']], secs)]
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
        if a.border_baseline:
            key = p['title'].split()[0]
            expected_border = border_signature((a.border_baseline / f'{key}.html').read_bytes().decode('utf-8'))
        else:
            expected_border = border_signature(exp, converted=True)
        actual_border = border_signature(h)
        changed_borders = border_differences(expected_border, actual_border)
        if changed_borders: probs.append('边框不同 表序 %s，写前 %s，读回 %s' % (changed_borders, border_counts(expected_border), border_counts(actual_border)))
        print(('!! ' if probs else 'OK ') + p['title'], '；'.join(probs)); bad += bool(probs)
    print('核对完：问题页', bad)
    return int(bool(bad))


if __name__ == '__main__':
    raise SystemExit(main())
