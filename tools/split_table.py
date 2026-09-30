#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""split_table.py —— 把一张大表按行断成几张（SD-86，2026-09-30 用户：「对于明显表格大内容多的表格，可以合理拆分」）
作为模块用：split(text, heading, starts) → 新文本
  heading：表格所在条目的标题行开头（如 "### B-1"），取其后第一张表；
  starts：新表从第几行数据开始（数据行从 1 数，表头 / 前提行不计），如 [4, 8] 表示断成 3 张。
规则：
  · 文字一字不改；前提行（premise）留在第一张；表末通栏注解行（note / warn）跟最后一张；
  · 每张新表重复原表头行（含 class 标注）；
  · 断点不能落在纵向合并格（rowspan）中间——落在中间就报错，不改；
  · 表后「注 / 出处 / 解释 / 公司差异：<行名>——…」段落按行名跟随所在那张表，相对顺序不变；认不出的留在最后一张之后。
命令行：python3 split_table.py <md> "<heading>" 4 8 --dry  只打印结果行数核对。"""
import re, sys

def _rows(tbl):
    return [m for m in re.finditer(r'(?s)<tr([^>]*)>.*?</tr>', tbl)]

def split(text, heading, starts):
    a = text.index(heading)
    t0 = text.index('<table', a); t1 = text.index('</table>', t0)
    open_tag = text[t0:text.index('>', t0) + 1]
    tbl = text[t0:t1]
    rows = _rows(tbl)
    cls = lambda m: (re.search(r'class="([^"]+)"', m.group(1)) or [None, ''])[1]
    hdr = [m for m in rows if 'hdr' in cls(m)]
    if not hdr: raise ValueError('没有表头行')
    hdr_html = hdr[0].group(0)
    data = [m for m in rows if not re.search(r'hdr|premise|note|warn', cls(m))]
    tail = [m for m in rows if re.search(r'note|warn', cls(m)) and m.start() > data[-1].start()]
    # rowspan 覆盖检查：每个数据行被上方合并格覆盖到第几行
    cover = [0] * (len(data) + 1)
    for i, m in enumerate(data, 1):
        for rs in re.findall(r'rowspan="(\d+)"', m.group(0)):
            for j in range(i + 1, min(len(data), i + int(rs) - 1) + 1): cover[j] = 1
    for st in starts:
        if not 2 <= st <= len(data): raise ValueError('断点 %d 超出范围（数据行 1～%d）' % (st, len(data)))
        if cover[st]: raise ValueError('断点 %d 落在合并格中间' % st)
    cuts = [data[st - 1].start() for st in starts]
    pieces, prev = [], 0
    for c in cuts:
        pieces.append(tbl[prev:c]); prev = c
    pieces.append(tbl[prev:])
    out = pieces[0].rstrip() + '\n</table>\n'
    for p in pieces[1:]:
        out += '\n\n' + open_tag + '\n' + hdr_html + '\n' + p.strip('\n') + ('\n' if p is not pieces[-1] else '')
        if p is not pieces[-1]: out = out.rstrip() + '\n</table>\n'
    res = text[:t0] + out.rstrip('\n') + ('\n' if not out.endswith('\n') else '') + text[t1:]
    return _route_notes(res, t0, len(starts) + 1)

def _route_notes(text, t0, n):
    """表后「注 / 出处 / 解释 / 公司差异：<行名>——…」段落跟随行名所在的那张表（相对顺序不变）；找不到或不唯一的留在最后一张后面。"""
    plain = lambda t: re.sub(r'\s', '', re.sub(r'<[^>]+>', '', t))
    ends, pos = [], t0
    for _ in range(n):
        e = text.index('</table>', pos) + len('</table>'); ends.append(e); pos = e
    starts_ = [text.rfind('<table', 0, e) for e in ends]
    nxt = re.search(r'(?m)^#{1,6} ', text[ends[-1]:])
    tail_end = ends[-1] + (nxt.start() if nxt else len(text) - ends[-1])
    tail = text[ends[-1]:tail_end]
    paras = re.split(r'\n\s*\n', tail.strip('\n'))
    NOTE = re.compile(r'^\s*(注|出处|解释|公司差异)[：:]\s*(.+?)——')
    pieces = [text[starts_[i]:ends[i]] for i in range(n)]
    firstcol = [' '.join(plain(m) for m in re.findall(r'(?s)<tr[^>]*>\s*<t[dh][^>]*>(.*?)</t[dh]>', pc)) for pc in pieces]
    buckets = [[] for _ in range(n)]
    for para in paras:
        m = NOTE.match(re.sub(r'<[^>]+>', '', para)); k = n - 1
        if m:
            core = plain(re.sub(r'（[^）]*）|\([^)]*\)', '', m.group(2))) or plain(m.group(2))
            hit = [i for i in range(n) if core and core in firstcol[i]] or [i for i in range(n) if core and core in plain(pieces[i])]
            if not hit:   # 再按最长共同字串：先比首列（权重 2），再比整张表；最高分唯一才移动
                full = plain(m.group(2))
                def lcs(a, b):
                    best = 0
                    for i0 in range(len(a)):
                        for j0 in range(i0 + best + 1, len(a) + 1):
                            if a[i0:j0] in b: best = j0 - i0
                            else: break
                    return best
                an = lambda t: ''.join(ch for ch in t if ch.isalnum())   # 只比字母数字，标点不截断共同字串
                full = an(full)
                sc = [max(2 * lcs(full, an(firstcol[i])), lcs(full, an(plain(pieces[i])))) for i in range(n)]
                top = max(sc)
                if top >= 4 and sc.count(top) == 1: hit = [sc.index(top)]
            if hit: k = hit[0]
        buckets[k].append(para)
    out = text[:t0]
    for i in range(n):
        out += pieces[i] + '\n'
        if buckets[i]: out += '\n' + '\n\n'.join(buckets[i]) + '\n'
        if i < n - 1: out += '\n\n'
    return out + '\n\n' + text[tail_end:].lstrip('\n') if text[tail_end:].strip() else out

if __name__ == '__main__':
    f, h = sys.argv[1], sys.argv[2]
    st = [int(x) for x in sys.argv[3:] if x.isdigit()]
    s = open(f, encoding='utf-8').read()
    n = split(s, h, st)
    plain = lambda t: re.sub(r'\s', '', re.sub(r'<[^>]+>', '', t))
    print('原表 → %d 张；去标签后字数 %d → %d（多出的应只是重复的表头）' % (len(st) + 1, len(plain(s)), len(plain(n))))
    if '--dry' not in sys.argv: open(f, 'w', encoding='utf-8').write(n)
