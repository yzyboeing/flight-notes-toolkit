# -*- coding: utf-8 -*-
"""页脚页码跳回本页知识点所在的速查目录页（SD-236，2026-10-24 用户：「点击右下角页码返回主题目录，能不能设置为返回到页码对应知识点的主题目录那一页」）

Word 页脚按节（一章一个）设置，每页跳不同位置做不到，所以在 PDF 生成后逐页改页码链接的落点（只改链接、不重排，不影响分页）。
落点规则：
  · 本页有条目开头 → 取第一个开头、且列在速查目录里的条目，跳它在速查目录里的第一处主位置（不取 ↗ 交叉引用位置；只有 ↗ 时取第一处 ↗）；
  · 续页（本页没有条目开头）→ 跳正在续写的条目；它没列在速查目录时往前找同节条目；
  · 章首页 → 跳本章第一个列在速查目录里的条目；
  · 都找不到 → 仍跳速查目录第 1 页。
Word 版页码仍跳速查目录第 1 页；OneNote 没有页脚，不受影响。

用法：footer_toc_links.py <pdf> [--check]
  不带 --check：改链接并写回 PDF；带 --check：只核对，落点不符的报错误（退出码 1）。
"""
import sys, re, os
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

ENTRY = re.compile(r'^(\d+\.\d+)\s+(\d+)(↗?)')
PAGENO = re.compile(r'^第\s*\d+\s*页$')


def _txt(pg, r):
    return pg.get_text(clip=r).strip()


def load(d):
    toc = d.get_toc(simple=False)
    dir0 = next((e[2] for e in toc if e[0] == 1 and e[1].replace('​', '').strip() == '速查目录'), None)
    if not dir0: return None
    body0 = next(e[2] for e in toc if e[0] == 1 and e[2] > dir0)            # 速查目录之后第一个一级书签 = 正文第一章
    # 条目：二级书签「x.y 节名」下的三级书签「N. 标题」→ 编号 x.y N
    ents, sec, chap = [], None, None
    for lv, t, p, *_ in toc:
        if p < body0: continue
        t = t.replace('​', '').strip()
        if lv == 1: chap = p
        elif lv == 2:
            m = re.match(r'^(\d+\.\d+)', t); sec = m.group(1) if m else None
        elif lv == 3 and sec:
            m = re.match(r'^(\d+)\.', t)
            if m: ents.append({'id': '%s %s' % (sec, m.group(1)), 'sec': sec, 'page': p, 'chap': chap})
    # 速查目录：条目编号 → 主位置（目录页）
    loc, xref = {}, {}
    for i in range(dir0 - 1, body0 - 1):
        pg = d[i]
        for l in pg.get_links():
            if l.get('kind') != pymupdf.LINK_GOTO: continue
            m = ENTRY.match(_txt(pg, l['from']))
            if not m: continue                                             # 主题小标题、页码数字格不算
            k = '%s %s' % (m.group(1), m.group(2))
            (xref if m.group(3) else loc).setdefault(k, i + 1)
    for k, v in xref.items(): loc.setdefault(k, v)
    return dir0, body0, ents, loc


def targets(d):
    r = load(d)
    if not r: return None
    dir0, body0, ents, loc = r
    chaps = sorted({e['chap'] for e in ents})
    out = {}
    for p in range(body0, len(d) + 1):
        here = [e for e in ents if e['page'] == p and e['id'] in loc]
        if here: out[p] = (loc[here[0]['id']], '条目 ' + here[0]['id']); continue
        prev = [e for e in ents if e['page'] < p]
        if p in chaps or not prev or prev[-1]['chap'] != max(c for c in chaps if c <= p):   # 章首页
            nxt = [e for e in ents if e['page'] >= p and e['id'] in loc]
            if nxt: out[p] = (loc[nxt[0]['id']], '章首 ' + nxt[0]['id']); continue
        cont = prev[-1] if prev else None
        if cont:
            same = [e for e in prev if e['sec'] == cont['sec'] and e['id'] in loc]
            if same: out[p] = (loc[same[-1]['id']], '续页 ' + same[-1]['id']); continue
        out[p] = (dir0, '兜底')
    return dir0, out


def footer_links(pg):
    h = pg.rect.height
    return [l for l in pg.get_links() if l.get('kind') == pymupdf.LINK_GOTO and l['from'].y0 > h * 0.9
            and PAGENO.match(re.sub(r'\s+', ' ', _txt(pg, l['from'])))]


def main():
    pdf = sys.argv[1]; check = '--check' in sys.argv
    d = pymupdf.open(pdf)
    r = targets(d)
    if not r:
        print('页脚页码落点：没有速查目录书签，跳过'); return 0
    dir0, out = r
    epage = {e['id']: e['page'] for e in load(d)[2]}
    ids_on = {}                                   # 目录页 → 该页列着的条目编号（独立复核用）
    def ids(t):
        if t not in ids_on:
            tp = d[t - 1]
            ids_on[t] = {' '.join(m.group(1, 2)) for x in tp.get_links() if x.get('kind') == pymupdf.LINK_GOTO
                         for m in [ENTRY.match(_txt(tp, x['from']))] if m}
        return ids_on[t]
    bad, nolink, fb, ch = [], [], 0, 0
    for p, (t, why) in out.items():
        pg = d[p - 1]; fl = footer_links(pg)
        if not fl: nolink.append(p); continue
        fb += why == '兜底'
        for l in fl:
            if check:
                if l['page'] + 1 != t: bad.append((p, l['page'] + 1, t, why))
                elif why != '兜底':   # 独立复核：落点目录页上确实列着这个条目；「条目」类的条目确实在本页开头，续页 / 章首的在本页之前 / 之后
                    k = why.split(' ', 1)[1]; ep = epage.get(k)
                    if k not in ids(t) or (why.startswith('条目') and ep != p) or (why.startswith('续页') and not ep < p) \
                       or (why.startswith('章首') and not ep >= p): bad.append((p, l['page'] + 1, t, why + ' 复核未过'))
            elif l['page'] + 1 != t:
                # LibreOffice 的页脚链接是同一个注释对象挂在每一页上（改落点 / delete_link 都会全书一起变）——
                # 只从本页的 /Annots 里摘掉它（对象保留给别的页），再插一个本页独有的链接
                rect = l['from']
                k, v = d.xref_get_key(pg.xref, 'Annots')
                arr = d.xref_object(int(v.split()[0])) if k == 'xref' else v
                refs = [m for m in re.findall(r'\d+ 0 R', arr) if int(m.split()[0]) != l['xref']]
                d.xref_set_key(pg.xref, 'Annots', '[' + ' '.join(refs) + ']')
                pg = d.reload_page(pg)
                pg.insert_link({'kind': pymupdf.LINK_GOTO, 'from': rect, 'page': t - 1, 'to': pymupdf.Point(0, 0)}); ch += 1
    if check:
        for x in bad[:10]: print('  ✗ 第 %d 页页码跳第 %d 页，应跳第 %d 页（%s）' % x)
        for p in nolink[:5]: print('  ✗ 第 %d 页页脚没有页码链接' % p)
        print('页脚页码落点：正文 %d 页，不符 %d 页、缺链接 %d 页、兜底到速查目录首页 %d 页' % (len(out), len(bad), len(nolink), fb))
        return 1 if bad or nolink else 0
    d.save(pdf + '.tmp', garbage=3, deflate=True); d.close(); os.replace(pdf + '.tmp', pdf)
    print('页脚页码落点：改 %d 个链接（正文 %d 页，兜底 %d 页）' % (ch, len(out), fb))
    return 0


if __name__ == '__main__':
    sys.exit(main())
