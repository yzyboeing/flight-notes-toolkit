#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""layout_measure.py —— 在成品 PDF 上实测「该加宽的格子」（SD-85，2026-09-30）
供 fit_fix.py 调用，也可单独运行：python3 layout_measure.py <全书.pdf>
找两类格子，给出所在表格的文字（用于匹配源表签名）、列号、需要加宽多少（pt）：
  · 末行孤字：任一格里的任一条（按 <br>、圆点、编号分条），折行后末行只剩 1～2 个字（2026-10-05 起不限短格）（如「警 / 戒」），上一行接近撑满（自然折行）；
  · 短格折行：约 20 字以内的表头或短格自然折成 2 行，一行放得下。
  · 加宽省行（grow，SD-144）：页面右侧有空余时，按实测行宽算出加宽哪一列能让某一条少折一行，只用空余、不动邻列（2026-10-05 用户：「表格宽度根据内容灵活调整，目标是行数最少」）；
  · 实测收窄（nw）：某列每一行距右边都有大片空白（≥ 24pt），按实测最长行收窄不会增加任何行。
原文 <br> 主动换行（上一行明显短于格宽）、括注行、分条（•）格不管。只读。"""
import sys, os, re, json, collections
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

PAD = 13   # 单元格左右内边距合计（pt）：左 170 + 右 90 DXA ≈ 13pt（2026-10-03 悬挂圆点后）

# 源文件 <br> 主动换行的位置（与 check_layout 同口径，2026-10-05 下沉到这里，检查与修复看同一批格子）：
# 把 build/book.md 去标签后只留字母数字，<br> 记为「|」；某行文字后紧跟「|」说明这是作者主动换行，不算孤字
_BR = {}
def _brtext(srcmd):
    if srcmd in _BR: return _BR[srcmd]
    t = ''
    if srcmd and os.path.exists(srcmd):
        raw = re.sub(r'<br\s*/?>', '\x01', open(srcmd, encoding='utf-8').read())
        raw = re.sub(r'<[^>]+>', '', raw)
        t = ''.join(ch if (ch.isalnum() or ch == '\x01') else '' for ch in raw).replace('\x01', '|')
    _BR[srcmd] = t
    return t
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

def measure(pdf, srcmd=None):
    if srcmd is None: srcmd = os.path.join(os.path.dirname(os.path.abspath(pdf)), 'book.md')
    br = _brtext(srcmd)
    ebr = lambda line: bool(br) and (lambda k: bool(k) and (k + '|') in br)(''.join(ch for ch in line if ch.isalnum()))
    d = pymupdf.open(pdf); H = d[0].rect.height; found = []
    CW = d[0].rect.width - 72   # 版心宽（左右页边距各 36pt）
    for i, p in enumerate(d):
        try: tabs = p.find_tables().tables
        except Exception: continue
        lines = lines_of(p)
        for t in tabs:
            if t.bbox[3] - t.bbox[1] < 8: continue   # 2026-10-05：不再跳过页顶的表——SD-97 后顶端没有页眉，页顶的表都是跨页续段，跳过它们曾让续段测不到、否决票发不出（C-3 教训）
            ttext = norm(p.get_text(clip=t.bbox))
            # 实测收窄（nw）：列里每一行距两边的空白都 ≥ 24pt（按行宽算，居中列同样适用）→ 收到最长行 + 6pt，
            # 不会增加任何行。有跨列 / 跨行合并格的表整表跳过（收窄会挤到合并格里的长句）。
            # 每一段都必须表态（跨页的表一页一段）：空白不足或本段有合并格 / 折行列的，发 0 当否决票，
            # 否则只看到空白大的那一段就会把另一段的长行挤折（C-3 首列的教训，2026-10-05）
            _spanfree = all(c is not None for r0 in t.rows for c in r0.cells)
            if not _spanfree or t.col_count < 2:
                found.append({'page': i + 1, 'col': -1, 'extra_pt': 0, 'cell': '', 'first': '', 'prev': '',
                              'orphan': False, 'nw': True, 'cx': 0,
                              'tw': round(t.bbox[2] - t.bbox[0], 1), 'table': ttext})
            else:
                for k0 in range(t.col_count):
                    cols = [r0.cells[k0] for r0 in t.rows if k0 < len(r0.cells) and r0.cells[k0]]
                    cw0 = max(c0[2] - c0[0] for c0 in cols); mx = 0; n_l = 0; wrapped = False
                    for c0 in cols:
                        ls0 = cell_lines(lines, c0)
                        if len(ls0) > 1: wrapped = True
                        for l0 in ls0:
                            mx = max(mx, l0[1] - l0[0]); n_l += 1
                    slack = 0 if (not n_l or wrapped) else max(0, cw0 - PAD - mx - 6)   # 有折行的列不收（宽度由加宽通道管）
                    found.append({'page': i + 1, 'col': k0, 'extra_pt': round(slack, 1),
                                  'cell': ''.join(l0[2] for l0 in cell_lines(lines, cols[0]))[:24] if cols else '', 'first': '', 'prev': '',
                                  'orphan': False, 'nw': True,
                                  'cx': round((cols[0][0] + cols[0][2]) / 2 - t.bbox[0], 1) if cols else 0,
                                  'tw': round(t.bbox[2] - t.bbox[0], 1), 'table': ttext})
            for r in t.rows:
                for k, c in enumerate(r.cells):
                    if not c: continue
                    span = k + 1 < len(r.cells) and r.cells[k + 1] is None   # 跨列格：不能单独加宽一列，fit_fix 直接收紧字距（2026-10-05，原来整格跳过，检查器却照查）
                    ls = cell_lines(lines, c)
                    if len(ls) < 2: continue
                    cw = c[2] - c[0]
                    txt = ''.join(l[2] for l in ls).strip()
                    # 2026-10-05：逐条查（格内按 <br>、圆点、①② 分成几条；长格、带圆点的格都查），原来只查 2～3 行的短格
                    full = lambda l: (l[1] - l[0]) >= cw - PAD - 20
                    seg0 = 0
                    for q in range(1, len(ls)):
                        cur, prev = ls[q], ls[q - 1]
                        if not full(prev) or re.match(r'\s*[•–▪①-⑳]', cur[2]):
                            seg0 = q; continue                      # 新的一条（上一条主动换行 / 圆点 / 编号开头）
                        nxt = ls[q + 1] if q + 1 < len(ls) else None
                        if nxt is not None and full(cur) and not re.match(r'\s*[•–▪①-⑳]', nxt[2]): continue   # 这一条还没完
                        last = re.sub(r'[\s，。；：、（）()「」.,;:]', '', cur[2])
                        extra = 0
                        if 0 < len(last) <= 2 and not ebr(prev[2]):   # 末行孤字：把末行摊到这一条前面各行；上一行止于原文 <br> 的是主动换行，不算
                            extra = (cur[1] - cur[0]) / max(1, q - seg0) + 3
                        elif len(ls) == 2 and len(re.sub(r'\s', '', txt)) <= 20 and not ebr(ls[0][2]):   # 短格折两行：给够一行（原文 <br> 的不算）
                            extra = sum(l[1] - l[0] for l in ls) + PAD + 3 - cw
                        if extra > 0:
                            found.append({'page': i + 1, 'col': k, 'extra_pt': round(extra, 1), 'cell': txt[:24], 'first': norm(ls[0][2]), 'prev': norm(prev[2]), 'orphan': 0 < len(last) <= 2, 'cx': round((c[0] + c[2]) / 2 - t.bbox[0], 1), 'tw': round(t.bbox[2] - t.bbox[0], 1), 'table': ttext, 'span': span})
                    # SD-144（2026-10-05 用户）：「在页面右侧空间足够的情况下，尽量增加文字多的表格宽度，以减少行数」——
                    # 表格右侧有空余时，按实测行宽算出加宽多少能让某一条少折一行；空余放得下就交给 fit_fix 加宽这一列（只用空余，不从邻列匀）
                    free = CW - (t.bbox[2] - t.bbox[0])
                    if free > 20 and not span:
                        segs, cur_s = [], []
                        for l in ls:
                            if cur_s and (not full(cur_s[-1]) or re.match(r'\s*[•●–▪①-⑳]', l[2])): segs.append(cur_s); cur_s = []
                            cur_s.append(l)
                        if cur_s: segs.append(cur_s)
                        best, bsg = None, None
                        for sg in segs:
                            if len(sg) < 2: continue
                            need = sum(l[1] - l[0] for l in sg) * 1.03 / (len(sg) - 1) + PAD + 2 - cw
                            if 0 < need <= free - 2 and (best is None or need < best): best, bsg = need, sg
                        if best is not None:
                            found.append({'page': i + 1, 'col': k, 'extra_pt': round(best, 1), 'cell': txt[:24], 'first': norm(ls[0][2]), 'prev': norm(bsg[-2][2]), 'brs': [norm(l[2]) for l in bsg[:-1]], 'orphan': False, 'grow': True, 'cx': round((c[0] + c[2]) / 2 - t.bbox[0], 1), 'tw': round(t.bbox[2] - t.bbox[0], 1), 'table': ttext, 'span': False})
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
        bl = [b for b in p.get_text('blocks') if b[4].strip() and b[1] > 0.06 * H]
        fy = max((b[1] for b in bl if re.search(r'第\s*\d+\s*页\s*$', b[4])), default=H + 1)   # SD-97 页脚一行：章名 / 节名 / 页码；取最下面一个「第 N 页」（正文「出处 …第 51 页」也会匹配）
        bl = [b for b in bl if b[1] < fy - 2]
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
        ln = [l for l in lines_of(d[i]) if l[1] > 0.06 * H]
        fy = max((l[1] for l in ln if re.search(r'第\s*\d+\s*页\s*$', l[4])), default=H + 1)   # SD-97 页脚整行去掉
        ln = [l for l in ln if l[1] < fy - 2]
        if not (0 < len(ln) <= 2): continue
        try: tabs = [t for t in d[i - 1].find_tables().tables if t.bbox[1] > 0.07 * H and t.bbox[3] - t.bbox[1] > 10]
        except Exception: tabs = []
        # 上一页最后一个条目标题（不在页顶）：表收紧也救不回时，fit_fix 让它另起一页（P:）
        heads = [l for l in lines_of(d[i - 1]) if l[1] > 0.15 * H and re.match(r'^\s*([A-Z]-\d+|\d{1,3}\.)\s+\S', l[4])]
        head = re.sub(r'\s+', '', heads[-1][4]) if heads else ''
        if not tabs and not head: continue
        t = max(tabs, key=lambda t: t.bbox[3]) if tabs else None
        out.append({'page': i + 1, 'table': norm(d[i - 1].get_text(clip=t.bbox)) if t else '', 'head': head})
    return out
