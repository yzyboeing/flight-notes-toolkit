#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""t13_post.py —— 定稿后收窄（2026-10-22，1021R4 列宽空白每轮二十余处收不干净）。
fit_fix 最后一遍排好以后，按成品 PDF 实测（与 check_layout T13 同口径：本格行数不增的严格余量、跨页表取各段最小值、只查没折行的列）
把右侧仍空 ≥ 10pt 的列直接在成品 docx 里收窄到「余量约 4pt」：改 gridCol、该列各格 tcW 和表宽。
只收不折行的列——列内行数不变、表高不变，分页不动。两型等宽列（col-eq）同步按较小余量收，保持等宽。
由 fit_fix.py 收尾时调用：run(docx, pdf, 源 md, tbls) → 改了几列；调用方随后重转 PDF。"""
import os, re, sys, zipfile, shutil, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import layout_measure as LM

N = lambda s: re.sub(r'[^\w]', '', s.replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&'))
THR, KEEP = 10.0, 6.0   # 余量 ≥ 10pt 才收（T13 口径），收到约剩 6pt


def _fixsigs(srcmd):
    """与 check_layout 同口径的门槛：w-NN 表 24pt、锁宽 30pt；等宽 / 点名满宽 / 图表并排的表不收（10**6）"""
    nosp = lambda t: re.sub(r'\s+', '', t)
    try: bm = open(srcmd, encoding='utf-8').read()
    except OSError: return []
    out = []
    for m in re.finditer(r'<table[^>]*>((?:(?!</table>).)*?)<tr class="hdr">((?:(?!</tr>).)*\bw-\d+(?:(?!</tr>).)*)</tr>((?:(?!</table>).)*?</tr>)', bm, re.S):
        out.append((nosp(re.sub(r'<[^>]+>', '', m.group(2) + m.group(3)))[:30], 10 ** 6 if ('keep-wide' in m.group(2) or 'col-eq' in m.group(2)) else 30 if re.search(r'\bw-lock\b', m.group(2)) else 24))
    for m in re.finditer(r'<table[^>]*>((?:(?!</table>).)*?)<tr class="hdr">((?:(?!</tr>).)*\b(?:col-eq|keep-wide)\b(?:(?!</tr>).)*)</tr>((?:(?!</table>).)*?</tr>)', bm, re.S):
        out.append((nosp(re.sub(r'<[^>]+>', '', m.group(2) + m.group(3)))[:30], 10 ** 6))
    for m in re.finditer(r'^%%FIG(?:SIDE|RIGHT)%%[^\n]*\n\s*<table[^>]*>((?:(?!</table>).)*?)<tr class="hdr">((?:(?!</tr>).)*)</tr>((?:(?!</table>).)*?</tr>)', bm, re.S | re.M):
        out.append((nosp(re.sub(r'<[^>]+>', '', m.group(1) + m.group(2) + m.group(3)))[:30], 10 ** 6))
        out.append((nosp(re.sub(r'<[^>]+>', '', m.group(2) + m.group(3)))[:30], 10 ** 6))
    return out


def plan(pdf, srcmd, kf=None):
    """[(表文字, 列, 可收 pt)]——同 check_layout T13 的分组口径与门槛"""
    res = LM.measure(pdf, srcmd=srcmd)
    FIX = _fixsigs(srcmd)
    NWL = set()
    try:
        for l in open(kf, encoding='utf-8'):
            if l.startswith('NWL:'): k0, sg = l.strip()[4:].split(':', 1); NWL.add((sg[:20], int(k0)))
    except (OSError, TypeError): pass
    segs = collections.OrderedDict()
    for x in res:
        if x.get('nw') and x.get('col', -1) >= 0 and 'tb' in x:
            segs.setdefault((x['page'], x['tb']), []).append(x)
    keys = list(segs); grp, gid = {}, 0
    for q, k in enumerate(keys):
        e = segs[k][0]
        cand = [kk for kk in keys[:q] if kk[0] == k[0] - 1 and segs[kk][0]['nc'] == e['nc'] and abs(segs[kk][0]['tw'] - e['tw']) < 3]
        if cand: grp[k] = grp[cand[-1]]
        else: gid += 1; grp[k] = gid
    byg = collections.defaultdict(list)
    for k in keys: byg[grp[k]].append(k)
    out = []
    for g, ks in byg.items():
        for c in range(segs[ks[0]][0]['nc']):
            vals = [next((e for e in segs[k] if e['col'] == c), None) for k in ks]
            if any(v is None for v in vals): continue
            sl = min(v.get('extra_strict', v['extra_pt']) for v in vals)
            tt = re.sub(r'\s+', '', vals[0].get('table', ''))
            thr = max([th for sg, th in FIX if sg and re.sub(r'[^\w]', '', tt).startswith(re.sub(r'[^\w]', '', sg)[:24])] or [THR])
            if sl < 30 and any(tt.startswith(sg) and kk == c for sg, kk in NWL): continue
            if sl >= thr and thr < 10 ** 6 and not any(v.get('wrapped') or v.get('anywrap') for v in vals):
                out.append((vals[0].get('table', ''), c, sl, vals[0]['nc']))
    return out


def _cells(tr):
    """一行里各格：(起始列, 跨列数, 格 XML 位置)"""
    out, col = [], 0
    for m in re.finditer(r'<w:tc>(?:(?!<w:tc>).)*?</w:tc>|<w:tc\b[^>]*>(?:(?!<w:tc\b).)*?</w:tc>', tr, re.S):
        sp = re.search(r'<w:gridSpan w:val="(\d+)"', m.group(0))
        n = int(sp.group(1)) if sp else 1
        out.append((col, n, m.start(), m.end())); col += n
    return out


def run(docx, pdf, srcmd, tbls=None, kf=None):
    todo = plan(pdf, srcmd, kf)
    if not todo: return 0
    eqs = {}   # 表文字前缀 → 等宽列
    for t in tbls or []:
        if t.get('eq'): eqs[N(t.get('text', ''))[:30]] = set(t['eq'])
    z = zipfile.ZipFile(docx); xml = z.read('word/document.xml').decode('utf-8')
    tabs = [(m.start(), m.end(), m.group(0)) for m in re.finditer(r'<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>', xml, re.S)]
    ttext = [N(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tb))) for _, _, tb in tabs]
    cuts = collections.defaultdict(dict)   # 表序号 → {列: 收窄 DXA}
    for txt, c, sl, nc in todo:
        key = N(txt)[:30]
        if len(key) < 6: continue
        hit = [i for i, t in enumerate(ttext) if t.startswith(key)]
        if len(hit) != 1: continue
        i = hit[0]
        grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tabs[i][2])]
        if len(grid) != nc: continue
        eq = eqs.get(key, set())
        d = int((sl - KEEP) * 20)
        if c in eq:   # 等宽列：各列按最小余量同步收
            others = [s2 for t2, c2, s2, _ in todo if N(t2)[:30] == key and c2 in eq]
            if len(others) < len(eq): continue
            d = int((min(others) - KEEP) * 20)
            for q in eq: cuts[i][q] = max(cuts[i].get(q, 0), d)
        else:
            cuts[i][c] = max(cuts[i].get(c, 0), d)
    if not cuts: return 0
    n = 0
    for i in sorted(cuts, reverse=True):   # 从后往前改，前面的位置不乱
        s0, e0, tb = tabs[i]
        grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tb)]
        cut = {c: min(d, max(0, grid[c] - 600)) for c, d in cuts[i].items()}   # 列宽不少于 30pt
        cut = {c: d for c, d in cut.items() if d > 0}
        if not cut: continue
        k = [0]
        def gsub(m):
            q = k[0]; k[0] += 1
            return '<w:gridCol w:w="%d"' % (int(m.group(1)) - cut.get(q, 0))
        tb2 = re.sub(r'<w:gridCol w:w="(\d+)"', gsub, tb)
        tot = sum(cut.values())
        tb2 = re.sub(r'(<w:tblW w:type="dxa" w:w=")(\d+)(")', lambda m: m.group(1) + str(max(600, int(m.group(2)) - tot)) + m.group(3), tb2, count=1)
        def rsub(mr):
            tr = mr.group(0); out, last = [], 0
            for col, span, a, b in _cells(tr):
                dd = sum(cut.get(q, 0) for q in range(col, col + span))
                cx = tr[a:b]
                if dd: cx = re.sub(r'(<w:tcW w:type="dxa" w:w=")(\d+)(")', lambda m: m.group(1) + str(max(300, int(m.group(2)) - dd)) + m.group(3), cx, count=1)
                out.append(tr[last:a]); out.append(cx); last = b
            out.append(tr[last:]); return ''.join(out)
        tb2 = re.sub(r'<w:tr\b[^>]*>.*?</w:tr>', rsub, tb2, flags=re.S)
        xml = xml[:s0] + tb2 + xml[e0:]; n += len(cut)
    tmp = docx + '.t13tmp'
    with zipfile.ZipFile(docx) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            zo.writestr(it, xml.encode('utf-8') if it.filename == 'word/document.xml' else zi.read(it.filename))
    os.replace(tmp, docx)
    return n


if __name__ == '__main__':
    if len(sys.argv) < 4: print(__doc__); sys.exit(1)
    print('定稿后收窄 %d 列' % run(sys.argv[1], sys.argv[2], sys.argv[3], kf=(sys.argv[4] if len(sys.argv) > 4 else None)))


# ---------------------------------------------------------------- 2026-10-22 定稿后加宽 / 借宽（T14 有空位却折行、T8 末行孤字）
TOTAL = 16838 - 1440   # 版心宽 DXA


def _tops(pdf):
    import pymupdf
    return [re.sub(r'\s+', '', p.get_text())[:60] for p in pymupdf.open(pdf)]


def grow_plan(pdf, srcmd):
    """[(表文字, {列: 增减 DXA})]——grow 项：表右侧空位加宽折行列；orphan 项：同表可让出列借宽（每列只让 lend 以内）"""
    res = LM.measure(pdf, srcmd=srcmd)
    FIX = {sg for sg, th in _fixsigs(srcmd) if th >= 10 ** 6}   # 等宽 / 点名满宽 / 图表并排表不动
    nosp = lambda t: re.sub(r'[^\w]', '', t or '')
    out = []
    lend = collections.defaultdict(lambda: collections.defaultdict(lambda: 10 ** 9))
    for x in res:
        if x.get('nw') and x.get('col', -1) >= 0: lend[nosp(x.get('table'))[:30]][x['col']] = min(lend[nosp(x.get('table'))[:30]][x['col']], x.get('lend') or 0)
    seen = set()
    for x in res:
        key = nosp(x.get('table'))[:30]
        if not key or any(key.startswith(nosp(sg)[:24]) for sg in FIX if sg): continue
        if x.get('grow') and not x.get('fill') and not x.get('span') and x.get('col', -1) >= 0 and (key, 'g', x['col']) not in seen:
            seen.add((key, 'g', x['col'])); out.append((key, {x['col']: int((x['extra_pt'] + 2) * 20)}, 'grow'))
        elif x.get('orphan') and not x.get('span') and x.get('col', -1) >= 0 and (key, 'o', x['col']) not in seen:
            need = x['extra_pt'] + 1; don = sorted(((v, c) for c, v in lend[key].items() if c != x['col'] and 2 <= v < 10 ** 9), reverse=True)
            if sum(v for v, c in don) < need: continue
            seen.add((key, 'o', x['col'])); d = {x['col']: int(need * 20)}; left = need
            for v, c in don:
                t = min(v - 1, left); d[c] = d.get(c, 0) - int(t * 20); left -= t
                if left <= 0: break
            out.append((key, d, 'borrow'))
    return out


def _apply(xml, tabs, ttext, key, delta):
    hit = [i for i, t in enumerate(ttext) if t.startswith(key)]
    if len(hit) != 1: return xml, False
    s0, e0, tb = tabs[hit[0]]
    grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tb)]
    if any(c >= len(grid) for c in delta): return xml, False
    tot = sum(delta.values())
    tw = re.search(r'<w:tblW w:type="dxa" w:w="(\d+)"', tb)
    if not tw or int(tw.group(1)) + tot > TOTAL - 20 or any(grid[c] + d < 600 for c, d in delta.items()): return xml, False
    k = [0]
    def gsub(m):
        q = k[0]; k[0] += 1
        return '<w:gridCol w:w="%d"' % (int(m.group(1)) + delta.get(q, 0))
    tb2 = re.sub(r'<w:gridCol w:w="(\d+)"', gsub, tb)
    tb2 = re.sub(r'(<w:tblW w:type="dxa" w:w=")(\d+)(")', lambda m: m.group(1) + str(int(m.group(2)) + tot) + m.group(3), tb2, count=1)
    def rsub(mr):
        tr = mr.group(0); out, last = [], 0
        for col, span, a, b in _cells(tr):
            dd = sum(delta.get(q, 0) for q in range(col, col + span))
            cx = tr[a:b]
            if dd: cx = re.sub(r'(<w:tcW w:type="dxa" w:w=")(\d+)(")', lambda m: m.group(1) + str(max(300, int(m.group(2)) + dd)) + m.group(3), cx, count=1)
            out.append(tr[last:a]); out.append(cx); last = b
        out.append(tr[last:]); return ''.join(out)
    tb2 = re.sub(r'<w:tr\b[^>]*>.*?</w:tr>', rsub, tb2, flags=re.S)
    return xml[:s0] + tb2 + xml[e0:], True


def _write(docx, xml):
    tmp = docx + '.postmp'
    with zipfile.ZipFile(docx) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            zo.writestr(it, xml.encode('utf-8') if it.filename == 'word/document.xml' else zi.read(it.filename))
    os.replace(tmp, docx)


def run_grow(docx, pdf, srcmd, to_pdf):
    """加宽 / 借宽后重转 PDF，逐页比对页首文字：分页动了就逐张试、只留不动分页的改动。返回 (保留数, 新 PDF 路径或 None)"""
    plan_ = grow_plan(pdf, srcmd)
    if not plan_: return 0, None
    base = zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')
    tops0 = _tops(pdf)
    def build(items):
        xml = base
        for key, d, kind in items:
            tabs = [(m.start(), m.end(), m.group(0)) for m in re.finditer(r'<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>', xml, re.S)]
            ttext = [N(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tb))) for _, _, tb in tabs]
            xml, ok = _apply(xml, tabs, ttext, key, d)
        _write(docx, xml); return to_pdf()
    p = build(plan_)
    if _tops(p) == tops0: return len(plan_), p
    keep = []
    for it in plan_:   # 分页动了：逐张试
        p = build(keep + [it])
        if _tops(p) == tops0: keep.append(it)
    p = build(keep)
    return len(keep), p
