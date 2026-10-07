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

# 2026-10-07（用户，速查第 15 页「座舱高度约机场标高 + 1000ft，压差 ≈ 4psi」）：断点前一行没排满、但下一行开头那个词接不回来，
# 也是被迫折行（排版器在「+」「（」等前面的断点换行，前一行右侧留一大截）。以前只认「排到右缘」，这种折行漏判、不会加宽。
_TOK = re.compile(r'\s*([（(【「]?[A-Za-z0-9.,/\-:+%°×±≈≤≥~～·\'’]+[）)】」]?|[（(【「]?.[）)】」，。；、：]?)')
def _wt(t): return sum(0.55 if ord(ch) < 0x2E80 and ch not in '±×≈≤≥→' else 1.0 for ch in t)
def forced_wrap(cw, lw, ltext, ntext, pad):
    """cw 格宽；lw / ltext 本行宽度与文字；ntext 下一行文字。下一行第一个词接到本行后超出可排宽度 → 被迫折行"""
    m = _TOK.match(ntext or '')
    if not m or not ltext.strip(): return False
    u = lw / max(_wt(ltext.strip()), 0.5)                 # 本行每单位字宽（按中文 1、西文 0.55 折算）
    # 两行合起来超出可排宽度 = 这一处换行是排不下造成的（排版器可能在更早的断点换行，只看下一行第一个词会漏判）
    return lw + _wt((ntext or '').strip()) * u > (cw - pad) - 1

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
            # 表后紧贴的「注：/公司差异/警告」会被识别成表格末几行：取表文字时去掉，否则与源表签名对不上，fit_fix 的加宽 / 收窄 / 去孤字都落不了账（2026-10-06）
            try:
                _rows = t.extract()
                _keep = [r0 for r0 in _rows if not re.match(r'\s*(注[：:]|公司差异|警告[：:])', (r0[0] or ''))]
                ttext = norm(''.join(c0 or '' for r0 in _keep for c0 in r0)) or norm(p.get_text(clip=t.bbox))
            except Exception:
                ttext = norm(p.get_text(clip=t.bbox))
            # 量列边界、表宽也不算表后注行：注行比表宽、铺满版心，会让 PyMuPDF 把表右侧空白当成一列（余量 0，否决真列），表宽也被算成满版心（拉满从不触发）（2026-10-06）
            try:
                _ex = t.extract()
                _note = {k for k, r0 in enumerate(_ex) if re.match(r'\s*(注[：:]|公司差异|警告[：:])', (r0[0] or ''))}
            except Exception:
                _note = set()
            _cells = [c0 for k, r0 in enumerate(t.rows) if k not in _note for c0 in r0.cells if c0]
            _txt = [c0 for c0 in _cells if cell_lines(lines, c0)]
            TX2 = max([c0[2] for c0 in _txt] + [t.bbox[0] + 10])   # 表宽取有文字格子的最右边界
            _cells = [c0 for c0 in _cells if c0[0] < TX2 - 1]       # 表右侧 PyMuPDF 虚构的空格子（注行比表宽时出现）丢掉
            # 实测收窄（nw）：列里每一行距两边的空白都 ≥ 24pt（按行宽算，居中列同样适用）→ 收到最长行 + 6pt，
            # 不会增加任何行。有跨列 / 跨行合并格的表整表跳过（收窄会挤到合并格里的长句）。
            # 每一段都必须表态（跨页的表一页一段）：空白不足或本段有合并格 / 折行列的，发 0 当否决票，
            # 否则只看到空白大的那一段就会把另一段的长行挤折（C-3 首列的教训，2026-10-05）
            # 2026-10-06（用户，第 10 页再循环风扇、第 62 页 EEC 备用方式）：有合并格的表也收窄——按格子横向位置还原列边界，
            # 单列格定每列的空白；跨列格的余量按所跨列数均分，作为这些列的上限（跨列格本身折行则这些列都不收）
            NEWIT = r'\s*([●•▪·–\-\u2460-\u2473【]|[A-H]-\d|\d{1,2}[.、)）]|注[：:])'
            def _measure_cell(c0):
                ls0 = cell_lines(lines, c0); wr = False; mx0 = 0
                for q0 in range(len(ls0) - 1):
                    if ((c0[2] - c0[0]) - (ls0[q0][1] - ls0[q0][0]) < PAD + 14 or forced_wrap(c0[2] - c0[0], ls0[q0][1] - ls0[q0][0], ls0[q0][2], ls0[q0 + 1][2], PAD)) and not re.match(NEWIT, ls0[q0 + 1][2]) and not ebr(ls0[q0][2]): wr = True   # 按整行宽度判：居中列的短行左右各空一点，只看右侧会误判成排满（2026-10-06）   # 真折行：排到右缘、下一行不是新的一条
                for l0 in ls0: mx0 = max(mx0, l0[1] - l0[0])
                return ls0, wr, mx0
            xs = []
            for c0 in sorted({round(c[0], 1) for c in _cells}):
                if not xs or c0 - xs[-1] > 2: xs.append(c0)
            xs.append(TX2)
            ncol_t = len(xs) - 1
            if ncol_t < 2:
                found.append({'page': i + 1, 'col': -1, 'extra_pt': 0, 'cell': '', 'first': '', 'prev': '',
                              'orphan': False, 'nw': True, 'cx': 0,
                              'tw': round(TX2 - t.bbox[0], 1), 'table': ttext})
            else:
                cover = lambda c0: [q for q in range(ncol_t) if xs[q] >= c0[0] - 2 and xs[q + 1] <= c0[2] + 2]
                # SD-151 按行高收窄（2026-10-06 用户，第 31 / 41 / 103 / 239 页：「在不增加行数的情况下尽量让这一列窄一点」）：
                # 单元格可容纳的行数按它自己的格高算（含跨行格），只要重排后不超过这个行数，这一列就能收窄——同一行别的格更高时，这格多折几行也不增加行高
                gaps = []   # 行距只在同一格内取（不同列的文字上下错开，混在一起会把行距算小）
                for c0 in _cells:
                    if not c0: continue
                    yy = sorted((l[1] + l[3]) / 2 for l in lines if c0[0] - 1 <= (l[0] + l[2]) / 2 <= c0[2] + 1 and c0[1] - 1 <= (l[1] + l[3]) / 2 <= c0[3] + 1)
                    gaps += [b2 - a2 for a2, b2 in zip(yy, yy[1:]) if 9 <= b2 - a2 <= 30]
                pitch = sorted(gaps)[len(gaps) // 2] if gaps else 17.0
                def _minw(c0, ls0, rowmode=True):
                    paras, cur = [], []
                    for l0 in ls0:
                        if cur and (re.match(NEWIT, l0[2]) or ebr(cur[-1][2]) or (((c0[2] - c0[0]) - (cur[-1][1] - cur[-1][0])) >= PAD + 14 and not forced_wrap(c0[2] - c0[0], cur[-1][1] - cur[-1][0], cur[-1][2], l0[2], PAD))): paras.append(cur); cur = []   # 源文件 <br> 处另起一段
                        cur.append(l0)
                    if cur: paras.append(cur)
                    cap = min(max(len(ls0), int((c0[3] - c0[1] - 4) / pitch + 0.35)), len(ls0) + 1) if (rowcap_ok and rowmode) else len(ls0)   # 表不挤时仍按「本格行数不增」；借行高时每格最多多折一行（「适当收窄」，不把短格挤成一列碎字）
                    tok = max([len(m0) * 5.2 for l0 in ls0 for m0 in re.findall(r'[A-Za-z0-9][A-Za-z0-9./\-:+%°]*', l0[2])] + [0])
                    widest = max((l0[1] - l0[0]) for l0 in ls0)
                    lo, hi = max(tok, 24) + PAD + 2, widest + PAD + 6
                    def need(w):
                        n = 0
                        for pg in paras:
                            L = sum(l0[1] - l0[0] for l0 in pg) * (1.10 if len(pg) > 1 else 1.0); av = w - PAD - 2
                            k = -(-L // av); n += k
                            if k > 1 and L - (k - 1) * av < 55: n += 9   # 末行不足约 6 个字：视为不行（不制造末行孤字）
                        return n
                    if need(hi) > cap: return hi
                    while hi - lo > 2:
                        mid = (lo + hi) / 2
                        if need(mid) <= cap: hi = mid
                        else: lo = mid
                    return hi
                # 按行高收窄的前提：表已占满版心（右侧空余 < 20pt）且有格真折行——确有列需要宽度，挪出来的宽度才有去处
                rowcap_ok = (CW - (TX2 - t.bbox[0]) < 20) and any(_measure_cell(c0)[1] for c0 in _cells if len(cover(c0)) == 1)
                mxk, nlk, wrk, capk, mxs, anyw = [0] * ncol_t, [0] * ncol_t, [False] * ncol_t, [10 ** 9] * ncol_t, [0] * ncol_t, [False] * ncol_t
                for c0 in _cells:
                    if not c0: continue
                    cv = cover(c0)
                    if not cv: continue
                    ls0, wr, mx0 = _measure_cell(c0)
                    if len(cv) == 1:
                        k0 = cv[0]; nlk[k0] += len(ls0); anyw[k0] = anyw[k0] or wr
                        if ls0:
                            mw = _minw(c0, ls0); mxk[k0] = max(mxk[k0], mw - PAD - 6)   # 以「行高不增」的最小宽度代替最长行
                            mxs[k0] = max(mxs[k0], _minw(c0, ls0, False) - PAD - 6)   # 严格口径：本格行数不增（已收窄过的列只用这个，防一遍遍往下收）
                            if wr and len(ls0) >= int((c0[3] - c0[1] - 4) / pitch + 0.35): wrk[k0] = True   # 真折行且这格就是撑起行高的那格：收窄过头的信号
                    elif len(cv) < ncol_t:   # 整行通栏的格（注、提示、前提行）不约束单列宽度：挪出的宽度由拉满补回，总宽不变
                        sl = 0 if (wr or not ls0) else max(0, (c0[2] - c0[0]) - PAD - mx0 - 6)
                        for q in cv: capk[q] = min(capk[q], sl / len(cv))
                rows_sl = []; f0 = len(found)
                for k0 in range(ncol_t):
                    cw0 = xs[k0 + 1] - xs[k0]
                    slack = 0 if (not nlk[k0] or wrk[k0]) else max(0, cw0 - PAD - mxk[k0] - 6)   # 撑起行高的格已折行到满：不收
                    slack = min(slack, capk[k0])
                    slack_s = 0 if (not nlk[k0] or wrk[k0] or anyw[k0]) else min(capk[k0], max(0, cw0 - PAD - mxs[k0] - 6))   # 有真折行的列不按严格口径收（只在「让给更挤的列」那一次收），防收窄 / 拉满来回拉锯
                    rows_sl.append((slack, slack_s))
                    found.append({'page': i + 1, 'col': k0, 'extra_pt': round(slack, 1), 'extra_strict': round(slack_s, 1), 'wrapped': wrk[k0],
                                  'anywrap': anyw[k0], 'nl': nlk[k0], 'tb': round(t.bbox[1], 1), 'bb': round(t.bbox[3], 1), 'ph': round(p.rect.height, 1), 'nc': ncol_t, 'head': (t.extract()[0][k0] or '')[:10] if t.extract() and k0 < len(t.extract()[0]) else '',
                                  'cell': '', 'first': '', 'prev': '',
                                  'orphan': False, 'nw': True,
                                  'cx': round((xs[k0] + xs[k0 + 1]) / 2 - t.bbox[0], 1),
                                  'tw': round(TX2 - t.bbox[0], 1), 'table': ttext})
                # 每张表一次只让一列借行高收窄（多借的那列）；其余列按「本格行数不增」——几列同时借同一行的行高会把行撑高（2026-10-06 页数来回跳的教训）
                gain = [a2 - b2 for a2, b2 in rows_sl]
                kb = max(range(len(gain)), key=lambda q: gain[q]) if gain and max(gain) > 0 else -1
                for q, e in enumerate(found[f0:f0 + len(rows_sl)]):
                    if q != kb: e['extra_pt'] = e['extra_strict']
            fillc = {}   # SD-148 实测拉满：给哪列都省不出行时，本表真折行最多的列（每表每遍一列）
            grew = False
            for _ri, r in enumerate(t.rows):
                if _ri in _note: continue   # 表后注行不当表格格子查（注段落的孤字由 SD-130 另管）
                for k, c in enumerate(r.cells):
                    if not c: continue
                    if c[0] >= TX2 - 1: continue   # 表右侧 PyMuPDF 虚构的空格子
                    span = k + 1 < len(r.cells) and r.cells[k + 1] is None and c[2] < TX2 - 2   # 后面的 None 落在表宽之外（虚构空列）不算跨列（2026-10-06 第 126 页 VNAV 目标速度：误判跨列，加宽 / 拉满全被跳过）   # 跨列格：不能单独加宽一列，fit_fix 直接收紧字距（2026-10-05，原来整格跳过，检查器却照查）
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
                            found.append({'page': i + 1, 'col': k, 'extra_pt': round(extra, 1), 'cell': txt[:24], 'first': norm(ls[0][2]), 'prev': norm(prev[2]), 'orphan': 0 < len(last) <= 2, 'cx': round((c[0] + c[2]) / 2 - t.bbox[0], 1), 'tw': round(TX2 - t.bbox[0], 1), 'table': ttext, 'span': span})
                    # SD-144（2026-10-05 用户）：「在页面右侧空间足够的情况下，尽量增加文字多的表格宽度，以减少行数」——
                    # 表格右侧有空余时，按实测行宽算出加宽多少能让某一条少折一行；空余放得下就交给 fit_fix 加宽这一列（只用空余，不从邻列匀）
                    free = CW - (TX2 - t.bbox[0])
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
                            found.append({'page': i + 1, 'col': k, 'extra_pt': round(best, 1), 'cell': txt[:24], 'first': norm(ls[0][2]), 'prev': norm(bsg[-2][2]), 'brs': [norm(l[2]) for l in bsg[:-1]], 'orphan': False, 'grow': True, 'cx': round((c[0] + c[2]) / 2 - t.bbox[0], 1), 'tw': round(TX2 - t.bbox[0], 1), 'table': ttext, 'span': False})
                            grew = True
                        else:
                            _, wr_c, _ = _measure_cell(c)   # 与收窄同一口径判真折行（整行排满、下一行不是新条目、不止于原文 <br>）
                            nwr = sum(len(sg) - 1 for sg in segs if len(sg) >= 2) if wr_c else 0
                            if nwr: fillc[k] = (fillc.get(k, (0, 0))[0] + nwr, round((c[0] + c[2]) / 2 - t.bbox[0], 1))
            # SD-148 / 151（2026-10-06 用户：「哪怕减少不了行数，也可以把空白区域利用起来」）：这一遍没有能省行的加宽项、表格右侧仍有空余时，
            # 空余整块给本表真折行最多的列（最挤的列），每表每遍只给一列；fit_fix 记 W:（fill 不受加宽上限与次数限制）
            free_t = CW - (TX2 - t.bbox[0])
            if fillc and free_t > 20:   # 有省行加宽项时也报：省行项可能是原文 <br> 被 fit_fix 跳过，由 fit_fix 判断本表这遍是否已加宽
                kf = max(fillc, key=lambda q: fillc[q][0])
                found.append({'page': i + 1, 'col': kf, 'extra_pt': round(free_t - 4, 1), 'cell': '', 'first': '', 'prev': '', 'brs': [],
                              'orphan': False, 'grow': True, 'fill': True, 'cx': fillc[kf][1], 'tw': round(TX2 - t.bbox[0], 1), 'table': ttext, 'span': False})
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
