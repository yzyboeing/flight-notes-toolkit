#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fixw_fit.py —— 定宽表 / 图旁表按「行数最少」重分列宽（SD-144 / SD-203，2026-10-17 用户：
「适当收窄一下图片，把右侧表格的独字解决一下，适当调整就可以减少两行或以上行数的表格都处理一下」）

fit_fix 只调自动宽度表；定宽（w-NN）表和图旁表（%%FIGSIDE%% / %%FIGRIGHT%%）的列宽写死在源文件里，以前没人按行数优化。
本脚本按源文字估算每格行数（9.5pt：汉字 190 DXA、西文约 100、大写约 135、格边距 260、「–」子项缩进 330），
在列宽组合里找「行数 + 独字」最少的一组，满足下列之一才改（check_struct T18 报的就是这些表，检查与修复同口径）：
  · 省 ≥ 2 行（定宽表的独字交给 fit_fix 实测借宽；图旁表另看独字）；收窄不自动做（合并格行高估算偏粗）
定宽表合计取原合计或 100%；图旁表另把图最多收窄 20%（≥ 55mm）；短标签列（每项 ≤ 约 10 字）保持一行。
用法：python3 fixw_fit.py [--repo 笔记库] [--apply] [--min 2]       只报 / 写回源文件"""
import io, os, re, glob, sys, math
REPO = os.path.abspath(os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else '~/flight-repos/gh-private'))
APPLY = '--apply' in sys.argv
MINSAVE = int(sys.argv[sys.argv.index('--min') + 1]) if '--min' in sys.argv else 2
FIG_MINSAVE = 1   # P06（SD-209，2026-10-19 用户：「适当减小插图大小以适当的减少孤字问题和表格行数」）：图旁表省 1 行或消孤字就收图
TOTAL = 16838 - 1440
CJ, LA, UP, PAD, SUB = 190, 100, 135, 260, 330

def plain(s): return re.sub(r'<[^>]+>', '', s).replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&nbsp;', ' ').strip()
def tw(p): return sum(CJ if ord(c) > 0x2000 else (UP if c.isupper() else LA) for c in p)
def cell_lines(cell, colw):
    tot = 0; orph = 0
    for it in re.split(r'<br\s*/?>', cell):
        p = plain(it)
        if not p: continue
        sub = bool(re.match(r'^[-–]\s', p))
        avail = colw - PAD - (SUB if sub else 0)
        if avail <= CJ * 2: return 99, 0
        w = tw(p); n = max(1, math.ceil(w / avail))
        tot += n
        if n > 1 and (w - (n - 1) * avail) / CJ <= 2.2: orph += 1
    return max(tot, 1), orph
def parse(tb):
    hm = re.search(r'<tr class="hdr">(.*?)</tr>', tb, re.S)
    ths = re.findall(r'<th([^>]*)>(.*?)</th>', hm.group(1), re.S)
    rows = [re.findall(r'<td([^>]*)>(.*?)</td>', r.group(1), re.S)
            for r in re.finditer(r'<tr(?![^>]*class="(?:hdr|premise|note)")[^>]*>(.*?)</tr>', tb[hm.end():], re.S)]
    return ths, [r for r in rows if r]
def layout(ths, rows, widths):
    n = len(widths); occ = {}; heights = []; orph = 0; spans = []
    for ri, cells in enumerate(rows):
        k = 0; h = 1
        for a, c in cells:
            while occ.get((ri, k)): k += 1
            cs = int((re.search(r'colspan="(\d+)"', a) or [0, 1])[1]); rs = int((re.search(r'rowspan="(\d+)"', a) or [0, 1])[1])
            w = sum(widths[k:k + cs]) if k < n else widths[-1]
            L, o = cell_lines(c, w); orph += o
            for dr in range(rs):
                for dc in range(cs): occ[(ri + dr, k + dc)] = True
            if rs == 1: h = max(h, L)
            else: spans.append((ri, rs, L))
            k += cs
        heights.append(h)
    for ri, rs, L in spans:
        have = sum(heights[ri:ri + rs])
        if L > have: heights[ri + rs - 1] += L - have
    return sum(heights), orph
def lows(ths, rows, total_dxa):
    n = len(ths)
    lo = [max(4, math.ceil(100 * (tw(plain(h[1])[:6]) + PAD) / total_dxa)) for h in ths]
    for k in range(n):
        its = [plain(x) for cells in rows for kk, (a, c) in enumerate(cells) if kk == k for x in re.split(r'<br\s*/?>', c) if plain(x)]
        if its and max(tw(x) for x in its) <= 10 * CJ:
            lo[k] = max(lo[k], math.ceil(100 * (max(tw(x) for x in its) + PAD) / total_dxa))
    return lo
def search(ths, rows, cur_pct, total_dxa, totals):
    """各总宽下找最佳分配；返回 [(总宽, 行数, 独字, 组合)]"""
    n = len(cur_pct); step = 1 if n <= 2 else (2 if n == 3 else 3)
    lo = lows(ths, rows, total_dxa); out = []
    # 不把一列收到比它现在还窄、且窄到最长一项放不下一行（实测教训：重要天气预告图首列被压到 8%，2026-10-17）
    for k in range(n):
        its = [plain(x) for cells in rows for kk, (a, c) in enumerate(cells) if kk == k for x in re.split(r'<br\s*/?>', c) if plain(x)]
        if its:
            need = math.ceil(100 * (max(tw(x) for x in its) + PAD) / total_dxa)
            lo[k] = max(lo[k], min(cur_pct[k], need))
    def gen(k, left):
        if k == n - 1:
            if lo[k] <= left: yield (left,)
            return
        for v in range(lo[k], left - sum(lo[k + 1:]) + 1, step):
            for rest in gen(k + 1, left - v): yield (v,) + rest
    for tot in totals:
        if tot < sum(lo): continue
        bestv = None
        for comb in gen(0, tot):
            L, o = layout(ths, rows, [p * total_dxa / 100 for p in comb])
            key = (L + 0.6 * o, sum(abs(a - b) for a, b in zip(comb, cur_pct)))
            if bestv is None or key < bestv[0]: bestv = (key, L, o, comb)
        if bestv: out.append((tot, bestv[1], bestv[2], bestv[3]))
    return out
def propose(tb, before, pct0=None):
    """返回 None 或 dict(kind, pct, fig, info)；pct0：没有 w-NN 的图旁表由调用方按成品实测列宽给出起始比例（P06）"""
    if '<tr class="hdr">' not in tb: return None
    ths, rows = parse(tb)
    pct = [int((re.search(r'\bw-(\d+)', a) or [0, 0])[1]) for a, _ in ths]
    fig = re.match(r'%%(FIGSIDE|FIGRIGHT)%%\s*([^|]+)\|([^|]*)\|\s*(\d+)', before)
    if fig and not all(pct) and pct0 and len(pct0) == len(ths): pct = [max(1, int(round(x))) for x in pct0]
    if not all(pct) or not rows or len(ths) > 5: return None
    if fig:
        fmm = int(fig.group(4)); cands = []
        fixed = '定' in [x.strip() for x in before.split('|')[3:]]   # P06：图行第 4 段写「定」= 图宽已人工看图定下，不再提议收图（只提列宽）
        for f2 in ([fmm] if fixed or fmm <= 70 else [fmm] + list(range(fmm - 5, max(55, int(fmm * 0.8)) - 1, -5))):
            tdx = TOTAL - (f2 * 56.7 + 240) - 300
            for tot, L, o, comb in search(ths, rows, pct, tdx, [100]): cands.append((f2, L, o, comb))
        if not cands: return None
        tdx0 = TOTAL - (fmm * 56.7 + 240) - 300
        L0, o0 = layout(ths, rows, [p * tdx0 / 100 for p in pct])
        f2, L, o, comb = min(cands, key=lambda x: (x[1] + 0.6 * x[2], -x[0]))
        if (L0 - L >= FIG_MINSAVE) or (o0 and o < o0 and L <= L0):
            return dict(kind='图旁表', pct=list(comb), fig=(fmm, f2), info='图 %d→%dmm，列 %s → %s，估算 %d → %d 行，独字 %d → %d' % (fmm, f2, '/'.join(map(str, pct)), '/'.join(map(str, comb)), L0, L, o0, o))
        return None
    cur_sum = sum(pct)
    L0, o0 = layout(ths, rows, [p * TOTAL / 100 for p in pct])
    totals = sorted({cur_sum, 100})
    res = search(ths, rows, pct, TOTAL, totals)
    if not res: return None
    bestL = min(r[1] + 0.6 * r[2] for r in res)
    tot, L, o, comb = min([r for r in res if r[1] + 0.6 * r[2] <= bestL + 1e-9], key=lambda r: (abs(r[0] - cur_sum)))   # 同等最好的取总宽变化最小
    kind = None
    if L0 - L >= MINSAVE: kind = '省行'
    # 定宽表的独字交给 fit_fix 实测借宽（T8）；估算消独字会和 autow 实测结果来回改（2026-10-17）
    # 「收窄」不自动做：合并格的行高估算偏粗，试跑时会把有内容的列压到 8%～11%（2026-10-17）
    if not kind: return None
    return dict(kind=kind, pct=list(comb), fig=None, info='%s：列 %s（合计 %d%%）→ %s（合计 %d%%），估算 %d → %d 行，独字 %d → %d' % (kind, '/'.join(map(str, pct)), cur_sum, '/'.join(map(str, comb)), tot, L0, L, o0, o))

def main():
    FILES = sorted(glob.glob(os.path.join(REPO, 'notes_src', '[1-5]*', '[0-9]*.md'))) + [os.path.join(REPO, '速查', '速查源.md')]
    changes = 0
    for f in FILES:
        t = io.open(f, encoding='utf-8').read(); out = t; fn = '速查' if f.endswith('速查源.md') else os.path.basename(f).split(' ')[0]
        for m in re.finditer(r'<table[^>]*>.*?</table>', t, re.S):
            tb = m.group(0); before = t[:m.start()].rstrip().rsplit('\n', 1)[-1]
            _hm = re.search(r'<tr class="hdr">(.*?)</tr>', tb, re.S)
            if _hm and re.search(r'w-lock|keep-wide|col-eq', _hm.group(1)): continue   # SD-209：实测锁定 / 用户点名满宽 / 等宽的表不按估算改（与 check_struct T18 同口径）
            pr = propose(tb, before)
            if not pr: continue
            title = (re.findall(r'(?m)^#{3,4} ([^\n]*)$', t[:m.start()]) or ['?'])[-1]
            print('[%s] %s「%s」：%s' % (pr['kind'], fn, title[:24], pr['info']))
            hdr = re.search(r'<tr class="hdr">.*?</tr>', tb, re.S).group(0); it = iter(pr['pct'])
            hdr2 = re.sub(r'\bw-\d+', lambda mm: 'w-%d' % next(it), hdr)
            out = out.replace(tb, tb.replace(hdr, hdr2), 1)
            if pr['fig'] and pr['fig'][0] != pr['fig'][1]:
                out = out.replace(before, re.sub(r'\|\s*%d\s*$' % pr['fig'][0], '| %d' % pr['fig'][1], before), 1)
            changes += 1
        if APPLY and out != t: io.open(f, 'w', encoding='utf-8').write(out)
    print('共 %d 张表%s' % (changes, '（已写回）' if APPLY else '（未写回，加 --apply）'))
if __name__ == '__main__':
    main()
