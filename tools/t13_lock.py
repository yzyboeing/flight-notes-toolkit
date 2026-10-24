#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""t13_lock.py —— 按排版检查报告的 T13（列没有折行、右侧仍空）把收不动的表按实测收窄并锁宽（SD-208，2026-10-18）
fit_fix 不动锁宽表、对起始比例表只按严格口径收，个别自动宽度表也会被旧记录卡住，T13 连续几轮不消。
做法：从报告取（书、页、列、空白 pt）→ layout_measure 找到那张表的文字 → 在成品 docx 里按文字找到表、读各列宽（DXA）
→ 该列减去「空白 − 6pt」→ 换算成版心百分比（合计不足 100 的保持窄表）→ 在源文件里按表格文字唯一定位，表头写 w-NN、首格加 w-lock。
对不上或不唯一的只列出不改。用法：python3 t13_lock.py --repo <gh-private> [--apply]"""
import argparse, io, os, re, sys, glob, zipfile, math
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import layout_measure as LM
TOTAL = 16838 - 1440
N = lambda s: re.sub(r'[^\w]', '', re.sub(r'<[^>]+>', '', s.replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&')))

def docx_tables(path):
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8'); out = []
    for m in re.finditer(r'<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>', x, re.S):
        tb = m.group(0)
        grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tb)]
        out.append((N(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tb))), grid))
    return out

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--repo', default='.'); ap.add_argument('--apply', action='store_true'); a = ap.parse_args()
    repo = os.path.abspath(a.repo)
    rep = io.open(os.path.join(repo, 'build/排版检查报告.md'), encoding='utf-8').read().split('## 新建议')[0]
    errs = [(bk, int(pg), int(col) - 1, float(bl)) for bk, pg, col, bl in re.findall(r'\[T13\] (全书|单册) 第 (\d+) 页：第 (\d+) 列（.*?右侧仍空约 (\d+)pt', rep, re.S)]
    books = {'全书': ('build/B737机型理论速查笔记.pdf', 'build/B737机型理论速查笔记.docx', sorted(glob.glob(os.path.join(repo, 'notes_src/[1-6]*/[0-9]*.md')))),
             '单册': ('build/B737机型理论速查.pdf', 'build/B737机型理论速查.docx', [os.path.join(repo, '速查/速查源.md')])}
    meas, dx = {}, {}
    plan = {}   # (文件, 源表头) -> (表 grid, {列: 收窄 DXA})
    for bk, pg, col, bl in errs:
        pdf, docx, srcs = books[bk]
        if bk not in meas: meas[bk] = LM.measure(os.path.join(repo, pdf), srcmd=os.path.join(repo, 'build', 'book.md' if bk == '全书' else 'qr_single.md')); dx[bk] = docx_tables(os.path.join(repo, docx))
        xs = [x for x in meas[bk] if x['page'] == pg and x['col'] == col and x.get('table')]
        x = min(xs, key=lambda x: abs(x.get('extra_strict', x['extra_pt']) - bl)) if xs else None   # 同页可能有几张表的同号列：按实测空白量最接近的那张
        if not x: print('  测不到', bk, pg); continue
        key = N(x['table'])[:30]
        # 2026-10-22（SD-223）：压缩过的表（本次运行字号 < 9pt）不锁宽——压缩只在本次运行有效，按压缩后字宽锁的列宽下次会对不上；交给 fit_fix 每次运行内实测收窄
        try:
            _kf = io.open(os.path.join(repo, 'build', 'keep_force_%s.txt' % os.path.splitext(os.path.basename(pdf))[0]), encoding='utf-8').read()
            if any(N(l.split(':', 1)[1])[:30] == key for l in _kf.splitlines() if re.match(r'Z?~\d:', l)): print('  压缩表不锁宽：%s 第 %d 页 %s' % (bk, pg, key[:16])); continue
        except OSError: pass
        cand = [g for t, g in dx[bk] if t.startswith(key)]
        hits = []
        for f in srcs:
            t = io.open(f, encoding='utf-8').read()
            for m in re.finditer(r'<table[^>]*>.*?</table>', t, re.S):
                tb = re.sub(r'<tr class="(premise|note|warn)">.*?</tr>', '', m.group(0), flags=re.S)
                if N(tb).startswith(key): hits.append((f, re.search(r'<tr class="hdr">.*?</tr>', m.group(0), re.S).group(0), m.start()))
        if len(cand) != 1 or len(hits) != 1: print('  对不上（docx %d / 源 %d），跳过：%s 第 %d 页 %s' % (len(cand), len(hits), bk, pg, key[:16])); continue
        grid = cand[0]; f, hdr, pos = hits[0]
        if 'col-eq' in hdr or 'keep-wide' in hdr: print('  等宽 / 点名满宽表，不收：%s 第 %d 页' % (bk, pg)); continue   # 与 check_layout T13 同口径
        if len(re.findall(r'<th\b', hdr)) != len(grid) or 'colspan' in hdr: print('  表头有跨列或列数不符，跳过：%s 第 %d 页' % (bk, pg)); continue
        k = (f, hdr, pos); g0, cut = plan.get(k, (grid, {}))
        cut[col] = max(cut.get(col, 0), int((bl - 14) * 20))   # 与 fit_fix 定宽起始表同口径多留 14pt（粗体 / 句末标点实宽偏大）
        plan[k] = (g0, cut)
    done = 0
    for (f, hdr, pos), (grid, cut) in sorted(plan.items(), key=lambda kv: -kv[0][2]):   # 同一文件从后往前写，位置不乱
        W = grid[:]
        for c, d in cut.items(): W[c] = max(600, W[c] - d)
        pct = [max(3, round(100 * w / TOTAL)) for w in W]
        while sum(pct) > 100: pct[pct.index(max(pct))] -= 1
        i = [0]
        def sub(m):
            th = m.group(0); q = i[0]; i[0] += 1
            th = re.sub(r'\s*\bw-(\d+|lock)\b', '', th)
            add = ' w-%d' % pct[q] + (' w-lock' if q == 0 else '')
            return re.sub(r'class="([^"]*)"', lambda mm: 'class="%s"' % ' '.join((mm.group(1) + add).split()), th) if 'class=' in th else th.replace('<th', '<th class="%s"' % add.strip(), 1)
        new = re.sub(r'<th\b[^>]*>', sub, hdr)
        print('  %s：%s → %s' % (os.path.basename(f), [round(100 * w / TOTAL) for w in grid], pct))
        if a.apply:
            t = io.open(f, encoding='utf-8').read()
            i0 = t.index(hdr, pos)   # 只改匹配到的那张表的表头
            io.open(f, 'w', encoding='utf-8').write(t[:i0] + new + t[i0 + len(hdr):]); done += 1
    print('已锁 %d 张' % done)

if __name__ == '__main__':
    main()
