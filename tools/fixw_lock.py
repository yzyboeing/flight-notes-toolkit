#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fixw_lock.py —— 实测来回拉锯的定宽表：按成品实测宽度 + 需加宽量定百分比并锁住（w-lock）（SD-204，2026-10-17）
SD-204 起定宽只是起始比例、按内容收放；少数表 fit_fix 收窄 / 拉满来回拉锯、收不敛（T14「定宽表右侧有空位，第 N 列可加宽」反复报）。
做法：fixw_grow 同口径找到这些表（layout_measure 的 grow 项），从成品 docx 读这张表现在的各列宽（DXA），
需要加宽的列加上实测缺口，其余列保持，换算成版心百分比写回源文件表头并加 w-lock（生成器不再收放、fit_fix 不再记账）。
用法：python3 fixw_lock.py --repo <gh-private> [--apply]"""
import argparse, io, os, re, sys, zipfile, math
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fixw_grow
TOTAL = 16838 - 1440
def docx_tables(path):
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8')
    out = []
    for m in re.finditer(r'<w:tbl>.*?</w:tbl>', x, re.S):
        tb = m.group(0)
        if '<w:tbl>' in tb[7:]: continue
        grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tb)]
        txt = re.sub(r'[\s​⁠­●•–]+', '', ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', tb)))
        out.append((txt, grid))
    return out
def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--repo', default='.'); ap.add_argument('--apply', action='store_true'); a = ap.parse_args()
    repo = os.path.abspath(a.repo)
    hits = fixw_grow.scan(repo)
    D = {'全书': docx_tables(os.path.join(repo, 'build/B737机型理论知识笔记.docx')), '单册': docx_tables(os.path.join(repo, 'build/B737机型理论知识速查.docx'))}
    done = 0
    for (f, pos), v in sorted(hits.items(), key=lambda kv: (kv[0][0], -kv[0][1])):
        t = io.open(f, encoding='utf-8').read()
        if t[pos:pos + len(v['hdr'])] != v['hdr']: print('  对不上，跳过', f); continue
        end = t.index('</table>', pos)
        key = fixw_grow.N(re.sub(r'<tr class="hdr">', '', t[pos:end]))[:40]
        cand = [g for txt, g in D[v['book']] if key and key[:30] in txt]
        ths = re.findall(r'<th\b[^>]*>', v['hdr'])
        if len(cand) != 1 or len(cand[0]) != len(ths): print('  docx 对不上，跳过', os.path.basename(f), v['page'], len(cand)); continue
        W = cand[0][:]
        for c, e in v['cols'].items():
            if c < len(W): W[c] += int(e * 20) + 60
        if sum(W) > TOTAL:   # 超出版心：从没加宽的最宽列让
            over = sum(W) - TOTAL
            for k in sorted(range(len(W)), key=lambda k: -W[k]):
                if k in v['cols']: continue
                d = min(over, W[k] // 3); W[k] -= d; over -= d
                if over <= 0: break
        pct = [max(3, math.ceil(100 * w / TOTAL)) for w in W]
        while sum(pct) > 100: pct[pct.index(max(pct))] -= 1
        k = [0]
        def sub(m):
            th = m.group(0); i = k[0]; k[0] += 1
            th = re.sub(r'\s*\bw-(\d+|lock)\b', '', th)
            th = re.sub(r'class="([^"]*)"', lambda mm: 'class="%s"' % ' '.join((mm.group(1) + ' w-%d' % pct[i] + (' w-lock' if i == 0 else '')).split()), th) if 'class=' in th else th.replace('<th', '<th class="w-%d%s"' % (pct[i], ' w-lock' if i == 0 else ''), 1)
            return th
        new = re.sub(r'<th\b[^>]*>', sub, v['hdr'])
        print('  %s 第 %d 页 %s：%s → %s' % (v['book'], v['page'], os.path.basename(f), re.findall(r'w-\d+', v['hdr']), pct))
        if a.apply and new != v['hdr']:
            io.open(f, 'w', encoding='utf-8').write(t[:pos] + new + t[pos + len(v['hdr']):]); done += 1
    print('已锁 %d 张' % done)
if __name__ == '__main__':
    main()
