#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_drift.py —— 排版漂移检查 K-B6（SD-238，2026-10-26：1026R1 只删了 6.5 第 5 条，重建后 6.2 等前面没改的表却换了压缩级、拆开）
用法：python3 check_drift.py [--repo ~/flight-repos/gh-private] [--pdf 成品.pdf] [--base 基线成品.pdf] [--tag baseline/…]
做法：取最近的基线标签（标签说明「基线 1025R3」→ 工作区《<书名>1025R3.pdf》），列出从基线到现在改过的源文件
（notes_src 下「x.y 节名.md」），找出其中最靠前的一节在基线成品里的起始页；正文从第一章起逐页对齐，
这一页之前的每一页页首、页尾文字（去掉空白与数字）必须与基线相同——源文件没改的部分，版面不该变。
不同的页报「第 N 页 …」（退出码 1）。改的是排版工具本身、版面本来就要变时，设 DRIFT_OK=1 只提示。"""
import os, re, sys, subprocess
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf


def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default


REPO = os.path.abspath(os.path.expanduser(arg('--repo', '~/flight-repos/gh-private')))
STEM = 'B737机型理论速查笔记'
PDF = arg('--pdf') or os.path.join(REPO, 'build', STEM + '.pdf')
WS = os.path.expanduser('~/Desktop/飞行理论笔记整理')
FOOT = re.compile(r'第\s*\d+\s*页\s*$')


def git(*a):
    return subprocess.run(['git', '-c', 'core.quotepath=off'] + list(a), cwd=REPO, capture_output=True, text=True).stdout


def body0(d):
    toc = d.get_toc()
    dir0 = next((p for lv, t, p in toc if lv == 1 and t.replace('​', '').strip() == '速查目录'), 0)
    return next((p for lv, t, p in toc if lv == 1 and p > dir0 and re.match(r'^\s*第', t.replace('​', ''))), None) \
        or next((p for lv, t, p in toc if lv == 1 and p > dir0), 1)


def sig(pg):
    ls = [l for l in pg.get_text().split('\n') if l.strip() and not FOOT.search(l.strip())]
    s = re.sub(r'[\s\d]+', '', ''.join(ls))
    return s[:40], s[-40:]


def main():
    tag = arg('--tag') or (git('tag', '-l', 'baseline/*', '--sort=-creatordate').split() or [None])[0]
    if not tag: print('排版漂移检查：没有基线标签，跳过'); return 0
    ed = re.search(r'(\d{4}R\d+)', git('tag', '-l', tag, '--format=%(contents:subject)'))
    base = arg('--base') or (ed and os.path.join(WS, '%s%s.pdf' % (STEM, ed.group(1))))
    if not base or not os.path.exists(base) or not os.path.exists(PDF):
        print('提示 排版漂移检查：找不到基线成品（%s）或当前成品，跳过' % (base or tag)); return 0
    files = [f for f in git('diff', '--name-only', tag, '--', 'notes_src').splitlines() if f.endswith('.md')]
    files += [f for f in git('ls-files', '--others', '--exclude-standard', 'notes_src').splitlines() if f.endswith('.md')]
    secs, unknown = set(), []
    for f in files:
        m = re.match(r'^(\d+\.\d+)\s', os.path.basename(f))
        (secs.add(m.group(1)) if m else unknown.append(f))
    b, n = pymupdf.open(base), pymupdf.open(PDF)
    b0, n0 = body0(b), body0(n)
    cut = len(b) + 1
    for f in unknown:   # 章目录页 / 章首简介等：从该章第一页起算改动；认不出章的（附件、MANIFEST）从正文第一页起算
        nm = re.sub(r'^\d+\s*', '', os.path.basename(os.path.dirname(f)))
        ch = nm and next((p for lv, t, p in b.get_toc() if lv == 1 and p >= b0 and nm.split()[-1] in t), None)
        cut = min(cut, ch or b0)
    for lv, t, p in b.get_toc():
        m = re.match(r'^\s*(\d+\.\d+)\s', t.replace('​', ''))
        if lv == 2 and m and m.group(1) in secs and p >= b0: cut = min(cut, p)
    bad = []
    for p in range(b0, cut):
        q = p - b0 + n0
        if q > len(n): bad.append((p, q, '当前成品没有这一页')); continue
        sb, sn = sig(b[p - 1]), sig(n[q - 1])
        if sb != sn: bad.append((p, q, '页首' if sb[0] != sn[0] else '页尾'))
    soft = os.environ.get('DRIFT_OK') == '1'
    for p, q, why in bad[:15]:
        print('%s第 %d 页（基线第 %d 页）%s与基线不同——源文件没改的部分版面变了' % ('提示 ' if soft else '', q, p, why))
    print('排版漂移检查：基线 %s，改动最早在基线第 %s 页（%s），之前正文 %d 页，漂移 %d 页%s'
          % (ed.group(1), cut if cut <= len(b) else '—', '、'.join(sorted(secs, key=lambda s: [int(x) for x in s.split('.')])) or '无源文件改动',
             max(0, cut - b0), len(bad), '（DRIFT_OK=1，只提示）' if soft and bad else ''))
    return 1 if bad and not soft else 0


if __name__ == '__main__':
    sys.exit(main())
