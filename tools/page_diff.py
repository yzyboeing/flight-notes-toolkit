#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""page_diff.py —— 变化页对比总览（2026-10-22 用户：「减少你人工过目的负担：每轮只给你这一轮变化过的页的对比总览，不用每次整本翻」）
用法：python3 page_diff.py <旧.pdf> <新.pdf> <输出目录> [--tag 速查]
做法：两版逐页取正文文字（去页脚），按文字序列对齐（difflib），找出内容或排版变化过的页：
  · 两版都有、文字或版面（文字块位置）不同 → 左旧右新并排；· 新增页 → 只放新页；· 删掉的页 → 只放旧页（标「已删」）。
输出：<输出目录>/<tag>_变化页_01.png …（每张 4 对，旧 / 新并排，标页码）＋ <tag>_变化页清单.md。"""
import sys, os, re, difflib
import pymupdf

OLD, NEW, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
TAG = sys.argv[sys.argv.index('--tag') + 1] if '--tag' in sys.argv else os.path.splitext(os.path.basename(NEW))[0][-2:]
os.makedirs(OUT, exist_ok=True)
for f in os.listdir(OUT):
    if f.startswith(TAG + '_变化页'): os.remove(os.path.join(OUT, f))


def sig(p):
    t = re.sub(r'(?m)^.*第\s*\d+\s*页\s*$', '', p.get_text())
    t = re.sub(r'\s+', '', t)
    lay = tuple(round(b[1] / 6) for b in p.get_text('blocks') if b[4].strip())[:40]   # 版面：文字块纵向位置（6pt 粒度）
    return t, lay


a, b = pymupdf.open(OLD), pymupdf.open(NEW)
sa, sb = [sig(p) for p in a], [sig(p) for p in b]
sm = difflib.SequenceMatcher(None, [x[0] for x in sa], [x[0] for x in sb], autojunk=False)
pairs = []   # (旧页号或 None, 新页号或 None, 说明)
for op, i1, i2, j1, j2 in sm.get_opcodes():
    if op == 'equal':
        for i, j in zip(range(i1, i2), range(j1, j2)):
            if sa[i][1] != sb[j][1]: pairs.append((i, j, '版面变化'))
    elif op == 'replace':
        for k in range(max(i2 - i1, j2 - j1)):
            i = i1 + k if i1 + k < i2 else None; j = j1 + k if j1 + k < j2 else None
            pairs.append((i, j, '内容变化' if i is not None and j is not None else ('新增页' if i is None else '已删页')))
    elif op == 'insert':
        for j in range(j1, j2): pairs.append((None, j, '新增页'))
    elif op == 'delete':
        for i in range(i1, i2): pairs.append((i, None, '已删页'))

TW = 360
def thumb(doc, n):
    p = doc[n]; sc = TW / p.rect.width
    return p.get_pixmap(matrix=pymupdf.Matrix(sc, sc), alpha=False)

PER = 4
lines = ['# %s 变化页清单' % TAG, '', '旧版 %s（%d 页）→ 新版 %s（%d 页）；变化 %d 处。' % (os.path.basename(OLD), len(a), os.path.basename(NEW), len(b), len(pairs)), '',
         '| 旧页 | 新页 | 变化 |', '|---|---|---|']
for i, j, why in pairs: lines.append('| %s | %s | %s |' % (i + 1 if i is not None else '—', j + 1 if j is not None else '—', why))
open(os.path.join(OUT, TAG + '_变化页清单.md'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
th = int(a[0].rect.height * TW / a[0].rect.width) if len(a) else 255
for k in range(0, len(pairs), PER):
    grp = pairs[k:k + PER]
    W, H = 2 * TW + 30, len(grp) * (th + 34) + 10
    sh = pymupdf.open(); pg = sh.new_page(width=W, height=H)
    for r, (i, j, why) in enumerate(grp):
        y = 10 + r * (th + 34)
        pg.insert_text((10, y + 14), '旧 %s' % (i + 1 if i is not None else '—'), fontsize=11, fontname='china-s', color=(0.4, 0.4, 0.4))
        pg.insert_text((TW + 20, y + 14), '新 %s　%s' % (j + 1 if j is not None else '—', why), fontsize=11, fontname='china-s', color=(0.7, 0, 0))
        for col, (doc, n) in enumerate(((a, i), (b, j))):
            x = 10 + col * (TW + 10)
            rect = pymupdf.Rect(x, y + 20, x + TW, y + 20 + th)
            if n is None: pg.draw_rect(rect, color=(0.8, 0.8, 0.8)); continue
            pg.insert_image(rect, pixmap=thumb(doc, n)); pg.draw_rect(rect, color=(0.6, 0.6, 0.6), width=0.6)
    pg.get_pixmap(matrix=pymupdf.Matrix(1.6, 1.6), alpha=False).save(os.path.join(OUT, '%s_变化页_%02d.png' % (TAG, k // PER + 1)))
print('%s：变化 %d 处 → %s（%d 张对比图）' % (TAG, len(pairs), OUT, (len(pairs) + PER - 1) // PER))
