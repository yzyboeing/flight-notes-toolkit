#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""quickref_pdf.py —— 速查区 iPad 版 PDF：书签 + 可点目录页（SD-23）

用法（笔记库根目录，先 assemble 再用竖版渲染）：
    DOC_PORTRAIT=1 node <工具链>/tools/build_docx.js build/mod0.md build/速查区_iPad.docx
    soffice --headless --convert-to pdf build/速查区_iPad.docx --outdir build
    python3 <工具链>/tools/quickref_pdf.py build/速查区_iPad.pdf [--src notes_src]

做三件事：去掉 Word 自动目录留下的空白目录页；按「块 → 条目」两级写入 PDF 书签；
在封面后插入一页可点击的块目录（块名 · 条目编号范围 · 页码）。页码按 PDF 实际页找，不估算。
"""
import io, os, re, sys, glob
try:
    import pymupdf
except ImportError:  # PyMuPDF 旧版模块名
    import fitz as pymupdf

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
PDF = sys.argv[1]
SRC = os.path.abspath(_arg('--src', 'notes_src'))

q = io.open(glob.glob(os.path.join(SRC, '0 *', '0 *.md'))[0], encoding='utf-8').read()
blocks = []                                   # [(块名, [(n, 标题)])]
for m in re.finditer(r'(?m)^(##|###) (.+)$', q):
    if m.group(1) == '##': blocks.append((m.group(2).strip(), []))
    else:
        mm = re.match(r'(\d+)\. (.+)$', m.group(2))
        if mm and blocks: blocks[-1][1].append((int(mm.group(1)), mm.group(2).strip()))

doc = pymupdf.open(PDF)
# 1 去掉空白目录页（只有「目　录」和页码）
for i in range(min(4, doc.page_count) - 1, -1, -1):
    txt = re.sub(r'第\s*\d+\s*页|\s', '', doc[i].get_text())
    if txt in ('目录', '目　　录'.replace('　', '')):
        doc.delete_page(i)
# 2 找每条所在页（从前往后，只接受不早于上一条的页）
def find(prefix, start):
    for p in range(start, doc.page_count):
        for ln in doc[p].get_text().split('\n'):
            if ln.strip().startswith(prefix): return p
    return None
pages, bpage, cur = {}, {}, 0
for bt, its in blocks:
    if not its:                                   # 没有编号条目的块（如「按物理量索引」）：按块标题找页
        p = find(bt, cur)
        if p is None: sys.exit('找不到块「%s」所在页' % bt)
        bpage[bt] = cur = p; continue
    for n, t in its:
        p = find('%d. ' % n, cur)
        if p is None: sys.exit('找不到第 %d 条所在页' % n)
        pages[n] = cur = p
    bpage[bt] = pages[its[0][0]]
# 3 插入目录页（封面之后）
W, H = doc[0].rect.width, doc[0].rect.height
idx = doc.new_page(pno=1, width=W, height=H)
# 中文用内置 china-s、数字与西文用 helv，分段绘制：两者都是 PDF 内置字体，不嵌入字形，文件不会变大
def draw(page, x, y, text, size, color=(0, 0, 0)):
    for seg in re.findall(r'[\x20-\x7e–]+|[^\x20-\x7e–]+', text):
        f = 'helv' if re.match(r'[\x20-\x7e–]', seg) else 'china-s'
        if f == 'helv': seg = seg.replace('–', '-')
        page.insert_text((x, y), seg, fontname=f, fontsize=size, color=color)
        x += pymupdf.get_text_length(seg, fontname=f, fontsize=size)
shift = lambda p: p + 1 if p >= 1 else p     # 插页后原第 1 页起全部后移
draw(idx, 56, 70, '速查区目录', 18)
draw(idx, 56, 92, '点块名跳转；条目按全章连续编号。', 9, (0.4, 0.4, 0.4))
y, LH = 124, (H - 180) / max(1, len(blocks))
LH = min(LH, 17)
for bt, its in blocks:
    if its:
        a, b = its[0][0], its[-1][0]
        rng = str(a) if a == b else '%d–%d' % (a, b)
    else:
        rng = ''
    tp = shift(bpage[bt])
    draw(idx, 56, y, bt, 10.5)
    draw(idx, W - 200, y, rng, 10.5, (0.35, 0.35, 0.35))
    draw(idx, W - 90, y, str(tp + 1), 10.5, (0.35, 0.35, 0.35))
    idx.insert_link({'kind': pymupdf.LINK_GOTO, 'from': pymupdf.Rect(50, y - 12, W - 50, y + 4), 'page': tp})
    y += LH
# 4 书签：块 → 条目
toc = [[1, '速查区目录', 2]]
for bt, its in blocks:
    toc.append([1, bt, shift(bpage[bt]) + 1])
    for n, t in its:
        toc.append([2, '%d. %s' % (n, t), shift(pages[n]) + 1])
doc.set_toc(toc)
out = PDF[:-4] + '_书签.pdf'
open(out, 'wb').write(doc.tobytes(garbage=4, deflate=True))   # 直接覆盖写，不走删除重建
print('写入 %s：%d 页，书签 %d 个（%d 块 %d 条）' % (out, doc.page_count, len(toc), len(blocks), sum(len(i) for _, i in blocks)))
