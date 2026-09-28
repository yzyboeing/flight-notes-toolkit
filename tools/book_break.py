#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""book_break.py —— 全书 iPad 竖版预排版：找出被拆到两页的条目，给其标题加「另起一页」

用法（每轮三步，在笔记主库下运行）：
  DOC_PORTRAIT=1 BREAK_IDX=$(cat build/breaks.txt) node ../pub/tools/build_docx.js build/book.md build/X.docx
  soffice --headless --convert-to pdf --outdir build build/X.docx
  python3 ../pub/tools/book_break.py build/X.pdf build/book.md build/breaks.txt
条目按全书出现顺序编号（与 build_docx.js 的 ITEM_IDX 一致）。条目内容越过页边界、且标题不在页顶的 → 加入另起一页；
标题已在页顶仍跨页的（一页放不下）→ 移出名单。breaks.txt 不再变化即收敛。
"""
import re, sys, io, os
import pymupdf

PDF, BOOK, STATE = sys.argv[1], sys.argv[2], sys.argv[3]
norm = lambda s: re.sub(r'[ \t\n]+', '', s)          # 保留全角空格：标题「A-1　名」与索引「A-1 名」靠它区分
book = io.open(BOOK, encoding='utf-8').read()
heads = [m.group(2).strip() for m in re.finditer(r'(?m)^(#{1,6}) (.+)$', book)]
items = [h for h in heads if re.match(r'^(\d+\. |[A-Z]-\d+　)', h)]
headset = {norm(h)[:12] for h in heads}
prev = set(int(x) for x in open(STATE).read().split(',') if x.strip()) if os.path.exists(STATE) else set()

d = pymupdf.open(PDF)
k = 0; cur = None; cur_top = False; split, big = set(), set()
for p in d:
    ls = [l.strip() for l in p.get_text().split('\n') if l.strip() and not re.fullmatch(r'第\s*\d+\s*页', l.strip())]
    if not ls: continue
    first = norm(ls[0])[:12]
    if cur is not None and first not in headset:
        (big if cur_top else split).add(cur)
    for i, l in enumerate(ls):
        if norm(l)[:12] in headset and not (k < len(items) and norm(l).startswith(norm(items[k])[:12])):
            if not re.match(r'^(\d+\. |[A-Z]-\d+)', l): cur = None          # 进入新的节 / 块标题：上一条已结束
            continue
        if k < len(items) and norm(l).startswith(norm(items[k])[:12]):
            k += 1; cur = k; cur_top = (i == 0 or all(norm(x)[:12] in headset for x in ls[:i]))
BIGF = STATE + '.big'                                   # 一页放不下的条目：永久移出名单，避免来回摆动
oldbig = set(int(x) for x in open(BIGF).read().split(',') if x.strip()) if os.path.exists(BIGF) else set()
allbig = oldbig | big
new = (prev | split) - allbig
io.open(STATE, 'w').write(','.join(str(x) for x in sorted(new)))
io.open(BIGF, 'w').write(','.join(str(x) for x in sorted(allbig)))
print('匹配条目 %d / %d；跨页 %d 条（其中一页放不下 %d 条）；另起一页名单 %d → %d%s'
      % (k, len(items), len(split) + len(big), len(big), len(prev), len(new), '（已收敛）' if new == prev else ''))
