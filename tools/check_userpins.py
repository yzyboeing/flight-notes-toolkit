#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_userpins.py —— 用户点名排版要求逐条在成品 PDF 上实测（第 116 条，2026-10-20 用户：说过的「定义」列居中，1020R1 没改到）
台账：笔记库根目录《用户点名排版台账.json》，每条 {书: 两本 / 完整版 / 速查, 表头: [...], 列, 要求: 居中 / 靠左加点, 首行?}
实测：按表头找到表，看该列每格——「靠左加点」每格以 • / ● 开头；「居中」不加点、单行格文字居中（±4pt）。
用法：check_userpins.py [笔记库] ——不符的打印「USERPIN 错误 …」，退出码 1；全部相符退出码 0"""
import sys, os, re, json
try: import pymupdf
except ImportError: import fitz as pymupdf
R = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser('~/flight-repos/gh-private'))
L = json.load(open(os.path.join(R, '用户点名排版台账.json'), encoding='utf-8'))
BOOKS = {'完整版': os.path.join(R, 'build', 'B737机型理论速查笔记.pdf'), '速查': os.path.join(R, 'build', 'B737机型理论速查.pdf')}
N = lambda s: re.sub(r'\s+', '', s or '')
def tables(pdf):
    d = pymupdf.open(pdf)
    for pg in d:
        for t in pg.find_tables().tables:
            rows = t.extract()
            if rows: yield pg, t, rows
errs, seen = [], {}
cache = {k: list(tables(v)) for k, v in BOOKS.items() if os.path.exists(v)}
OFF = [it for it in L if it.get('停用')]   # SD-234：停用的点名（表只在速查单册）不查，报告里单列条数
for it in [x for x in L if not x.get('停用')]:
    hdr = [N(h) for h in it['表头']]; col = hdr.index(N(it['列']))
    # SD-234（2026-10-24）：只维护《机型理论速查笔记》一本——速查单册不再构建时，「两本」只查这一本，「速查」点名的要求也落到这一本（用户规则用到唯一成品）
    QR_ON = '速查' in cache
    bks = (['完整版', '速查'] if it['书'] == '两本' else [it['书']]) if QR_ON else ['完整版']
    for bk in bks:
        hits = 0
        for pg, t, rows in cache.get(bk, []):
            h = [N(x) for x in rows[0] if x]
            if ''.join(h) != ''.join(hdr): continue
            if len(h) != len(hdr): continue
            if it.get('首行') and not any(N(it['首行']) in N(c or '') for r in rows[1:2] for c in r): continue
            hits += 1
            for ri, r in enumerate(t.rows[1:], 1):
                cells = [c for c in r.cells]
                if col >= len(cells) or cells[col] is None: continue
                if col + 1 < len(cells) and cells[col + 1] is None: continue   # 2026-10-22：跨列共有格（如再循环风扇「输送路径」）按规则居中，不算这一列的独有格
                txt = (rows[ri][col] if col < len(rows[ri]) else '') or ''
                s = txt.strip()
                if not s or re.fullmatch(r'[—\-–/无]+', s): continue
                bullet = bool(re.search(r'[•●]', s))   # PDF 表格取字时圆点可能排在格末
                if it['要求'] == '靠左加点' and not bullet:
                    errs.append('%s 第 %d 页「%s」列「%s」没有靠左加点' % (bk, pg.number + 1, it['列'], s[:16])); break
                if it['要求'] == '居中':
                    if bullet: errs.append('%s 第 %d 页「%s」列「%s」带圆点（应居中不加点）' % (bk, pg.number + 1, it['列'], s[:16])); break
                    rc = pymupdf.Rect(cells[col]); ls = [l for b in pg.get_text('dict', clip=rc)['blocks'] for l in b.get('lines', [])]
                    if len(ls) == 1:
                        cx = (ls[0]['bbox'][0] + ls[0]['bbox'][2]) / 2
                        if abs(cx - (rc.x0 + rc.x1) / 2) > 4:
                            errs.append('%s 第 %d 页「%s」列「%s」没有居中' % (bk, pg.number + 1, it['列'], s[:16])); break
        if not hits and it.get('并排'): print('USERPIN 提示 %s「%s」在图表并排里，PDF 表格识别不到，需人工看页' % (bk, it['列'])); continue
        if not hits: errs.append('%s 找不到表头「%s」的表（台账条目：%s）' % (bk, '｜'.join(it['表头']), it['来源']))
for e in errs: print('USERPIN 错误', e)
print('用户点名排版核对：%d 条（停用 %d 条，见台账「停用」），不符 %d 处' % (len(L) - len(OFF), len(OFF), len(errs)))
sys.exit(1 if errs else 0)
