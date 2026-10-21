#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_splits.py —— 成品 PDF 断页检查（2026-09-29，SD-71）
用法：python3 check_splits.py <全书.pdf> [--all]
逐页用 PyMuPDF 识别表格，找出「页底表格接到下一页页顶继续」的断表，按 SD-218（2026-10-21 用户：表格可跨页，
「每截至少 3 行」「续页重复表头」）报坏断点（fit_fix 修、check_layout 报错误，同一口径）：
  · 截行不足：上截或下截的数据行（不算表头、续页重复的表头）少于 3 行；上截 0 行＝页底只剩表头；
  · 页底分组行：上截最后一行是通栏分组行（只有一格有字）；
  · 断在合并格：下截第一行数据的首列是空的（纵向合并格被拆到两页）；
  · 标题孤立：页底最后一段是条目标题，正文到了下一页。
  另记 fits（两截连同表上方标题合起来放得进一页）供 fit_fix 选修法：放得下 → 整表移下页；放不下 → 改断点。
默认只列有问题的；--all 列出全部断表；--json 输出机器可读结果（fit_fix.py 用）。只读，不改文件。"""
import sys, re
try:
    import pymupdf
except ImportError:  # PyMuPDF 旧版模块名
    import fitz as pymupdf

PDF = sys.argv[1]; ALL = '--all' in sys.argv; JSON = '--json' in sys.argv
import json
norm = lambda x: ''.join(ch for ch in str(x or '') if ch.isalnum())
recs = []
EXT = {}
import os
MINR = int(os.environ.get('SPLIT_MINROWS', '1'))   # SD-218：每截至少几行数据（2026-10-21 用户定 1；仍不许页底只剩表头、分组行留页底、拆合并格）

d = pymupdf.open(PDF)
info = []
for i, pg in enumerate(d):
    H = pg.rect.height
    try:
        tabs = pg.find_tables().tables
    except Exception:
        tabs = []
    blocks = [b for b in pg.get_text('blocks') if b[4].strip()]
    # 去掉页眉页脚：按内容识别（页码行、页眉书名 / 章名是本页最上面、且整块在版心之上的那一块），不写死坐标——页边距改了也适用
    # SD-97：页脚一行含章名 / 节名 / 页码，可能分成几个块——以页码块所在高度为准，同一行及以下都算页脚
    fy = max((x[1] for x in blocks if re.search(r'第\s*\d+\s*页\s*$', x[4])), default=H)
    foot = [x for x in blocks if x[1] >= fy - 2]
    tops = sorted(blocks, key=lambda x: x[1])
    hy = tops[0][3] if tops and tops[0][3] < 0.08 * H and tops[0] not in foot else 0
    body = [x for x in blocks if x[1] >= hy and x[3] <= fy and x not in foot and (hy == 0 or x is not tops[0])]
    # 页眉 / 页脚的横线会被识别成一行「表格」，剔掉
    tabs = [t for t in tabs if t.bbox[1] >= hy - 2 and t.bbox[3] <= fy + 2 and t.bbox[3] - t.bbox[1] > 4]
    top = min((b[1] for b in body), default=H); bot = max((b[3] for b in body), default=0)
    EXT.update({id(t): (t.extract() or []) for t in tabs})   # 2026-10-21：当场取表格文字——循环结束后页面对象已换，t.extract() 会读到别的页（曾把断表误当表后注跳过）
    info.append({'tabs': tabs, 'top': top, 'bot': bot, 'H': H, 'blocks': body,
                 'idx': any(re.match(r'\s*(按主题查|主题目录|目录)', x[4]) for x in foot)})   # SD-139：目录 / 按主题查页的条目是索引行，不是标题
AREA = max(x['bot'] - x['top'] for x in info if x['blocks']) if info else 0
problems, allsplits = [], []
head_re = re.compile(r'^\s*(\d{1,3}\.\s+\S|[A-Z]-\d+\s|\d\.\d+\s|块索引\s*$)')   # 条目标题、节标题、「块索引」小标题
for i in range(len(info) - 1):
    a, b = info[i], info[i + 1]
    # 标题孤立：本页最后一个文字块是条目标题
    if a['blocks'] and not a['idx']:
        last = max(a['blocks'], key=lambda x: x[3])
        txt = last[4].strip().split('\n')[0]
        if head_re.match(txt) and len(txt) < 40 and not any(t.bbox[3] > last[1] for t in a['tabs']):
            problems.append('第 %d 页页底只有标题「%s」，正文在下一页' % (i + 1, txt))
    if not a['tabs'] or not b['tabs']:
        continue
    ta = max(a['tabs'], key=lambda t: t.bbox[3]); tb = min(b['tabs'], key=lambda t: t.bbox[1])
    # 上页表格贴到版心底部、下页表格从版心顶部开始、中间没有别的文字 → 视为同一张表断开
    if ta.bbox[3] < a['bot'] - 6 or tb.bbox[1] > b['top'] + 6:
        continue
    def cell0(t, last=False):
        try:
            rows = EXT.get(id(t)) or [['']]; r = rows[-1] if last else rows[0]
            return ''.join(str(x or '') for x in r).strip()
        except Exception: return ''
    TAILRE = re.compile(r'^(注|解释|出处|公司差异)[：:]')
    firstb = min(b['blocks'], key=lambda x: x[1])[4].strip() if b['blocks'] else ''
    if TAILRE.match(cell0(tb)) or TAILRE.match(cell0(ta)) or TAILRE.match(firstb):
        continue                                    # 表后说明段（带左竖线，会被识别成单行表），不是表格本体断开
    # 同一张表断开时两截的列分界线完全一致；不一致说明是上一张表恰好排到页底、下一页另起一张表（不是断表）
    def inner(t):
        xs = {round(c[k]) for r in t.rows for c in r.cells if c for k in (0, 2)}
        return sorted(x for x in xs if 45 < x < pg0w - 45)
    pg0w = d[i].rect.width
    ia, ib = inner(ta), inner(tb)
    if len(ia) != len(ib) or any(abs(x - y) > 3 for x, y in zip(ia, ib)):
        continue
    ra, rb = ta.row_count, tb.row_count
    rows_a, rows_b = EXT.get(id(ta)) or [], EXT.get(id(tb)) or []
    rowtxt = lambda r: norm(''.join(str(x or '') for x in r))
    rep = 0   # 续页重复的表头行数（多行表头整组重复，如「EEC 选择条件」下再分「737-NG ｜ 737-8」）
    while rep < 3 and rep < len(rows_a) and rep < len(rows_b) - 1 and rowtxt(rows_a[rep]) == rowtxt(rows_b[rep]): rep += 1
    da, db = ra - rep, rb - rep                                                     # 两截的数据行数
    filled = lambda r: [x for x in r if str(x or '').strip()]
    group_tail = bool(rows_a) and ta.col_count >= 2 and len(filled(rows_a[-1])) == 1 and len(rows_a[-1]) >= 2 and not str(rows_a[-1][0] or '').strip() == ''
    c0a = [str(r[0] or '').strip() for r in rows_a[rep:]] if rows_a else []
    c0b = [str(r[0] or '').strip() for r in rows_b[rep:]] if rows_b else []
    # 断在合并格：下截第一行首列空、上截首列有字（首列整列没字的是图标列，不算）；span_back＝上截末尾属于这一组的行数（从组首行起）
    span_cut = ta.col_count >= 2 and bool(c0b) and not c0b[0] and any(c0a) and len(filled(rows_b[rep])) >= 1
    span_back = 0
    if span_cut:
        for x in reversed(c0a):
            span_back += 1
            if x: break
    ha, hb = ta.bbox[3] - ta.bbox[1], tb.bbox[3] - tb.bbox[1]
    # 条目标题与表前说明必须跟表同页：表上方（同页）最近的条目标题到表顶这段高度也要算进去
    heads = [x for x in a['blocks'] if x[3] <= ta.bbox[1] + 1 and head_re.match(x[4].strip().split('\n')[0])]
    above = ta.bbox[1] - max(heads, key=lambda x: x[1])[1] if heads else 0
    if above > AREA * 0.3: above = 0
    name = ' '.join(d[i].get_text(clip=ta.rows[0].bbox).split())[:40] if ra else ''
    msg = '第 %d→%d 页：上截 %d 行 / 下截 %d 行（表头：%s）' % (i + 1, i + 2, ra, rb, name)
    allsplits.append(msg)
    why = []
    fits = above + ha + hb <= AREA * 0.97
    if da <= 0: why.append('页底只剩表头')
    elif da < MINR or db < MINR: why.append('截行不足（上截 %d 行 / 下截 %d 行数据，每截至少 %d 行）' % (da, db, MINR))
    if group_tail: why.append('页底是分组行')
    if span_cut: why.append('断在合并格中间')
    orph = bool(why)
    if orph and not fits and da + db < 2 * MINR: why.append('（表本身超过一页、行数太少无法满足，人工看）')
    msg = '第 %d→%d 页：上截 %d 行 / 下截 %d 行数据（表头：%s）' % (i + 1, i + 2, da, db, name)
    recs.append({'page': i + 1, 'ra': ra, 'rb': rb, 'da': da, 'db': db, 'rep': rep, 'group_tail': group_tail, 'span_cut': span_cut, 'span_back': span_back, 'first_b': norm(''.join(str(x or '') for x in rows_b[rep]))[:24] if len(rows_b) > rep else '', 'last_a': norm(''.join(str(x or '') for x in rows_a[-1]))[:24] if rows_a else '', 'fits': fits, 'orphan': orph, '_ta': id(ta), '_tb': id(tb), 'over': round((above + ha + hb) / AREA, 3), 'msg': msg, 'why': why,
                 'text': norm(d[i].get_text(clip=ta.bbox)) + norm(d[i + 1].get_text(clip=tb.bbox))})
byp = {r['page']: r for r in recs}
for r in recs:
    # 跨 3 页以上的同一张表：本页页顶续接的表就是本页页底又断开的那张（2026-10-21：原来只看相邻两页都有断表，不同的表也被当成一张）
    prv, nxt = byp.get(r['page'] - 1), byp.get(r['page'] + 1)
    r['multi'] = bool((prv and prv['_tb'] == r['_ta']) or (nxt and r['_tb'] == nxt['_ta']))
    if r['multi'] and r['fits']: r['fits'] = False
    if r['why']: problems.append(r.pop('msg') + '　→ ' + '、'.join(r['why']))
for r in recs: r.pop('_ta', None); r.pop('_tb', None)
if JSON: print(json.dumps({'pages': len(d), 'splits': recs, 'problems': problems}, ensure_ascii=False)); sys.exit(0)
print('共 %d 页；断表 %d 处；问题 %d 处' % (len(d), len(allsplits), len(problems)))
for m in (allsplits if ALL else problems): print('  ' + m)
