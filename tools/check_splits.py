#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_splits.py —— 成品 PDF 断页检查（2026-09-29，SD-71）
用法：python3 check_splits.py <全书.pdf> [--all]
逐页用 PyMuPDF 识别表格，找出「页底表格接到下一页页顶继续」的断表，报告：
  · 孤行：断开后上半截或下半截只有 1～2 行、且高度不到版心的 1/4；
  · 本可整页：两截高度之和（连同表上方同页的条目标题与表前说明）不超过一页可用高度（说明这张表本来放得进一页）。
  · 标题孤立：页底最后一段是条目标题（数字编号或 A-1 这类），正文到了下一页。
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
    info.append({'tabs': tabs, 'top': top, 'bot': bot, 'H': H, 'blocks': body,
                 'idx': any(re.match(r'\s*(按主题查|目录)', x[4]) for x in foot)})   # SD-139：目录 / 按主题查页的条目是索引行，不是标题
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
            rows = t.extract(); r = rows[-1] if last else rows[0]
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
    ha, hb = ta.bbox[3] - ta.bbox[1], tb.bbox[3] - tb.bbox[1]
    # 条目标题与表前说明必须跟表同页：表上方（同页）最近的条目标题到表顶这段高度也要算进去
    heads = [x for x in a['blocks'] if x[3] <= ta.bbox[1] + 1 and head_re.match(x[4].strip().split('\n')[0])]
    above = ta.bbox[1] - max(heads, key=lambda x: x[1])[1] if heads else 0
    if above > AREA * 0.3: above = 0
    name = ' '.join(d[i].get_text(clip=ta.rows[0].bbox).split())[:40] if ra else ''
    msg = '第 %d→%d 页：上截 %d 行 / 下截 %d 行（表头：%s）' % (i + 1, i + 2, ra, rb, name)
    allsplits.append(msg)
    why = []
    orph = (ra <= 2 and ha < AREA * 0.25) or (rb <= 2 and hb < AREA * 0.25)   # 行数少但每行很高（大段处置）不算孤行
    if re.sub(r'\s', '', name).startswith('块主题条目') and ra >= 2 and rb >= 2: orph = False   # 块索引：一行一块，两截各 ≥ 2 块即可（SD-87）
    if orph: why.append('孤行')
    fits = above + ha + hb <= AREA * 0.97
    if fits: why.append('本可整页')
    recs.append({'page': i + 1, 'ra': ra, 'rb': rb, 'fits': fits, 'orphan': orph, 'over': round((above + ha + hb) / AREA, 3), 'msg': msg, 'why': why,
                 'text': norm(d[i].get_text(clip=ta.bbox)) + norm(d[i + 1].get_text(clip=tb.bbox))})
pgs = {r['page'] for r in recs}
for r in recs:
    r['multi'] = (r['page'] - 1) in pgs or (r['page'] + 1) in pgs   # 连续两页都断：跨 3 页以上的大表，只看两截会误判「放得下」
    if r['multi'] and r['fits']: r['fits'] = False; r['why'].remove('本可整页')
    if r['why']: problems.append(r.pop('msg') + '　→ ' + '、'.join(r['why']))
if JSON: print(json.dumps({'pages': len(d), 'splits': recs, 'problems': problems}, ensure_ascii=False)); sys.exit(0)
print('共 %d 页；断表 %d 处；问题 %d 处' % (len(d), len(allsplits), len(problems)))
for m in (allsplits if ALL else problems): print('  ' + m)
