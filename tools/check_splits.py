#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_splits.py —— 成品 PDF 断页检查（2026-09-29，SD-71）
用法：python3 check_splits.py <全书.pdf> [--all]
逐页用 PyMuPDF 识别表格，找出「页底表格接到下一页页顶继续」的断表，报告：
  · 孤行：断开后上半截或下半截只有 1～2 行；
  · 本可整页：两截高度之和不超过一页可用高度（说明这张表本来放得进一页）。
  · 标题孤立：页底最后一段是条目标题（数字编号或 A-1 这类），正文到了下一页。
默认只列有问题的；--all 列出全部断表。只读，不改文件。"""
import sys, re
import pymupdf

PDF = sys.argv[1]; ALL = '--all' in sys.argv
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
    foot = [x for x in blocks if re.match(r'\s*第\s*\d+\s*页\s*$', x[4])]
    fy = min((x[1] for x in foot), default=H)
    tops = sorted(blocks, key=lambda x: x[1])
    hy = tops[0][3] if tops and tops[0][3] < 0.08 * H and tops[0] not in foot else 0
    body = [x for x in blocks if x[1] >= hy and x[3] <= fy and x not in foot and (hy == 0 or x is not tops[0])]
    top = min((b[1] for b in body), default=H); bot = max((b[3] for b in body), default=0)
    info.append({'tabs': tabs, 'top': top, 'bot': bot, 'H': H, 'blocks': body})
AREA = max(x['bot'] - x['top'] for x in info if x['blocks']) if info else 0
problems, allsplits = [], []
head_re = re.compile(r'^\s*(\d{1,3}\.\s+\S|[A-Z]-\d+\s)')
for i in range(len(info) - 1):
    a, b = info[i], info[i + 1]
    # 标题孤立：本页最后一个文字块是条目标题
    if a['blocks']:
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
    ra, rb = ta.row_count, tb.row_count
    ha, hb = ta.bbox[3] - ta.bbox[1], tb.bbox[3] - tb.bbox[1]
    first = (ta.extract()[0] if ra else [''])
    name = ' | '.join(str(x or '').replace('\n', '') for x in first)[:40]
    msg = '第 %d→%d 页：上截 %d 行 / 下截 %d 行（表头：%s）' % (i + 1, i + 2, ra, rb, name)
    allsplits.append(msg)
    why = []
    if ra <= 2 or rb <= 2: why.append('孤行')
    if ha + hb <= AREA * 0.97: why.append('本可整页')
    if why: problems.append(msg + '　→ ' + '、'.join(why))
print('共 %d 页；断表 %d 处；问题 %d 处' % (len(d), len(allsplits), len(problems)))
for m in (allsplits if ALL else problems): print('  ' + m)
