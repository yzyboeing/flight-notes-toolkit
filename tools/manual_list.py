#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""manual_list.py —— 从手册资料包登记表重新生成《手册版本清单》（SD-113；2026-10-27 起改为脚本生成）
用法：python3 manual_list.py [--manuals ~/flight-repos/manuals] [--kb ~/Desktop/B737飞行理论知识库] --out 清单.md [--out 另一份.md]
  - 读 manuals/MANIFEST-*.csv（知识库导出时生成，路径以它为准）；
  - 波音手册（AFM / FCOM / FCTM / QRH）从逐页文本里抽「最新页日期」（页脚日期的最大值，供核对时确认版本；以 PDF 修订页为准）；
  - 登记表里路径已不存在的条目标「已取代 / 不在库」，不删行；
  - 换版记录一节从旧清单原样保留。"""
import csv, os, re, sys, io, glob, datetime, collections

def arg(k, d=None, many=False):
    v = [sys.argv[i + 1] for i, a in enumerate(sys.argv) if a == k and i + 1 < len(sys.argv)]
    return v if many else (v[0] if v else d)

M = os.path.expanduser(arg('--manuals', '~/flight-repos/manuals'))
K = os.path.expanduser(arg('--kb', '~/Desktop/B737飞行理论知识库'))
OUTS = [os.path.expanduser(x) for x in arg('--out', many=True)]
if not OUTS: sys.exit('用法：manual_list.py --out <清单.md>')
TODAY = datetime.date.today()
MON = 'January February March April May June July August September October November December'.split()

def latest_date(md):
    try: t = io.open(md, encoding='utf-8', errors='ignore').read()
    except OSError: return ''
    fm = re.search(r'^revision_date:\s*"?([^"\n]+?)"?\s*$', t[:3000], re.M)   # 逐页文本文件头有修订日期就用它
    if fm and fm.group(1) != 'null':
        m = re.match(r'(%s)\s+(\d{1,2}),\s+(20\d\d)$' % '|'.join(MON), fm.group(1))
        return datetime.date(int(m.group(3)), MON.index(m.group(1)) + 1, int(m.group(2))).isoformat() if m else fm.group(1)
    c, cs = collections.Counter(), collections.Counter()
    for y, mo, d in re.findall(r'(20\d\d)[/.](\d{1,2})[/.](\d{1,2})(?!\d)', t):   # 斜杠 / 点分日期多是印章、批注，要 ≥ 10 处才算页脚
        try: cs[datetime.date(int(y), int(mo), int(d))] += 1
        except ValueError: pass
    for mo, d, y in re.findall(r'(%s)\s+(\d{1,2}),\s+(20\d\d)' % '|'.join(MON), t):
        try: c[datetime.date(int(y), MON.index(mo) + 1, int(d))] += 1
        except ValueError: pass
    for y, mo, d in re.findall(r'(20\d\d)\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', t):
        try: c[datetime.date(int(y), int(mo), int(d))] += 1
        except ValueError: pass
    ok = [k for k, v in c.items() if k <= TODAY and v >= 3] + [k for k, v in cs.items() if k <= TODAY and v >= 10]   # 至少 3 处，避开正文里偶然提到的日期
    return max(ok).isoformat() if ok else ''

rows = []
for f, cat in (('MANIFEST-公司手册.csv', '手册'), ('MANIFEST-扩充资料.csv', '扩充资料')):
    p = os.path.join(M, f)
    if not os.path.exists(p): continue
    for r in csv.DictReader(io.open(p, encoding='utf-8-sig')):
        sp = r.get('source_path', '')
        mdp = r.get('markdown_path', '')
        exists = sp.startswith('资料/') and os.path.exists(os.path.join(K, sp))
        pdf = os.path.join(M, r.get('pdf_path', '')) if r.get('pdf_path') else ''
        size = round(os.path.getsize(pdf) / 1048576) if pdf and os.path.exists(pdf) else ''
        rev = r.get('revision', '')
        rev = '未知' if rev in ('', 'UNKNOWN') else rev
        boeing = r.get('document_type', '') in ('AFM', 'FCOM', 'FCTM', 'QRH')
        rows.append(dict(cat=cat, title=r['title'], ac=re.sub(r'[\[\]"]', '', r.get('aircraft_series', '')).replace(',', '、') or '通用', lang=r.get('language', ''), rev=rev,
                         last=latest_date(os.path.join(M, mdp)) if boeing and mdp else '', size=size,
                         path=('`%s`' % sp) if exists else ('已取代 / 不在库（%s）' % sp.strip('（）') if not sp.startswith('资料/') else '不在库：`%s`' % sp),
                         md='`%s`' % mdp if mdp else ''))

old_log = ''
for o in OUTS:
    if os.path.exists(o):
        t = io.open(o, encoding='utf-8').read()
        if '## 换版记录' in t: old_log = '## 换版记录' + t.split('## 换版记录', 1)[1]; break
if not old_log: old_log = '## 换版记录（新的在前）\n\n（暂无。）\n'

n1 = sum(1 for r in rows if r['cat'] == '手册'); n2 = len(rows) - n1
L = ['# 手册版本清单（给噜噜）', '',
     '- **清单更新日期：%s**（每次换版或新增资料后，由 Claude 按 SD-113 用 `pub/tools/manual_list.py` 重新生成）' % TODAY.isoformat(),
     '- 用法：开新对话、基线换新、引用手册前，先看上面的「清单更新日期」。日期和你上次记下的不同，就去读同目录的《手册换版通知_…》，提醒用户哪本换了版，以后按新版引用，并记下新日期。',
     '- 只读下表列出的现行路径（总库根目录 `~/Desktop/B737飞行理论知识库/`）；**历史版本目录里的旧版不作依据**。',
     '- 单个文件超过 50MB 读不了时，改读「逐页文本」列的文件（在 `~/flight-repos/manuals/` 下，带 PDF 页、印刷页、章节），或按章节分段读。数值仍以 PDF 原页为准。',
     '- 共 %d 份：手册 %d 份，扩充资料 %d 份。修订号「未知」表示登记表里没有记录，不代表没有修订号；波音手册另列「最新页日期」（逐页文本里页脚日期的最大值），引用时写这个日期，有疑问以 PDF 修订页为准。' % (len(rows), n1, n2),
     '', '| 类别 | 手册 | 机型 | 语言 | 修订号 | 最新页日期 | 大小(MB) | 总库路径 | 逐页文本 |', '|---|---|---|---|---|---|---|---|---|']
for r in rows:
    L.append('| %(cat)s | %(title)s | %(ac)s | %(lang)s | %(rev)s | %(last)s | %(size)s | %(path)s | %(md)s |' % r)
L += ['', old_log.rstrip('\n'), '']
for o in OUTS:
    io.open(o, 'w', encoding='utf-8').write('\n'.join(L))
print('手册版本清单：%d 份（手册 %d、扩充资料 %d），波音手册最新页日期 %d 本 → %s' % (len(rows), n1, n2, sum(1 for r in rows if r['last']), '、'.join(OUTS)))
