#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_consistency.py —— 全书数值一致性扫描（SD-241 B3，2026-10-27；按需跑，不进 sync）
用法：python3 check_consistency.py [--repo ~/flight-repos/gh-private] [--out build/数值一致性报告.md] [--min 2]
做法：把每条正文拆成句 / 表格单元，取「数值 + 单位」和它前面紧挨着的关键词（去掉标点、HTML、机型标注后最后 4～10 个字）作为键；
同一个键在 ≥ 2 个不同条目里出现不同数值，就列为候选。两型不同（同句带【737-NG】/【737-8】或在两型对比表的不同列）不算冲突，单独标出。
结果须人工判断：同一说法换了语境（如「襟翼 15」既是着陆襟翼又是复飞襟翼）不是错误。"""
import re, io, os, glob, html, sys, collections

def arg(k, d=None): return sys.argv[sys.argv.index(k) + 1] if k in sys.argv else d
REPO = os.path.abspath(os.path.expanduser(arg('--repo', '~/flight-repos/gh-private')))
OUT = arg('--out', os.path.join(REPO, 'build', '数值一致性报告.md'))
MIN = int(arg('--min', '2'))
UNIT = r'(?:ft|kt|nm|NM|fpm|psi|hPa|kg|lb|℃|°C|%|°|min|s|h|m|km|g|L|Mach|M)'
NUM = re.compile(r'(?<![\w.])(\d+(?:\.\d+)?)\s*(' + UNIT + r')(?![A-Za-z])')
TYPE = re.compile(r'【737-(?:NG|8)[^】]*】|737-NG|737-8')
skey = lambda s: [int(v) for v in s.split('.')]

def cells(body):
    """表格按单元格、正文按句拆；两型对比表的列标记带出来。"""
    out = []
    for tb in re.findall(r'(?s)<table.*?</table>', body):
        hdr = re.search(r'(?s)<tr class="hdr">(.*?)</tr>', tb)
        cols = [re.sub(r'<[^>]+>', '', h) for h in re.findall(r'(?s)<th[^>]*>(.*?)</th>', hdr.group(1))] if hdr else []
        for tr in re.findall(r'(?s)<tr(?! class="hdr").*?</tr>', tb):
            tds = re.findall(r'(?s)<td[^>]*>(.*?)</td>', tr)
            first = html.unescape(re.sub(r'<[^>]+>', '', tds[0])).strip() if tds else ''
            for i, td in enumerate(tds):
                col = cols[i] if i < len(cols) else ''
                out.append((first if i else '', col, td))
    rest = re.sub(r'(?s)<table.*?</table>', '', body)
    for para in re.split(r'[。；\n]', rest):
        out.append(('', '', para))
    return out

def plain(t): return re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', '', t))).strip()

IDX = collections.defaultdict(list)   # key → [(value, addr, type-tag, snippet)]
for f in glob.glob(os.path.join(REPO, 'notes_src', '[1-6] *', '[0-9]*.md')):
    b = os.path.basename(f); m = re.match(r'(\d+\.\d+) ', b)
    if not m: continue
    t = io.open(f, encoding='utf-8').read()
    for im in re.finditer(r'(?ms)^### (\d+)\. (.*?)\n(.*?)(?=^### |\Z)', t):
        addr = '%s 第 %s 条' % (m.group(1), im.group(1))
        for rowlab, col, c in cells(im.group(3)):
            s = plain(c)
            for nm in NUM.finditer(s):
                pre = re.sub(r'[\W\d_]+', '', TYPE.sub('', s[max(0, nm.start() - 16):nm.start()]))[-8:]
                if len(pre) < 3: pre = re.sub(r'[\W\d_]+', '', TYPE.sub('', rowlab))[-8:] + pre
                if len(pre) < 3: continue
                tag = (TYPE.search(s[max(0, nm.start() - 40):nm.end() + 10]) or TYPE.search(col) or [''])[0]
                IDX[(pre, nm.group(2))].append((nm.group(1), addr, tag, s[max(0, nm.start() - 30):nm.end() + 12]))

rep = ['# 数值一致性报告（check_consistency.py）', '',
       '> 同一说法（数值前的关键词 + 单位）在不同条目里出现不同数值。**须人工判断**：语境不同、两型不同、举例都不是错误；真不一致的按手册改并联动。', '']
n = 0; TWO = []
for (pre, unit), occ in sorted(IDX.items(), key=lambda kv: -len({o[1] for o in kv[1]})):
    vals = {o[0] for o in occ}; addrs = {o[1] for o in occ}
    if len(vals) < 2 or len(addrs) < MIN: continue
    by = collections.defaultdict(set)
    for v, a, tag, sn in occ: by[v].add(a)
    tags = collections.defaultdict(set)
    for v, a, tag, sn in occ: tags[v].add(tag)
    if all(len(tg) == 1 and '' not in tg for tg in tags.values()) and len({next(iter(tg)) for tg in tags.values()}) == len(tags):
        TWO.append((pre, unit)); continue   # 两型各一个值、每处都标了机型：两型差异，不是冲突
    n += 1
    rep.append('## %d. 「%s」+ %s：%s' % (n, pre, unit, ' / '.join('%s%s（%d 处）' % (v, unit, len(by[v])) for v in sorted(by, key=lambda x: float(x)))))
    for v, a, tag, sn in sorted(occ, key=lambda o: (float(o[0]), skey(o[1].split(' ')[0]))):
        rep.append('- %s%s　%s%s　…%s…' % (v, unit, a, '（%s）' % tag if tag else '', sn))
    rep.append('')
rep.append('> 另有 %d 组是两型差异（每个数值都标了机型），不列出。' % len(TWO))
io.open(OUT, 'w', encoding='utf-8').write('\n'.join(rep) + '\n')
print('数值一致性：候选 %d 组（两型差异 %d 组不列）→ %s' % (n, len(TWO), OUT))
