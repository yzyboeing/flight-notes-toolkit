#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_coverage.py —— 回原手册查漏（SD-209 K8，2026-10-19；按需跑，不进 sync）

用法：python3 check_coverage.py [--repo ~/flight-repos/gh-private] [--only 1.2,1.4] [--out build/手册覆盖报告.md]

做法（两步，和 1018 审查报告二同口径，工具化）：
  ① 节 front matter 的 `fcom:` 字段（如 fcom: ["2.10", "2.20"]，FCOM 章.节）→ 取两型 FCOM 中文版该章节全部页块，
     列出两本笔记（含第四章）都没有的「数值 + 单位」和「灯 / 电门 / 方式名（大写英文）」所在句子；
     没写 fcom 字段的节跳过（提示补字段）。
  ② 每个条目按原句片段在整个中文手册库里定位原文段落，段内带数值 / 限制词而两本没有的句子列为候选。
单位先换算统一（英尺 → ft、节 → kt、千克 → kg …），否则大量假阳性；结果仍需人工按上下文复核（改写过的句子机器认不出）。
报告：build/手册覆盖报告.md（按节 → 条目 → 候选句，附手册名与 PDF 页）。"""
import re, io, os, glob, html, bisect, collections, sys
REPO = os.path.abspath(os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else '~/flight-repos/gh-private'))
ONLY = set(sys.argv[sys.argv.index('--only') + 1].split(',')) if '--only' in sys.argv else None
OUT = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else os.path.join(REPO, 'build', '手册覆盖报告.md')
M = os.path.expanduser('~/flight-repos/manuals/Markdown/')
strip = lambda s: html.unescape(re.sub(r'<[^>]+>', '', re.sub(r'<!--.*?-->', '', s, flags=re.S)))
N = lambda s: re.sub(r'[^\w]', '', s)
U = [('英尺', 'ft'), ('节', 'kt'), ('海里', 'nm'), ('分钟', 'min'), ('秒钟', 's'), ('秒', 's'), ('摄氏度', '℃'), ('华氏度', '℉'), ('度', '°'), ('小时', 'h'),
     ('磅/平方英寸', 'psi'), ('PSI', 'psi'), ('千克', 'kg'), ('公斤', 'kg'), ('伏', 'V'), ('米', 'm')]
def nu(x):
    x = re.sub(r'[\s,，]', '', x)
    for a, b in U: x = x.replace(a, b)
    return x
NUM = re.compile(r'\d+(?:[.,]\d+)?\s*(?:英尺|节|海里|分钟|秒钟|秒|摄氏度|华氏度|度|小时|psi|PSI|磅/平方英寸|%|伏|安培|米|千克|公斤|磅|ft|kt|nm|min|℃)')
UPW = re.compile(r'\b[A-Z][A-Z0-9]{2,}(?:\s[A-Z][A-Z0-9]{2,}){0,3}\b')
SKIPW = {'FCOM', 'QRH', 'FCTM', 'NNC', 'SOP', 'FMC', 'CDU', 'MCP', 'EFIS', 'PFD', 'TEXT', 'NOTE', 'ON', 'OFF', 'AUTO', 'NORMAL', 'TEST', 'PDF', 'OCR'}
_JUNK_EXTRA = [l.strip() for l in io.open(os.path.expanduser('~/.config/flight-notes/coverage_junk.txt'), encoding='utf-8') if l.strip()] if os.path.exists(os.path.expanduser('~/.config/flight-notes/coverage_junk.txt')) else []   # 本机私有：公司名、文件编号前缀等不进公开库
JUNK = re.compile('|'.join([r'章节：', 'OCR', r'页面引用键', r'File Highlight', r'\?\?\?', '目录', '修订', '有效页', '版权所有', '印刷页', 'Copyright', '参见', '见第', r'^\s*\d+\s*$'] + [re.escape(x) for x in _JUNK_EXTRA]))
SIG = re.compile(r'\d|必须|不得|禁止|不能|只有|除非|不可|切勿|警告|注意|限制|最大|最小|不超过|超过|低于|高于|以上|以下|自动|失效|故障|不工作')

# 两本笔记全文（含第四章；别处收了就不算漏）
notes = ''.join(strip(io.open(f, encoding='utf-8').read()) for f in glob.glob(os.path.join(REPO, 'notes_src', '*', '*.md')))
notes += strip(io.open(os.path.join(REPO, '速查', '速查源.md'), encoding='utf-8').read())
NZ = re.sub(r'[\s,，]', '', notes); NZU = NZ.upper()
nn = N(notes); NS = set(nn[i:i + 4] for i in range(len(nn) - 3))
def cov(s):
    s = N(s); sh = [s[i:i + 4] for i in range(len(s) - 3)]
    return sum(1 for x in sh if x in NS) / max(1, len(sh))
def novel_tokens(s):
    toks = NUM.findall(s) + [x for x in UPW.findall(s) if x not in SKIPW and not re.match(r'^[A-Z]\d', x)]
    return [x for x in toks if nu(x) not in NZ and re.sub(r'\s', '', x).upper() not in NZU]

# ① 按 fcom 字段扫两型 FCOM
FCOM = {}
for tag, pat in (('737-8', 'B737-8 飞行机组操作手册 FCOM（中文）'), ('737-NG', 'B737-NG 飞行机组操作手册 FCOM（中文）')):
    fs = [x for x in glob.glob(M + '公司手册/*.md') if pat in x]
    if not fs: continue
    t = io.open(fs[0], encoding='utf-8').read()
    for m in re.finditer(r'### PDF 第 (\d+) 页 · 印刷页 ([^\n]*)\n(?:[^\n]*\n){0,5}?- 章节：([^·\n]*)· 节：([^·\n]*)· 标题：([^\n]*)\n(?:[^\n]*\n){0,5}?```text\n(.*?)```', t, re.S):
        pg, pp, ch, sec, tit, body = m.groups()
        FCOM.setdefault(tag, []).append((pg, pp.strip(), sec.strip(), tit.strip(), body))
def fcom_pages(tag, key):   # key 形如 2.10 → 印刷页以 2.10. 开头
    return [x for x in FCOM.get(tag, []) if x[1].startswith(key + '.') or x[1] == key]

# ② 手册库定位用的大文本
files = [f for f in glob.glob(M + '*/*.md') if '英文' not in os.path.basename(f)]
LINES, BIG, OFF, pos = [], [], [], 0
for fi, f in enumerate(files):
    for li, l in enumerate(io.open(f, encoding='utf-8', errors='ignore').read().split('\n')):
        n = N(l)
        if not n: continue
        LINES.append((fi, li, l)); OFF.append(pos); BIG.append(n); pos += len(n) + 1
BIGS = '\n'.join(BIG)
def where(p): return bisect.bisect_right(OFF, p) - 1

rep = ['# 手册覆盖报告（check_coverage.py，%s）\n' % os.path.basename(REPO), '> 候选 = 手册句子里有两本笔记都没有的数值 / 名称；须人工按上下文复核，改写过的句子机器认不出。\n']
tot = collections.Counter()
for f in sorted([x for x in glob.glob(os.path.join(REPO, 'notes_src', '[1235]*', '*.md')) if re.match(r'\d+\.\d+ ', os.path.basename(x))], key=lambda x: [int(v) for v in os.path.basename(x).split(' ')[0].split('.')]):
    b = os.path.basename(f); sec = b.split(' ')[0]
    if ONLY and sec not in ONLY: continue
    t = io.open(f, encoding='utf-8').read()
    fm = re.match(r'\A---\n(.*?)\n---\n', t, re.S)
    fc = re.search(r'^fcom:\s*\[(.*?)\]', fm.group(1), re.M) if fm else None
    keys = [x.strip().strip('"\'') for x in fc.group(1).split(',') if x.strip()] if fc else []
    rep.append('\n## %s\n' % b[:-3])
    if not keys: rep.append('- （front matter 无 fcom 字段，跳过 FCOM 章节扫描；补写如 `fcom: ["2.10"]`）')
    for key in keys:
        for tag in ('737-NG', '737-8'):
            seen = set()
            for pg, pp, s2, tit, body in fcom_pages(tag, key):
                if '目录' in tit or '目录' in s2: continue
                txt = re.sub(r'\n(?=[^\n•\d])', '', body)
                for s in re.split(r'(?<=[。；])|\n', txt):
                    s = s.strip()
                    if not (10 <= len(s) <= 220) or JUNK.search(s) or len(re.findall(r'[\u4e00-\u9fff]', s)) < 6: continue   # 面板标签、图上英文不算
                    nov = [x for x in novel_tokens(s) if not re.match(r'^(D\d|[A-Z]{1,2}\d|[A-Z]{3}-\d)', x) and x not in _JUNK_EXTRA]
                    if not nov or N(s)[:30] in seen: continue
                    seen.add(N(s)[:30]); tot['fcom'] += 1
                    rep.append('- [FCOM %s %s · PDF %s · %s] 缺 %s：%s' % (tag, pp, pg, tit[:16], ' / '.join(nov[:3]), s[:140]))
    # ② 逐条目定位原文
    for mm in re.finditer(r'(?ms)^### (\d+)\. (.+?)\n(.*?)(?=^### |\Z)', t):
        body = strip(mm.group(3)); nb = N(body)
        if len(nb) < 20: continue
        L, step = 7, max(3, len(nb) // 60)
        probes = list(dict.fromkeys(nb[i:i + L] for i in range(0, len(nb) - L, step)))
        hits = collections.defaultdict(set)
        for pi, p in enumerate(probes):
            s0 = c0 = 0
            while c0 < 4:
                j = BIGS.find(p, s0)
                if j < 0: break
                w = where(j); hits[(LINES[w][0], LINES[w][1] // 80)].add((pi, w)); s0 = j + 1; c0 += 1
        if not hits: continue
        best = max(hits, key=lambda k: len({x[0] for x in hits[k]}) + 0.5 * len({x[0] for x in hits.get((k[0], k[1] + 1), ())}))
        hs = hits[best] | hits.get((best[0], best[1] + 1), set()) | hits.get((best[0], best[1] - 1), set())
        if len({x[0] for x in hs}) < max(2, len(probes) * 0.15): continue
        ws = sorted(x[1] for x in hs); med = ws[len(ws) // 2]; ws = [w for w in ws if abs(w - med) <= 120]
        a0, b0 = max(0, ws[0] - 12), min(len(LINES) - 1, ws[-1] + 12); fi = LINES[med][0]
        txt = ''.join(LINES[i][2].strip() for i in range(a0, b0 + 1) if LINES[i][0] == fi and not LINES[i][2].startswith('#'))
        cands = []
        for s in re.split(r'(?<=[。；！？])', txt):
            s = s.strip()
            if not (12 <= len(N(s)) <= 220) or JUNK.search(s) or not SIG.search(s) or cov(s) >= 0.35: continue
            nov = novel_tokens(s)
            if nov or cov(s) < 0.15: cands.append((nov, s))
        if cands:
            tot['item'] += len(cands)
            rep.append('- 第 %s 条「%s」← %s：' % (mm.group(1), mm.group(2)[:24], os.path.basename(files[fi]).split('--')[0][:24]))
            for nov, s in cands[:6]: rep.append('  - 缺 %s：%s' % (' / '.join(nov[:3]) or '（整句）', s[:160]))
os.makedirs(os.path.dirname(OUT), exist_ok=True)
io.open(OUT, 'w', encoding='utf-8').write('\n'.join(rep) + '\n')
print('手册覆盖：FCOM 章节候选 %d 句、条目定位候选 %d 句 → %s' % (tot['fcom'], tot['item'], OUT))
