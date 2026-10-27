#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""move_items.py —— 节与条目的移动、拆节、并节、调序与改号（SD-209 批 3：S1 调序 / S2 拆杂烩节 / S3 技术提示归系统 / S4 拆大节）。
按方案 JSON 重写 notes_src 里的节文件，条目按方案顺序重新编号，并一次性改全部引用（同一遍替换，链式改号不会二次改）：
  · 双链 [[x.y 节名|x.y 第 N 条]]、[[x.y 节名]]、正文「x.y 第 N 条」、图文件 figures/x.y-N_ 与图注「图 x.y-N」；
  · front matter 的 related；章目录 _N … 目录.md 与 MANIFEST.txt 按新节重排；
  · JSON（两本主题索引、裁定台账、白名单等）里的「x.y N」与「x.y 第 N 条」；可选：学习系统题库、规则正本。
方案格式：
  {"sections": [{"id": "1.2", "dir": "1 系统描述", "title": "引气、空调与增压", "fm": "1.2",
                 "items": ["1.2 1", "1.2 5", "%%SUB%% 技术提示", "5.9 1", …],        # 未写的老条目会报错，不会悄悄丢
                 "retitle": {"5.9 8": "人工增压常见问题与操作提示"},                    # 改标题（星级保留）
                 "append": {"5.10 1": "注：依据——…"}}],                              # 条目末尾追加一段（并入节的节首说明）
   "rename": {"3.7": "3.9"},          # 整节改号（条目不变）
   "remove": ["5.9", "5.10"],         # 条目已全部并走的旧节：文件移到 --trash
   "merge": {"1.14 2": "1.14 1"}}      # 并掉的条目：内容已手工并进目标条目，本条不再写出；指向它的引用改指目标条目的新号
用法：python3 move_items.py --repo <gh-private> --plan plan.json [--study <study 库>] [--prompt <pub/prompt>] [--trash <目录>] [--map-out 对照表.json --from 旧版名 --to 新版名] [--apply]
  --map-out：把改了号的条目写成对照表 JSON（{from, to, 说明, map}，学习系统 remap 与噜噜用；同名 .md 是人读版）。
不加 --apply 只出对照表与检查结果。"""
import re, io, os, sys, json, glob, shutil, datetime, collections

def arg(k, d=None):
    return sys.argv[sys.argv.index(k) + 1] if k in sys.argv else d
REPO = os.path.abspath(os.path.expanduser(arg('--repo', '.')))
PLAN = json.load(io.open(arg('--plan'), encoding='utf-8'))
STUDY = arg('--study'); PROMPT = arg('--prompt')
TRASH = os.path.expanduser(arg('--trash', '~/flight-repos/_to_delete/move_items_%s' % datetime.datetime.now().strftime('%Y%m%d_%H%M%S')))
APPLY = '--apply' in sys.argv
SRC = os.path.join(REPO, 'notes_src')
TODAY = datetime.date.today().isoformat()
skey = lambda s: [int(v) for v in s.split('.')]

# ---------- 1 读现有节 ----------
SUBS = {}   # 节 → {在其后出现分组标记的条目号: 标记行}
SEC = {}   # id → dict(path, dir, fm(str), fmd(dict-ish lines), title, pre, items[(n, title, star, body)])
for f in glob.glob(os.path.join(SRC, '[1-6]*', '[0-9]*.md')):
    b = os.path.basename(f); m = re.match(r'(\d+\.\d+) ', b)
    if not m: continue
    t = io.open(f, encoding='utf-8').read()
    fm = re.match(r'(?s)\A---\n(.*?)\n---\n', t)
    rest = t[fm.end():] if fm else t
    tm = re.search(r'(?m)^# (\d+\.\d+)[　 ](.+)$', rest)
    first = re.search(r'(?m)^### ', rest)
    pre = rest[tm.end():first.start() if first else len(rest)].strip('\n') if tm else ''
    items = []
    for im in re.finditer(r'(?ms)^### (\d+)\. (.*?)(?:\s*\{(★+)\})?[ \t]*\n(.*?)(?=^### |\Z)', rest):
        body = im.group(4)
        if re.search(r'(?m)^%%SUB%%', body):   # 节内分组标记是结构，不是条目内容：记下位置、从正文去掉（2026-10-19 批 5 重复标记教训）
            SUBS.setdefault(m.group(1), {})[int(im.group(1))] = re.search(r'(?m)^(%%SUB%%.*)$', body).group(1)
            body = re.sub(r'(?m)^%%SUB%%.*\n?', '', body)
        items.append((int(im.group(1)), im.group(2).strip(), im.group(3) or '', body.strip('\n')))
    SEC[m.group(1)] = dict(path=f, dir=os.path.basename(os.path.dirname(f)), fm=fm.group(1) if fm else '', title=tm.group(2).strip() if tm else b[:-3].split(' ', 1)[1], pre=pre, items=items)

# ---------- 2 方案 → 映射 ----------
plan_secs = list(PLAN.get('sections', []))
for old, new in PLAN.get('rename', {}).items():
    s = SEC[old]
    its = []
    for it in s['items']:
        its.append('%s %d' % (old, it[0]))
        if it[0] in SUBS.get(old, {}): its.append(SUBS[old][it[0]])
    plan_secs.append({'id': new, 'dir': s['dir'], 'title': s['title'], 'fm': old, 'items': its})
IMAP, SMAP, NEWT = {}, collections.defaultdict(list), {}
errs = []; used = set()
MERGE = {tuple([k.split(' ')[0], int(k.split(' ')[1])]): tuple([v.split(' ')[0], int(v.split(' ')[1])]) for k, v in PLAN.get('merge', {}).items()}
for ps in plan_secs:
    NEWT[ps['id']] = ps['title']; n = 0
    for k in ps['items']:
        if k.startswith('%%'): continue
        o, on = k.split(' '); on = int(on); n += 1
        if (o, on) in used: errs.append('条目重复列入：%s' % k)
        used.add((o, on)); IMAP[(o, on)] = (ps['id'], n)
        if ps['id'] not in SMAP[o]: SMAP[o].append(ps['id'])
for ps in plan_secs:   # 节级引用的主去向：fm 指向的节排第一
    if ps['id'] in SMAP.get(ps['fm'], []):
        SMAP[ps['fm']].remove(ps['id']); SMAP[ps['fm']].insert(0, ps['id'])
src_secs = {k.split(' ')[0] for ps in plan_secs for k in ps['items'] if not k.startswith('%%')}
for s in src_secs:
    for it in SEC[s]['items']:
        if (s, it[0]) in MERGE: continue
        if (s, it[0]) not in IMAP: errs.append('漏列：%s %d %s（会丢条目）' % (s, it[0], it[1]))
removed = set(PLAN.get('remove', []))
for s in src_secs:
    if s not in {ps['id'] for ps in plan_secs} and s not in removed and s not in PLAN.get('rename', {}):
        errs.append('节 %s 的条目全部移走，但没写进 remove' % s)
for s, v in SEC.items():   # 未动的节：恒等映射（标题可能因改号变化）
    if s in src_secs: continue
    for it in v['items']: IMAP[(s, it[0])] = (s, it[0])
    SMAP[s] = [s]; NEWT.setdefault(s, v['title'])
for k, v in MERGE.items():   # 并掉的条目：引用改指目标条目
    if v not in IMAP: errs.append('merge 目标 %s %d 不在新结构里' % v); continue
    if k in used: errs.append('merge 的条目 %s %d 又被列进了 items' % k)
    IMAP[k] = IMAP[v]
if errs:
    print('\n'.join('✗ ' + e for e in errs)); sys.exit(1)

# 图：移到别节的条目里的图改号（按新节内出现顺序接在已有图号后）
FMAP = {}
def figs_in(body): return re.findall(r'figures/(\d+\.\d+)-(\d+)_', body)
for ps in plan_secs:
    have = [int(n) for k in ps['items'] if not k.startswith('%%') for s, n in figs_in(dict((i[0], i[3]) for i in SEC[k.split(' ')[0]]['items'])[int(k.split(' ')[1])]) if s == ps['id'] and ps['fm'] == ps['id']]
    nxt = max(have or [0])
    for k in ps['items']:
        if k.startswith('%%'): continue
        o, on = k.split(' ')
        body = dict((i[0], i[3]) for i in SEC[o]['items'])[int(on)]
        for s, n in figs_in(body):
            if (s, n) in FMAP: continue
            if s == ps['id'] and ps['fm'] == ps['id']: FMAP[(s, n)] = (s, n); continue
            nxt += 1; FMAP[(s, n)] = (ps['id'], str(nxt))
FMAP = {k: v for k, v in FMAP.items() if k != v}

# ---------- 3 一遍替换 ----------
UNRES = collections.Counter()
def newname(s): return '%s %s' % (s, NEWT.get(s, SEC.get(s, {}).get('title', '?')))
RX = re.compile(r'\[\[(\d+\.\d+) ([^\]|]*?)(?:\|(\d+\.\d+) 第 (\d+) 条)?(\|[^\]]*)?\]\]'      # 1-4 双链（条目 / 节），5 其他标签
                r'|(?<![\d.])(\d+\.\d+) 第 (\d+) 条'                                              # 6-7 正文条目地址
                r'|figures/(\d+\.\d+)-(\d+)_'                                                    # 8-9 图文件
                r'|图 (\d+\.\d+)-(\d+)(?!\d)')                                                   # 10-11 图号
def rep(m):
    if m.group(1):
        s, lab_s, lab_n, other = m.group(1), m.group(3), m.group(4), m.group(5)
        if lab_s:
            k = (lab_s, int(lab_n))
            if k not in IMAP: UNRES['链接 %s 第 %s 条' % k] += 1; return m.group(0)
            ns, nn = IMAP[k]; return '[[%s|%s 第 %d 条]]' % (newname(ns), ns, nn)
        if s not in SMAP or not SMAP[s]: UNRES['节链接 %s' % s] += 1; return m.group(0)
        return '[[%s%s]]' % (newname(SMAP[s][0]), other or '')
    if m.group(6):
        k = (m.group(6), int(m.group(7)))
        if k[0] not in SEC: return m.group(0)
        if k not in IMAP: UNRES['%s 第 %d 条' % k] += 1; return m.group(0)
        return '%s 第 %d 条' % IMAP[k]
    if m.group(8):
        k = (m.group(8), m.group(9)); return 'figures/%s-%s_' % FMAP.get(k, k)
    k = (m.group(10), m.group(11)); return '图 %s-%s' % FMAP.get(k, k)
RXJ = re.compile(r'(?<![\d.])(\d\.\d{1,2}) (\d{1,2})(?![\d.])')   # JSON / 题库的「x.y N」
def repj(m):
    k = (m.group(1), int(m.group(2)))
    if k[0] not in SEC: return m.group(0)
    if k not in IMAP: UNRES['%s %d' % k] += 1; return m.group(0)
    return '%s %d' % IMAP[k]
def conv(t, json_style=False):
    t = RX.sub(rep, t)
    return RXJ.sub(repj, t) if json_style else t
def fix_related(fm, self_id):
    def r(m):
        out = []
        for x in re.findall(r'"([^"]*)"', m.group(1)):
            for y in (SMAP.get(x) or [x]):
                if y != self_id and y not in out: out.append(y)
        return 'related: [%s]' % ', '.join('"%s"' % x for x in out)
    return re.sub(r'(?m)^related: \[(.*)\]$', r, fm)

writes = {}   # path → text
moves = []    # (old, new)
for ps in plan_secs:
    base = SEC[ps['fm']]; fm = base['fm']
    srcs = []; [srcs.append(k.split(' ')[0]) for k in ps['items'] if not k.startswith('%%') and k.split(' ')[0] not in srcs]
    fm = re.sub(r'(?m)^id: .*$', 'id: "%s"' % ps['id'], fm)
    fm = re.sub(r'(?m)^title_cn: .*$', 'title_cn: "%s"' % ps['title'], fm)
    def union(key):
        vals = []
        for s in [ps['fm']] + srcs:
            mm = re.search(r'(?m)^%s: \[(.*)\]$' % key, SEC[s]['fm'])
            for x in re.findall(r'"([^"]*)"', mm.group(1)) if mm else []:
                if x not in vals: vals.append(x)
        return vals
    for key in ('aircraft', 'source'):
        fm = re.sub(r'(?m)^%s: \[.*\]$' % key, '%s: [%s]' % (key, ', '.join('"%s"' % x for x in union(key))), fm)
    rel = union('related') + [s for s in srcs if s != ps['id']] + [ps['fm']]
    fm = re.sub(r'(?m)^related: \[.*\]$', 'related: [%s]' % ', '.join('"%s"' % x for x in rel), fm)
    if len(srcs) > 1 or srcs != [ps['id']] or ps.get('items') != ['%s %d' % (ps['fm'], i[0]) for i in base['items']]:
        fm = re.sub(r'(?m)^updated: .*$', 'updated: "%s"' % TODAY, fm)
    fm = fix_related(fm, ps['id'])
    if ps.get('related_add'):   # 新节号，不再映射
        fm = re.sub(r'(?m)^related: \[(.*)\]$', lambda m: 'related: [%s]' % ', '.join('"%s"' % x for x in dict.fromkeys(re.findall(r'"([^"]*)"', m.group(1)) + [x for x in ps['related_add'] if x != ps['id']])), fm)
    parts = ['---\n%s\n---\n\n# %s　%s' % (fm, ps['id'], ps['title'])]
    pre = ps.get('pre', base['pre'] if ps['fm'] == ps['id'] or ps.get('keep_pre') else '')
    if pre: parts.append(pre)
    n = 0
    for k in ps['items']:
        if k.startswith('%%'): parts.append(k); continue
        o, on = k.split(' '); it = dict((i[0], i) for i in SEC[o]['items'])[int(on)]; n += 1
        title = ps.get('retitle', {}).get(k, it[1]); body = it[3]
        if k in ps.get('append', {}): body = body.rstrip('\n') + '\n\n' + ps['append'][k]
        parts.append('### %d. %s%s\n\n%s' % (n, title, ' {%s}' % it[2] if it[2] else '', body.strip('\n')))
    text = conv('\n\n'.join(parts[:2]) + '\n\n' + '\n\n\n'.join(parts[2:]) + '\n')
    path = os.path.join(SRC, ps['dir'], '%s %s.md' % (ps['id'], ps['title']))
    writes[path] = text
old_paths = {SEC[s]['path'] for s in src_secs | set(PLAN.get('rename', {}))}
# 未动的节：只改引用与 related
for s, v in SEC.items():
    if v['path'] in old_paths: continue
    t = io.open(v['path'], encoding='utf-8').read()
    fm = re.match(r'(?s)\A---\n(.*?)\n---\n', t)
    t2 = conv(t)
    if fm: t2 = t2.replace(fm.group(1), fix_related(conv(fm.group(1)), s), 1)
    if t2 != t: writes[v['path']] = t2
for p in old_paths:
    if p not in writes: moves.append(p)

# 章目录与 MANIFEST
final = {}   # id → (dir, title)
for s, v in SEC.items():
    if v['path'] not in old_paths: final[s] = (v['dir'], v['title'])
for ps in plan_secs: final[ps['id']] = (ps['dir'], ps['title'])
for moc in glob.glob(os.path.join(SRC, '[1-6]*', '_*目录.md')):
    d = os.path.basename(os.path.dirname(moc)); t = io.open(moc, encoding='utf-8').read()
    rows = ''.join('| %s | [[%s %s]] |\n' % (s, s, final[s][1]) for s in sorted([x for x in final if final[x][0] == d], key=skey))
    t2 = re.sub(r'(?ms)^(\| 编号 \| 知识点 \|\n\|---\|---\|\n)(?:\|.*\n)*', lambda m: m.group(1) + rows, t, 1)
    if t2 != t: writes[moc] = t2
mf = os.path.join(SRC, 'MANIFEST.txt'); t = io.open(mf, encoding='utf-8').read()
lines = [l for l in t.split('\n') if not re.match(r'^\d+\.\d+    ', l)]
lines[1:1] = ['%s    %s/%s %s.md' % (s, final[s][0], s, final[s][1]) for s in sorted(final, key=skey)]
writes[mf] = '\n'.join(lines)

# 速查源与数据文件
others = [os.path.join(REPO, '速查', '速查源.md'), os.path.join(REPO, '协作', '插图清单.md')]
jsons = [f for f in glob.glob(os.path.join(REPO, '*.json')) + glob.glob(os.path.join(REPO, '速查', '*.json'))
         if os.path.basename(f) not in ('package.json', 'package-lock.json', '相关节索引.json', '排版建议裁定.json', '句末标点裁定.json')]
for f in others:
    if os.path.exists(f):
        t = io.open(f, encoding='utf-8').read(); t2 = conv(t)
        if t2 != t: writes[f] = t2
for f in jsons:
    t = io.open(f, encoding='utf-8').read(); t2 = conv(t, json_style=True)
    if t2 != t: writes[f] = t2
if STUDY:
    for f in glob.glob(os.path.join(os.path.expanduser(STUDY), 'bank', '*.json')):
        t = io.open(f, encoding='utf-8').read(); t2 = conv(t, json_style=True)
        if t2 != t: writes[f] = t2
if PROMPT:
    for b in ('layout-checklist.md', 'content-core.md'):
        f = os.path.join(os.path.expanduser(PROMPT), b); t = io.open(f, encoding='utf-8').read(); t2 = conv(t)
        if t2 != t: writes[f] = t2

# ---------- 4 报告 ----------
REV = {v: k for k, v in IMAP.items()}
print('对照（变化的条目）：')
for ps in plan_secs:
    for k in ps['items']:
        if k.startswith('%%'): continue
        o, on = k.split(' '); ns, nn = IMAP[(o, int(on))]
        if (o, int(on)) != (ns, nn): print('  %s 第 %s 条 → %s 第 %d 条' % (o, on, ns, nn))
print('图改号：', {'%s-%s' % k: '%s-%s' % v for k, v in FMAP.items()} or '无')
print('写入文件 %d 个；整节移走 %d 个' % (len(writes), len(moves)))
for p in sorted(writes): print('  ·', os.path.relpath(p, REPO) if p.startswith(REPO) else p)
if UNRES: print('未能映射的引用（保持原样，需人工看）：', dict(UNRES))
# 正文里仍出现被移走 / 改号节号的纯文字（不是链接、地址、图号）
changed = {s for s in src_secs if SMAP[s] != [s]} | set(PLAN.get('rename', {})) | removed
if changed:
    rx = re.compile(r'(?<![\d.\-/])(%s)(?![\d])(?! 第 \d+ 条)(?!-\d)' % '|'.join(re.escape(s) for s in sorted(changed, key=len, reverse=True)))
    for p, t in list(writes.items()):
        if not p.endswith('.md') or '/notes_src/' not in p and not p.endswith('速查源.md'): continue
        for ln in t.split('\n'):
            if ln.startswith(('id:', 'related:', '# ', '|')) or re.match(r'^\d+\.\d+    ', ln): continue
            s = re.sub(r'\[\[.*?\]\]|figures/\S+|图 \d+\.\d+-\d+', '', ln)
            for m in rx.finditer(s):
                tail = s[m.end():m.end() + 3]
                if re.match(r'\s*(m|ft|kt|psi|%|°|s\b|km|kg|NM|hPa|<|）)', tail) or re.match(r'[\d<]', s[max(0, m.start() - 2):m.start()][-1:] or ' '): continue
                print('  纯文字节号？ %s：%s' % (os.path.basename(p)[:14], ln[:110])); break
MAPOUT = arg('--map-out')
if MAPOUT:
    chg = {'%s %d' % k: '%s %d' % v for k, v in sorted(IMAP.items(), key=lambda kv: (skey(kv[0][0]), kv[0][1])) if k != v}
    old_t = {('%s %d' % (s, it[0])): it[1] for s, v in SEC.items() for it in v['items']}
    json.dump({'from': arg('--from', '?'), 'to': arg('--to', '?'), '说明': '键 = %s「x.y N」，值 = %s「x.y N」；未列出的条号不变；%s' % (arg('--from', '?'), arg('--to', '?'), ('并掉的 %d 条（merged 列出）指向并入后的条目' % len(MERGE)) if MERGE else '本次没有删除条目'), 'map': chg, 'merged': {'%s %d' % k: '%s %d' % IMAP[k] for k in MERGE}},
              io.open(MAPOUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    md = ['# 旧号新号对照表（%s → %s）' % (arg('--from', '?'), arg('--to', '?')), '',
          '> 只列号码变了的 %d 条，未列出的条号不变；机读版是同名 .json。' % len(chg), '', '## 节级变化', '', '| 旧节 | 新节 |', '|---|---|']
    for s in sorted(SEC, key=skey):
        if SMAP.get(s) != [s]: md.append('| %s %s | %s |' % (s, SEC[s]['title'], '、'.join('%s %s' % (x, NEWT.get(x, '')) for x in sorted(SMAP.get(s, []), key=skey))))
    md += ['', '## 逐条对照', '', '| 旧号 | 新号 | 条目 |', '|---|---|---|'] + ['| %s 第 %s 条 | %s 第 %s 条 | %s |' % tuple(k.split(' ') + v.split(' ') + [old_t.get(k, '')]) for k, v in chg.items()]
    io.open(os.path.splitext(MAPOUT)[0] + '.md', 'w', encoding='utf-8').write('\n'.join(md) + '\n')
    print('对照表：%d 条 → %s（及 .md）' % (len(chg), MAPOUT))
if not APPLY:
    print('（未写入；加 --apply 执行）'); sys.exit(0)
os.makedirs(TRASH, exist_ok=True)
for p in moves:
    shutil.move(p, os.path.join(TRASH, os.path.basename(p)))
tmp = []   # 两步改名，避免链式改号互相覆盖（3.6-2 → 3.7-1 时旧 3.7-1 还在）
for (s, n), (ns, nn) in FMAP.items():
    for f in glob.glob(os.path.join(SRC, 'figures', '%s-%s_*' % (s, n))):
        t1 = f + '.mvtmp'; shutil.move(f, t1); tmp.append((t1, os.path.join(SRC, 'figures', os.path.basename(f).replace('%s-%s_' % (s, n), '%s-%s_' % (ns, nn), 1))))
for t1, dst in tmp:
    assert not os.path.exists(dst), dst; shutil.move(t1, dst)
for p, t in writes.items():
    io.open(p, 'w', encoding='utf-8').write(t)
print('已写入；移走的旧节文件 →', TRASH)
