#!/usr/bin/env python3
"""联动检查（SD-114）：每次更新后，查改动有没有按逻辑同步到全笔记。

用法：python3 check_linkage.py --repo <gh-private> [--base <git 引用>] [--out build/联动检查报告.md]

比较基准：--base 未给时，取最近的 baseline/* 标签；没有标签时取 HEAD~1。比较对象是工作区（含未提交改动）。

四项检查：
  一  引用完整性（全书，错误级）：
      [[x.y 标题|x.y A-n]]、[[x.y 标题|x.y 第 n 条]]、正文里的「见 / 详见 x.y A-n」「x.y 第 n 条」，
      速查区里的「见第 N 条」、其他节里的「见第 N 条」（指本节），目标条目必须存在。
  二  旧写法残留（只看本次改动，提醒级）：本次改动里被替换掉的片段（含颜色标记），
      如果在全书其他位置原样还在，列出来。（例：3.6 把「8kt 以上」改成「大于 8kt」，速查 120 还是「8kt 以上」。）
  三  同值异色（只看本次改动涉及的数值，提醒级）：本次改动行里的带单位数值，在全书上下文相近的地方颜色不同。
  四  新增内容的联动面（只看本次新增或改写的行，提醒级）：同一数值 + 相近上下文在其他章节出现、而那一处本次没动的，列出供人工确认。

白名单：<repo>/联动检查白名单.json，格式 [{"type": "残留|异色|联动面", "key": "片段或数值", "where": "文件名片段（可空）", "why": "理由"}]。
只有「一」是错误（退出码 1）；其余是提醒，逐条看过：要改的改，有意不同的写进白名单。
"""
import argparse, json, os, re, subprocess, sys
from difflib import SequenceMatcher

ap = argparse.ArgumentParser()
ap.add_argument('--repo', default='.')
ap.add_argument('--base', default=None)
ap.add_argument('--out', default='build/联动检查报告.md')
A = ap.parse_args()
REPO = os.path.abspath(A.repo)
SRC = os.path.join(REPO, 'notes_src')

def git(*a):
    return subprocess.run(['git', '-C', REPO, *a], capture_output=True, text=True).stdout

base = A.base
if not base:
    tags = [t for t in git('tag', '--list', 'baseline/*', '--sort=-creatordate').split() if t]
    base = tags[0] if tags else 'HEAD~1'

# ---------- 读全书 ----------
files = {}
for root, _, fs in os.walk(SRC):
    if '/_bak' in root or '/_附件' in root:
        continue
    for f in fs:
        if f.endswith('.md') and not f.startswith('_') and f != 'MANIFEST.txt':
            p = os.path.join(root, f)
            files[os.path.relpath(p, REPO)] = open(p, encoding='utf-8').read().split('\n')

TAG = re.compile(r'<[^>]+>')
plain = lambda s: TAG.sub('', s).replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&')

def sec_id(rel):
    m = re.search(r'/(\d+\.\d+) ', '/' + os.path.basename(rel))
    return m.group(1) if m else ('0' if '速查区' in rel else None)

# 每节的条目集合
items = {}
for rel, L in files.items():
    sid = sec_id(rel)
    s = set()
    for l in L:
        m = re.match(r'^###\s+([A-Z]-\d+)\s', l) or re.match(r'^###\s+(\d+)\.\s', l)
        if m:
            s.add(m.group(1))
    items[sid] = s
QR = [r for r in files if '速查区' in r]
QRF = QR[0] if QR else None

def short(rel):
    return os.path.basename(rel)[:-3]

# ---------- 一、引用完整性 ----------
errs = []
pat_link = re.compile(r'\[\[(\d+\.\d+)[^\]|]*\|(\d+\.\d+) ([A-Z]-\d+|第 ?(\d+) ?条)\]\]')
pat_txt = re.compile(r'(?:详见|见)\s*(?:<strong>)?(\d+\.\d+)\s*([A-Z]-\d+)')
pat_num = re.compile(r'(\d+\.\d+)\s*第\s*(\d+)\s*条')
pat_self = re.compile(r'见第\s*(\d+)\s*条')
for rel, L in files.items():
    sid = sec_id(rel)
    for n, l in enumerate(L, 1):
        for m in pat_link.finditer(l):
            tgt = m.group(2); key = m.group(4) or m.group(3)
            if tgt not in items:
                errs.append((rel, n, f'「{m.group(0)}」指向的节 {tgt} 不存在'))
            elif key not in items[tgt]:
                errs.append((rel, n, f'「{m.group(0)}」指向的条目 {tgt} {m.group(3)} 不存在'))
        lp = plain(l)
        for m in pat_txt.finditer(lp):
            tgt, key = m.group(1), m.group(2)
            if tgt in items and key not in items[tgt]:
                errs.append((rel, n, f'「{m.group(0)}」指向的条目不存在'))
        for m in pat_num.finditer(lp):
            tgt, key = m.group(1), m.group(2)
            if tgt in items and items[tgt] and key not in items[tgt]:
                errs.append((rel, n, f'「{m.group(0)}」指向的条目不存在'))
        if '[[' not in l:
            for m in pat_self.finditer(lp):
                if lp[max(0, m.start() - 6):m.start()].rstrip().endswith(('条', '节')) or re.search(r'\d+\.\d+\s*$', lp[:m.start()]):
                    continue
                key = m.group(1)
                if key not in items.get(sid, set()):
                    errs.append((rel, n, f'「{m.group(0)}」在本{"区" if sid == "0" else "节"}找不到第 {key} 条'))

# ---------- 改动 ----------
diff = subprocess.run(['git', '-C', REPO, 'diff', '-U0', base, '--', 'notes_src'], capture_output=True, text=True).stdout
hunks = []   # (rel, [removed], [added])
cur = None; rel = None
for ln in diff.split('\n'):
    if ln.startswith('+++ '):
        p = ln[4:].strip()
        rel = None if p == '/dev/null' else p[2:] if p.startswith('b/') else p
        if rel and rel.startswith('"'):
            rel = bytes(rel.strip('"'), 'utf-8').decode('unicode_escape').encode('latin-1').decode('utf-8')
        continue
    if ln.startswith('@@'):
        cur = [rel, [], []]; hunks.append(cur); continue
    if cur is None or ln.startswith('--- ') or ln.startswith('diff ') or ln.startswith('index '):
        continue
    if ln.startswith('-'):
        cur[1].append(ln[1:])
    elif ln.startswith('+'):
        cur[2].append(ln[1:])
changed_rels = {h[0] for h in hunks if h[0]}
added_lines = {(h[0], a) for h in hunks for a in h[2]}

wl = []
wp = os.path.join(REPO, '联动检查白名单.json')
if os.path.exists(wp):
    wl = json.load(open(wp, encoding='utf-8'))
def allowed(typ, key, where):
    return any(w.get('type') == typ and w.get('key') == key and (not w.get('where') or w['where'] in where) for w in wl)

# ---------- 二、旧写法残留 ----------
BOUND = set('，。；：、（）「」【】<>／/ \t')
def frags(a, b):
    out = []
    sm = SequenceMatcher(None, a, b, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op in ('replace', 'delete'):
            l, r = i1, i2
            # 扩到完整标签和左右各最多 4 个可见字
            while l > 0 and a[l - 1] != '>' and a[l - 1] not in BOUND and i1 - l < 4: l -= 1
            while r < len(a) and a[r] not in BOUND and r - i2 < 4: r += 1
            seg = a[l:r]
            # 不切断标签
            if seg.count('<') != seg.count('>'):
                ll = a.rfind('<', 0, l + 1); rr = a.find('>', r - 1)
                if ll >= 0 and rr >= 0: seg = a[ll:rr + 1]
            vis = plain(seg).strip()
            # 只有数字和标点的片段（条目重编号、页码）不算写法变化
            if not re.search(r'[\u4e00-\u9fffA-Za-z]', vis):
                continue
            if len(vis) >= 3 and seg not in b:
                out.append(seg)
    return out

left = []
seen = set()
for rel_h, rem, add in hunks:
    if not rel_h or not rem or not add:
        continue
    for r_line in rem:
        if re.match(r'^### \d+\. ', r_line) or re.match(r'^<!-- 详见', r_line):
            continue   # 速查条目标题重编号、详见注释不算
        best = max(add, key=lambda x: SequenceMatcher(None, r_line, x, autojunk=False).ratio())
        if SequenceMatcher(None, r_line, best, autojunk=False).ratio() < 0.5:
            continue
        for fr in frags(r_line, best):
            if fr in seen: continue
            seen.add(fr)
            hits = []
            for rel2, L in files.items():
                for n, l in enumerate(L, 1):
                    if fr in l and (rel2, l) not in added_lines:
                        hits.append((rel2, n, l))
            hits = [h for h in hits if not allowed('残留', fr, h[0])]
            if hits:
                left.append((fr, rel_h, hits[:8]))
# 同一批命中只报一次：保留含数字或更长的片段
uniq = {}
for fr, rel_h, hits in left:
    k = frozenset((h[0], h[1]) for h in hits)
    old = uniq.get(k)
    score = (bool(re.search(r'\d', plain(fr))), len(plain(fr)))
    if not old or score > old[0]:
        uniq[k] = (score, (fr, rel_h, hits))
left = [v[1] for v in uniq.values()]

# ---------- 三、同值异色 / 四、联动面 ----------
UNIT = r'(?:kt|ft|fpm|psi|min|s|m|kg|°|%|℃|Hz|V|nm|NM|lb|英尺|节)'
NUM = re.compile(r'(?<![\d.])(\d+(?:\.\d+)?)\s?(' + UNIT + r')(?![A-Za-z])')
CJK = re.compile(r'[一-鿿]')
def color_of(raw, pos_plain_tok):
    return None
def num_ctx(raw_line):
    """返回 [(token, color, ctx_bigrams)]"""
    out = []
    # 按标签切分，记录每段颜色
    segs = re.split(r'(<em>|</em>|<b>|</b>|<strong>|</strong>)', raw_line)
    col = []; buf = ''; spans = []
    for sgm in segs:
        if sgm in ('<em>', '<b>', '<strong>'): col.append({'<em>': '红', '<b>': '蓝', '<strong>': '黑粗'}[sgm]); continue
        if sgm in ('</em>', '</b>', '</strong>'):
            if col: col.pop()
            continue
        t = plain(sgm); spans.append((len(buf), len(buf) + len(t), col[-1] if col else '无')); buf += t
    for m in NUM.finditer(buf):
        c = '无'
        for s0, s1, cc in spans:
            if s0 <= m.start() < s1: c = cc
        ctx = ''.join(CJK.findall(buf[max(0, m.start() - 10):m.start()]))[-6:]
        big = {ctx[i:i + 2] for i in range(len(ctx) - 1)}
        out.append((m.group(1) + m.group(2), c, big, buf[max(0, m.start() - 14):m.end() + 6]))
    return out

def same_text(snip, other_raw):
    """新写法在别处已经一字不差（取数值前后各约 8 字），视为已一致"""
    core = snip.strip()
    return core[-16:] in plain(other_raw) or core[:16] in plain(other_raw) and core[-8:] in plain(other_raw)

index = {}
for rel2, L in files.items():
    for n, l in enumerate(L, 1):
        if l.startswith('来源：') or l.startswith('出处：') or l.startswith('---'):
            continue
        for tok, c, big, snip in num_ctx(l):
            index.setdefault(tok, []).append((rel2, n, c, big, snip, l))

clash, reach = [], []
done = set()
for rel_h, rem, add in hunks:
    if not rel_h: continue
    # 速查条目自己「详见」指向的正文节，本来就该一样，不算联动面 / 异色
    own = set(re.findall(r'\[\[(\d+\.\d+) ', '\n'.join(add)))
    for a in add:
        for tok, c, big, snip in num_ctx(a):
            if len(big) < 2: continue
            key = (tok, frozenset(big))
            if key in done: continue
            done.add(key)
            for rel2, n, c2, big2, snip2, l2 in index.get(tok, []):
                if (rel2, l2) in added_lines: continue
                if sec_id(rel2) in own: continue
                if len(big & big2) < 2: continue
                if c2 != c and not (c == '无' and c2 == '无'):
                    if not allowed('异色', tok, rel2):
                        clash.append((tok, rel_h, c, snip, rel2, n, c2, snip2))
                elif rel2 not in changed_rels and sec_id(rel2) != sec_id(rel_h) and not same_text(snip, l2):
                    if not allowed('联动面', tok, rel2):
                        reach.append((tok, rel_h, snip, rel2, n, snip2))

# ---------- 报告 ----------
o = [f'# 联动检查报告', '', f'比较基准：`{base}` → 工作区。改动文件 {len(changed_rels)} 个、改动片段 {len(hunks)} 处。', '',
     f'**错误 {len(errs)} 条（必须改）**；提醒：旧写法残留 {len(left)}、同值异色 {len(clash)}、联动面 {len(reach)}（逐条看过：该改的改，有意不同的写进 `联动检查白名单.json`）。', '']
o += ['## 一、引用完整性（错误）', '']
o += [f'- {short(r)} 第 {n} 行：{m}' for r, n, m in errs] or ['- 无']
o += ['', '## 二、旧写法残留（本次改掉的写法，别处还在）', '']
if not left: o.append('- 无')
for fr, rel_h, hits in left:
    o.append(f'- 「{plain(fr)}」（原文 `{fr}`，本次在 {short(rel_h)} 改掉）仍见于：')
    o += [f'  - {short(r)} 第 {n} 行：{plain(l).strip()[:90]}' for r, n, l in hits]
o += ['', '## 三、同值异色（本次改动的数值，上下文相近处颜色不同）', '']
o += [f'- {tok}：{short(r1)}【{c1}】「{s1}」 ↔ {short(r2)} 第 {n} 行【{c2}】「{s2}」' for tok, r1, c1, s1, r2, n, c2, s2 in clash[:60]] or ['- 无']
o += ['', '## 四、联动面（同一数值与相近上下文出现在其他章节、本次没动，供人工确认）', '']
o += [f'- {tok}：{short(r1)}「{s1}」 → {short(r2)} 第 {n} 行「{s2}」' for tok, r1, s1, r2, n, s2 in reach[:80]] or ['- 无']
os.makedirs(os.path.dirname(os.path.join(REPO, A.out)), exist_ok=True)
open(os.path.join(REPO, A.out), 'w', encoding='utf-8').write('\n'.join(o) + '\n')
print(f'联动检查：错误 {len(errs)}，残留 {len(left)}，异色 {len(clash)}，联动面 {len(reach)}（基准 {base}）→ {A.out}')
sys.exit(1 if errs else 0)
