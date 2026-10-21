#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""muse_import_check.py —— Muse 变更单导入前检查（SD-81，2026-09-30 用户）
用法：python3 muse_import_check.py <变更单.md> [--repo ~/flight-repos/gh-private]
只读，不改任何文件。逐条检查：
  1  基线：变更单写的基线版本号 / 提交号与仓库记录是否一致；基线之后 notes_src 改过哪些节（可能冲突）
  2  修改 / 删除：「现文」去掉标签和空白后，在 notes_src 里是否恰好出现 1 处（0 处＝基线后已改或抄错；多处＝需加长现文）
  3  改为：HTML 标签是否成对、有无同名嵌套、行首项目符号、正文里的「待补 / 待核」字样
  4  多条变更落在同一处 → 提示合并顺序；「取代」关系是否闭合
  5  汇总：各条的待核数值（这些都必须回手册库核实）、结构提议（须先报用户批准）
输出按条列「可直接落地 / 需处理」，最后给总表。真正落地仍按 AI交接/10_Muse变更单导入流程.md 手工逐条做。"""
import sys, os, re, glob, subprocess, collections

def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
if len(sys.argv) < 2 or not os.path.isfile(sys.argv[1]): print(__doc__); sys.exit(1)
CL = sys.argv[1]
REPO = os.path.abspath(os.path.expanduser(arg('--repo', '~/flight-repos/gh-private')))
text = open(CL, encoding='utf-8').read()

def plain(s):
    s = re.sub(r'<!--.*?-->', '', s, flags=re.S)
    s = re.sub(r'<[^>]+>', '', s)
    s = s.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&nbsp;', ' ')
    return re.sub(r'[\s•▪–]+', '', s)

# ---------- 1 基线 ----------
issues_global = []
m = re.search(r'基线[：:]\s*(?:版本号\s*)?(\d{8})(?:\s*/\s*gh-private\s*([0-9a-f]{7,40}))?', text)
base_ed, base_commit = (m.group(1), m.group(2)) if m else (None, None)
cur_ed = subprocess.run(['git', '-C', REPO, 'config', '--get', 'notes.docEdition'], capture_output=True, text=True).stdout.strip()
print('变更单：%s' % CL)
print('基线：版本号 %s，提交 %s；仓库当前版本号 %s' % (base_ed, base_commit, cur_ed))
changed_secs = []
if base_commit:
    r = subprocess.run(['git', '-C', REPO, '-c', 'core.quotepath=off', 'diff', '--name-only', base_commit, 'HEAD', '--', 'notes_src'],
                       capture_output=True, text=True)
    if r.returncode: issues_global.append('基线提交 %s 在仓库里找不到' % base_commit)
    changed_secs = [re.sub(r'^notes_src/[^/]+/', '', x).split(' ')[0] for x in r.stdout.split() if x.endswith('.md')]
    dirty = subprocess.run(['git', '-C', REPO, 'status', '--porcelain', 'notes_src'], capture_output=True, text=True).stdout.strip()
    if dirty: issues_global.append('notes_src 有未提交改动，先处理干净再导入')
    print('基线之后改过的节：%s' % ('、'.join(changed_secs) or '无'))
else:
    issues_global.append('变更单没写基线提交号，无法判断冲突')

src = {}
for f in glob.glob(os.path.join(REPO, 'notes_src', '[0-5]*', '*.md')):
    src[os.path.basename(f)[:-3]] = plain(open(f, encoding='utf-8').read())

# ---------- 2-4 逐条 ----------
items = re.split(r'(?m)^## (C-\d{3})', text)
recs = []
for k in range(1, len(items), 2):
    cid, body = items[k], items[k + 1]
    head = body.split('\n', 1)[0]
    kind = next((t for t in ('新增', '修改', '删除', '结构提议') if t in head), '?')
    def sect(name):
        mm = re.search(r'(?ms)^### ' + name + r'.*?\n(.*?)(?=^### (?:现文|改为|说明)|\Z)', body)   # 改为里本身有 ### 标题，只按三个固定栏名切分
        return mm.group(1).strip() if mm else ''
    cur, new = sect('现文'), sect('改为')
    code = re.search(r'```(?:html|markdown)?\n(.*?)```', new, re.S)
    new_src = code.group(1) if code else new
    field = lambda n: (re.search(r'(?m)^- ' + n + r'[：:]\s*(.*)$', body) or [None, ''])[1].strip()
    status, todo = field('状态'), field('待核')
    probs, notes = [], []
    if '已确认' not in status and '取代' not in status: probs.append('状态不是「用户已确认」：%s' % (status or '空'))
    if '已被' in status and '取代' in status: notes.append('已被取代，跳过')
    loc = None
    if kind in ('修改', '删除'):
        p = plain(cur)
        if len(p) < 6: probs.append('「现文」太短或为空，无法唯一定位')
        else:
            hits = [(n, s.count(p)) for n, s in src.items() if p in s]
            total = sum(c for _, c in hits)
            if total == 0: probs.append('「现文」在 notes_src 里找不到（基线后已改，或抄写有误）')
            elif total > 1: probs.append('「现文」出现 %d 处（%s），需要加长以唯一定位' % (total, '、'.join(n for n, _ in hits)))
            else: loc = hits[0][0]; notes.append('定位：%s' % loc)
            if loc and loc.split(' ')[0] in changed_secs: probs.append('所在节 %s 在基线之后改过，需人工比对冲突' % loc.split(' ')[0])
    if kind in ('新增', '修改'):
        if not new_src.strip() or new_src.strip() == '（删除）': probs.append('「改为」为空')
        for tag in ('em', 'strong', 'b', 'td', 'th', 'tr', 'table'):
            o, c = len(re.findall(r'<%s[\s>]' % tag, new_src)), len(re.findall(r'</%s>' % tag, new_src))
            if o != c: probs.append('<%s> 开 %d 闭 %d，不成对' % (tag, o, c))
        for tag in ('em', 'strong', 'b'):
            if re.search(r'<%s>(?:(?!</%s>).)*<%s>' % (tag, tag, tag), new_src, re.S): probs.append('<%s> 同名嵌套' % tag)
        if re.search(r'(?m)^\s*[•▪\-\*]\s', new_src): probs.append('行首有项目符号（源文件禁用）')
        if re.search(r'待补|待核|资料库内无原文', plain(new_src)): probs.append('「改为」正文里有「待补 / 待核」字样（应写在待核栏）')
    if kind == '新增':
        sec = re.search(r'(\d\.\d{1,2})', head)
        if sec and sec.group(1) in changed_secs: notes.append('目标节 %s 在基线后改过，核对提议的条目号' % sec.group(1))
    recs.append({'id': cid, 'kind': kind, 'head': re.sub(r'^[\s｜|]*(新增|修改|删除|结构提议)[\s｜|]*', '', head), 'probs': probs, 'notes': notes, 'todo': todo, 'loc': loc})

# 同一处多条变更
bylo = collections.defaultdict(list)
for r in recs:
    if r['loc']: bylo[r['loc']].append(r['id'])
for loc, ids in bylo.items():
    if len(ids) > 1:
        for r in recs:
            if r['id'] in ids: r['notes'].append('与 %s 落在同一节，按编号顺序依次落地' % '、'.join(i for i in ids if i != r['id']))

# ---------- 5 输出 ----------
print()
for g in issues_global: print('！' + g)
ok = [r for r in recs if not r['probs']]
print('共 %d 条：%s' % (len(recs), '、'.join('%s %d' % (k, v) for k, v in collections.Counter(r['kind'] for r in recs).items())))
print('格式与定位无问题 %d 条，需处理 %d 条\n' % (len(ok), len(recs) - len(ok)))
for r in recs:
    print('%s｜%s｜%s' % (r['id'], r['kind'], r['head'][:60]))
    for p in r['probs']: print('    ✗ ' + p)
    for n in r['notes']: print('    · ' + n)
    if r['todo'] and r['todo'] not in ('无', '—'): print('    待核（必须回手册库核实）：' + r['todo'])
struct = [r['id'] for r in recs if r['kind'] in ('删除', '结构提议')]
if struct: print('\n须先列清单报用户批准（删除 / 结构 / 编号）：' + '、'.join(struct))
sys.exit(1 if issues_global or len(ok) < len(recs) else 0)
