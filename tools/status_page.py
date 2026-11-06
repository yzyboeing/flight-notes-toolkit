#!/usr/bin/env python3
"""生成《现状一页纸.md》（规则重组轮方案 A8，2026-10-21）：给任何接手的 AI 一页看清现在的状态。
用法：python3 ~/flight-repos/pub/tools/status_page.py   （任意位置运行；每轮收尾、出基线后重跑）
只读各仓库与成品，不改任何源文件；输出写到工作区根目录《现状一页纸.md》。"""
import datetime, glob, os, re, subprocess

HOME = os.path.expanduser('~')
WS = os.path.join(HOME, 'Desktop', '飞行理论笔记整理')
G = os.path.join(HOME, 'flight-repos', 'gh-private')
PUB = os.path.join(HOME, 'flight-repos', 'pub')


def sh(*a, cwd=None):
    try:
        return subprocess.run(a, cwd=cwd, capture_output=True, text=True, timeout=60).stdout.strip()
    except Exception:
        return ''


def repo(path):
    head = sh('git', 'log', '-1', '--format=%h %ad %s', '--date=format:%m-%d %H:%M', cwd=path)
    dirty = [l for l in sh('git', 'status', '--porcelain', cwd=path).splitlines() if l]
    ahead = sh('git', 'rev-list', '--count', '@{u}..HEAD', cwd=path) or '?'
    return head, len(dirty), ahead


L = ['# 现状一页纸', '',
     '> 自动生成（`pub/tools/status_page.py`），%s。只读汇总，不是规则；规则看 `gh-private/规则/` 三份现行规则。' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), '']

# 以当前内容提交和已成功的本机上传日志为准。
def app_status():
    flight = os.path.join(HOME, 'flight-repos')
    content = sh('git', 'log', '-1', '--format=%H', 'HEAD', '--', 'notes_src', '速查/速查源.md', cwd=G)
    roots = glob.glob(os.path.join(flight, '_work', '学习系统快照_*', '基线.txt'))
    matching = []
    for meta in roots:
        with open(meta, encoding='utf-8') as f:
            revision = re.search(r'(?m)^提交：gh-private ([0-9a-f]+)', f.read())
        if content and revision and revision[1] == content:
            matching.append(meta)
    snapshot = os.path.basename(os.path.dirname(max(matching, key=os.path.getmtime))) if matching else '尚无匹配快照'
    uploads = []
    for log in glob.glob(os.path.join(flight, '_work', 'archives', '737note_*_upload.log')):
        version = re.fullmatch(r'737note_(\d+(?:\.\d+)*)_(\d+)_upload\.log', os.path.basename(log))
        if not version:
            continue
        with open(log, encoding='utf-8', errors='replace') as f:
            result = f.read()
        if 'Upload succeeded' in result and 'EXPORT SUCCEEDED' in result:
            uploads.append((os.path.getmtime(log), version[1], version[2]))
    tf = '本机无成功上传记录'
    if uploads:
        _, version, number = max(uploads)
        tf = f'{version} ({number}) 已上传（本机日志，未查询 Apple 当前处理状态）'
    return content[:7] or '未知', snapshot, tf


content, snapshot, tf = app_status()
L += ['## App 快照与 TestFlight', '',
      '| 项目 | 现状 |', '|---|---|',
      '| 当前内容提交 | `%s` |' % content,
      '| 与当前内容一致的 App 快照 | `%s` |' % snapshot,
      '| 最近成功上传的 TestFlight | %s |' % tf, '']

# 两本规模
secs = [f for f in glob.glob(os.path.join(G, 'notes_src', '[1-9] *', '*.md')) if re.match(r'\d+\.\d+ ', os.path.basename(f))]
ents = sum(len(re.findall(r'(?m)^### \d+\. ', open(f, encoding='utf-8').read())) for f in secs)
chs = sorted({os.path.basename(os.path.dirname(f)) for f in secs})
q = open(os.path.join(G, '速查', '速查源.md'), encoding='utf-8').read()
L += ['## 两本规模', '',
      '| 本 | 结构 | 条目 |', '|---|---|---|',
      '| 完整版 | %d 章%s %d 节（%s） | %d 条 |' % (len([c for c in chs if '附录' not in c]), ' + 附录' if any('附录' in c for c in chs) else '', len(secs), '、'.join(chs), ents),
      '| 旧速查源（SD-234 起只作溯源） | %d 段 %d 个主题 | %d 条 |' % (len(re.findall(r'(?m)^%%PART%%', q)), len(re.findall(r'(?m)^## ', q)), len(re.findall(r'(?m)^### ', q))), '']

# 仓库
L += ['## 仓库', '', '| 仓库 | 最新提交 | 未提交文件 | 未推送提交 |', '|---|---|---|---|']
for nm, p in (('私有笔记库 gh-private', G), ('公开工具库 pub', PUB), ('工作区', WS)):
    h, d, a = repo(p)
    L.append('| %s | %s | %d | %s |' % (nm, h, d, a))
L.append('')

# 规则
sd = sorted(set(int(x) for x in re.findall(r'SD-(\d+)', open(os.path.join(PUB, 'prompt', 'standing-decisions.md'), encoding='utf-8').read())))
L += ['## 规则', '', '- 长期决策编到 SD-%d；现行规则：总纲、排版规则总表、内容校对核心规则（`pub/prompt/`）。' % (sd[-1] if sd else 0), '']

# 待办
todo = ['## 待办与待用户定', '']
opin = sorted(glob.glob(os.path.join(WS, '优化意见记录_*.md')), key=os.path.getmtime)
if opin:
    t = open(opin[-1], encoding='utf-8').read()
    rows = [r for r in t.split('\n') if re.match(r'\|\s*\d+\s*\|', r)]
    pend = [r.split('|')[1].strip() for r in rows if not re.search(r'已执行|不改|作废|已撤回', r.split('|')[-2] if r.count('|') > 2 else '')]
    todo.append('- 当期意见记录 `%s`：%d 条，未执行 %d 条%s' % (os.path.basename(opin[-1]), len(rows), len(pend), ('（第 %s 条）' % '、'.join(pend[:15])) if pend else ''))
nk = os.path.join(WS, '新知识点待入库清单.md')
if os.path.exists(nk):
    part = open(nk, encoding='utf-8').read().split('## 待入库', 1)[-1].split('## 已入库', 1)[0]
    todo.append('- 新知识点待入库：%d 张卡' % len(re.findall(r'【知识点卡】', part)))
one = os.path.join(WS, 'OneNote待同步记录.md')
if os.path.exists(one):
    pend = [l for l in open(one, encoding='utf-8') if '待同步' in l]
    todo.append('- OneNote：%d 条记录标「待同步」（只在用户指令时更新）' % len(pend))
L += todo + ['']

# 保留人工维护的剩余批次区；重新生成只替换自动段。
dest = os.path.join(WS, '现状一页纸.md')
marker = '<!-- status-page:manual -->'
if os.path.isfile(dest):
    with open(dest, encoding='utf-8') as f:
        old = f.read()
    _, sep, manual = old.partition(marker)
    if sep:
        L += [marker + manual.rstrip('\n'), '']
with open(dest, 'w', encoding='utf-8') as f:
    f.write('\n'.join(L))
print('已生成 现状一页纸.md')
