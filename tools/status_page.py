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


def pages(pdf):
    try:
        import fitz
        return len(fitz.open(pdf))
    except Exception:
        return '—'


def repo(path):
    head = sh('git', 'log', '-1', '--format=%h %ad %s', '--date=format:%m-%d %H:%M', cwd=path)
    dirty = [l for l in sh('git', 'status', '--porcelain', cwd=path).splitlines() if l]
    ahead = sh('git', 'rev-list', '--count', '@{u}..HEAD', cwd=path) or '?'
    return head, len(dirty), ahead


L = ['# 现状一页纸', '',
     '> 自动生成（`pub/tools/status_page.py`），%s。只读汇总，不是规则；规则看 `gh-private/规则/` 三份现行规则。' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), '']

# 基线
tags = sh('git', 'tag', '-l', 'baseline/*', '--sort=-creatordate', cwd=G).splitlines()
base = tags[0] if tags else '（无）'
since = sh('git', 'rev-list', '--count', '%s..HEAD' % base, cwd=G) if tags else '?'
eds = sorted({re.sub(r'\.(docx|pdf)$', '', os.path.basename(p)) for p in glob.glob(os.path.join(WS, 'B737机型理论*.pdf'))})
L += ['## 基线与成品', '',
      '| 项目 | 现状 |', '|---|---|',
      '| 现行基线标签 | `%s`（%s）；之后私有库又有 %s 个提交 |' % (base, sh('git', 'rev-list', '-n1', '--abbrev-commit', base, cwd=G) if tags else '—', since),
      '| 桌面成品 | %s |' % ('、'.join(eds) or '（无）')]
for nm in ('B737机型理论速查笔记',):   # SD-234
    p = os.path.join(G, 'build', nm + '.pdf')
    if os.path.exists(p):
        L.append('| build/%s.pdf | %s 页，生成于 %s |' % (nm, pages(p), datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%m-%d %H:%M')))
try: _lag = int(open(os.path.join(G, 'build', '完整版落后轮数.txt')).read().strip() or 0)
except Exception: _lag = 0
L.append('| 完整版落后速查 | %d 轮速查构建%s |' % (_lag, '——**已满 5 轮，请提醒用户全量收口（SD-227）**' if _lag >= 5 else '（满 5 轮提醒全量收口，SD-227）'))
L.append('')

# 两本规模
secs = [f for f in glob.glob(os.path.join(G, 'notes_src', '[1-9] *', '*.md')) if re.match(r'\d+\.\d+ ', os.path.basename(f))]
ents = sum(len(re.findall(r'(?m)^### \d+\. ', open(f, encoding='utf-8').read())) for f in secs)
chs = sorted({os.path.basename(os.path.dirname(f)) for f in secs})
q = open(os.path.join(G, '速查', '速查源.md'), encoding='utf-8').read()
L += ['## 两本规模', '',
      '| 本 | 结构 | 条目 |', '|---|---|---|',
      '| 完整版 | %d 章 %d 节（%s） | %d 条 |' % (len(chs), len(secs), '、'.join(chs), ents),
      '| 速查 | %d 段 %d 个主题 | %d 条 |' % (len(re.findall(r'(?m)^%%PART%%', q)), len(re.findall(r'(?m)^## ', q)), len(re.findall(r'(?m)^### ', q))), '']

# 检查
rep = os.path.join(G, 'build', '排版检查报告.md')
if os.path.exists(rep):
    head = open(rep, encoding='utf-8').read().split('\n')[2]
    L += ['## 上次全量检查', '', head, '']

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
top = open(os.path.join(WS, '00_先读我_AI交接.md'), encoding='utf-8').read()
m = re.search(r'(?m)^### (.+)$', top)
if m:
    todo.append('- 最近一轮：%s' % m.group(1))
L += todo + ['']

open(os.path.join(WS, '现状一页纸.md'), 'w', encoding='utf-8').write('\n'.join(L))
print('已生成 现状一页纸.md')
