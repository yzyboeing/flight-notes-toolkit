#!/usr/bin/env python3
"""K-RULES 规则自检（2026-10-21 规则重组轮，方案 A1）。

用法：check_rules.py [笔记库路径] [--no-copies | --copies-only]   ——问题打印「RULES 错误 …」，退出码 1；全部通过退出码 0
（check_layout 调用时加 --no-copies：副本在 sync 收尾才复制，sync 复制后再跑 --copies-only）

查三件事：
1. 过时说法：现行规则与入口文件里不得再出现已作废的口径（rules.md、第 1～5 章、只本地提交绝不推送、块索引地址等）。
   写在「已作废」表里、或同一行带「作废 / 不再 / 已删 / 不设 / 原 / 旧」这类否定说明的不算。
2. K- 规则号与检查器一一对应：《排版规则总表》「检查器规则号」表里每个 K- 号都要在某个检查脚本里找得到；
   检查脚本 err() / sug() / rep() 报出的规则号都要登记在表里。
3. 规则副本与真本一致：私有库 规则/、桌面 Muse交接包/ 里的副本（存在时）与 pub/prompt 真本逐字相同。
"""
import os, re, sys
from pathlib import Path

PUB = Path(__file__).resolve().parents[1]
TOOLS = PUB / 'tools'
ARGS = [a for a in sys.argv[1:] if not a.startswith('--')]
NOTES = Path(ARGS[0]).resolve() if ARGS else PUB.parent / 'gh-private'
NO_COPIES, COPIES_ONLY = '--no-copies' in sys.argv, '--copies-only' in sys.argv
HB = Path.home() / 'Desktop/飞行理论笔记整理/Muse交接包'

CURRENT = ['SKILL.md', 'AGENTS.md', 'prompt/AGENT.md', 'prompt/portable-instruction.md',
           'prompt/layout-checklist.md', 'prompt/content-core.md', 'prompt/lessons.md']
PRIVATE = ['START_HERE.md', 'AGENTS.md', '规则/示例.md']

# 过时说法：(正则, 说明)
STALE = [
    (r'prompt/rules\.md|整理规则（真源）', '旧整理规则 rules.md 已归档，现行规则是总纲 / 排版规则总表 / 内容校对核心规则'),
    (r'5 章 76 节|第二章\s*机组训练手册|2 机组训练手册|第五章\s*模拟机训练|5 模拟机训练', 'SD-234（2026-10-24）：新框架 5 章 + 附录 72 节，第二章「操作程序」、第五章「模拟机笔记」'),
    (r'4 章 60 节|第四章\s*模拟机训练|4 模拟机训练|第五章\s*技术提示|5 技术提示', '2026-10-23 两本整合：完整版 5 章 76 节——第四章 飞行技术提示、第五章 模拟机训练'),
    (r'第三章\s*运行手册|3 运行手册|第四章\s*飞行技术提示|4 飞行技术提示|第五章\s*模拟机笔记|5 模拟机笔记', '2026-10-25 用户：章名改为 第三章 运行规范 / 第四章 技术提示 / 第五章 故障处置'),
    (r'第四章\s*技术提示|4 技术提示 \||06 技术提示', 'SD-241（2026-10-28）：第四章改名「特殊情况与风险」（7 节）；速查目录 06 段同名'),
    (r'绝不推送|不远程 git push', '出基线时推送两库并打标签（总纲硬约束第 7 条）'),
    (r'[A-H]-\d+\s*表|C 块|块索引地址|「A-1」', '全书不设块，条目写「x.y 第 N 条」（SD-199 / 201）'),
    (r'0 基础知识速查区|速查区（第零章', '旧第零章速查区已删除，速查独立成册'),
    (r'规则/排版检查清单\.md', '私有库副本已改名 规则/排版规则总表.md'),
    (r'必背清单（★★）|限制数据总表」和', '必背清单、限制数据总表已删（SD-213）'),
    # 2026-10-23 规则统一整理（SD-230）：已作废 / 已改的口径，现行规则里再出现就报
    (r'注\s*≤\s*3|一条目注不超过 3', '注不再限 3 条、不建注表（B37 / Q06，SD-230）'),
    (r'原地只差一点|原地收紧（≤\s*8pt）|方案 C）', '9pt 一页放得下的表不压缩、整表移页（G6，SD-229）'),
    (r'4～5 级 7\.5pt|4 级＝7\.5pt', '4 级是 8pt 收行距（G6，SD-230）'),
    (r'任何情况\s*≥\s*7\.5pt|表内字号\s*≥\s*8pt', '压缩底线 7pt（G6）'),
    (r'约\s*17\s*分钟', '全量重建现在约 40 分钟'),
    (r'B737机型理论知识\{?(笔记|速查)', '成品名已改为《B737机型理论速查笔记》（SD-234；SD-225 的《详解》《速查》两本只作溯源）'),
    (r'SD-1～217', '长期决策已到 SD-230，索引见 decisions-index.md'),
    (r'使用说明」一页|含学习路线：初学', '两本都不设使用说明（L05）'),
]
EXEMPT = re.compile(r'作废|不再|已删|不设|已并入|原第|旧|删除|去掉|取消|改名|不收|无前言|不标|没有|已归档')

CHECKERS = ['check_layout.py', 'check_struct.py', 'check_src.py', 'check_blocks.py', 'check_quickref.py',
            'check_pair.py', 'check_userpins.py', 'check_linkage.py', 'check_figs.py', 'check_rules.py',
            'verify.py', 'check_splits.py', 'check_coverage.py', 'punct_rule.py', 'fit_fix.py']

errs = []


def stale_scan(path, label):
    if not path.is_file():
        return
    in_void = False
    for i, line in enumerate(path.read_text(encoding='utf-8').splitlines(), 1):
        if line.startswith('#'):
            in_void = '作废' in line
        if in_void or EXEMPT.search(line):
            continue
        for pat, why in STALE:
            m = re.search(pat, line)
            if m:
                errs.append('%s:%d 过时说法「%s」——%s' % (label, i, m.group(0), why))


if not COPIES_ONLY:
    for rel in CURRENT:
        stale_scan(PUB / rel, rel)
    for rel in PRIVATE:
        stale_scan(NOTES / rel, '私有库/' + rel)

# ---- 2. K- 规则号 ↔ 检查器 ----
lc = (PUB / 'prompt/layout-checklist.md').read_text(encoding='utf-8')
sec = lc.split('## 检查器规则号（K-）', 1)
if len(sec) < 2:
    errs.append('排版规则总表找不到「## 检查器规则号（K-）」一节')
    table = ''
else:
    table = sec[1].split('\n## ', 1)[0]


def expand(cell):
    """K-T20～T23 / K-T6、T7 / K-P1、P2（两本） → ['T20','T21',...]"""
    cell = re.sub(r'（[^）]*）', '', cell).replace('K-', '')
    out = []
    for part in re.split(r'[、,]', cell):
        part = part.strip()
        m = re.fullmatch(r'([A-Z]+)(\d+)\s*～\s*([A-Z]*)(\d+)', part)
        if m:
            width = len(m.group(2)) if m.group(2).startswith('0') else 0
            for n in range(int(m.group(2)), int(m.group(4)) + 1):
                out.append(m.group(1) + (str(n).zfill(width) if width else str(n)))
        elif re.fullmatch(r'[A-Z]+\d*', part):
            out.append(part)
    return out


registered = set()
for line in table.splitlines():
    m = re.match(r'\|\s*(K-[^|]+?)\s*\|', line)
    if m:
        registered.update(expand(m.group(1)))

src_text = {}
for f in CHECKERS:
    p = TOOLS / f
    if p.is_file():
        src_text[f] = p.read_text(encoding='utf-8')
alltext = '\n'.join(src_text.values())

# 只能在 sync 或别的脚本里实现的号：写明实现位置
ELSEWHERE = {'RULES': 'check_rules.py', 'L7': 'check_figs.py', 'B5': 'check_gaps.py', 'Z3': 'check_overshrink.py'}
for code in sorted(registered) if not COPIES_ONLY else ():
    if code in ELSEWHERE:
        if not (TOOLS / ELSEWHERE[code]).is_file():
            errs.append('K-%s 登记的实现 %s 不存在' % (code, ELSEWHERE[code]))
        continue
    if not re.search(r'(?<![A-Za-z0-9_-])%s(?![0-9])' % re.escape(code), alltext):
        errs.append('K-%s 在排版规则总表里登记了，检查脚本里找不到实现' % code)

emitted = set()
for f, t in src_text.items():
    emitted.update(re.findall(r"\b(?:err|sug|rep)\(\s*'([A-Z]+\d*)'", t))
    emitted.update(re.findall(r"(?:errors|warns)\.append\(\s*'([A-Z]+\d+) ", t))
for code in sorted(emitted - registered) if not COPIES_ONLY else ():
    errs.append('检查脚本报出的规则号 %s 没有登记在排版规则总表「检查器规则号」表' % code)

# ---- 3. 规则副本 ----
COPIES = [
    (PUB / 'prompt/portable-instruction.md', [NOTES / '规则/规则总纲.md', HB / '规则总纲.md']),
    (PUB / 'prompt/layout-checklist.md', [NOTES / '规则/排版规则总表.md', HB / '排版规则总表.md']),
    (PUB / 'prompt/content-core.md', [NOTES / '规则/内容校对核心规则.md', HB / '内容校对核心规则.md']),
    (PUB / 'prompt/standing-decisions.md', [NOTES / '规则/长期决策.md']),
    (PUB / 'prompt/lessons.md', [NOTES / '规则/经验与教训.md', HB / '经验与教训.md']),
    (PUB / 'prompt/SKILL.md', [NOTES / '规则/B737笔记总规范_单文件.md', HB / 'SKILL.md']),
]
for real, copies in (COPIES if not NO_COPIES else ()):
    body = real.read_bytes()
    for c in copies:
        if c.is_file() and c.read_bytes() != body:
            errs.append('规则副本与真本不一致：%s（真本 %s；sync --full 会自动复制）' % (c, real.relative_to(PUB)))
stale_copy = NOTES / '规则/排版检查清单.md'
if stale_copy.exists():
    errs.append('私有库还有旧副本 规则/排版检查清单.md（已改名 排版规则总表.md）')

for e in errs:
    print('RULES 错误', e)
print('规则自检：%d 个 K- 号、%d 个检查脚本，问题 %d' % (len(registered), len(src_text), len(errs)))
sys.exit(1 if errs else 0)
