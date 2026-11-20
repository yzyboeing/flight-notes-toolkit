#!/usr/bin/env python3
"""K-RULES 规则自检（2026-10-21 规则重组轮，方案 A1）。

用法：check_rules.py [笔记库路径] [--no-copies | --copies-only]   ——问题打印「RULES 错误 …」，退出码 1；全部通过退出码 0
（check_layout 调用时加 --no-copies：副本在 sync 收尾才复制，sync 复制后再跑 --copies-only）

查三件事：
1. 过时说法：现行规则与入口文件里不得再出现已作废的口径（rules.md、第 1～5 章、只本地提交绝不推送、块索引地址等）。
   写在「已作废」表里、或同一行带「作废 / 不再 / 已删 / 不设 / 原 / 旧」这类否定说明的不算。
2. K- 规则号与检查器一一对应：《排版规则总表》「检查器规则号」表里每个 K- 号都要在某个检查脚本里找得到；
   检查脚本 err() / sug() / rep() 报出的规则号都要登记在表里。
3. 规则副本：2026-11-20 规则瘦身后不再保留副本；私有库 规则/ 下不得再出现旧副本文件。
4. 现行规则（私有库 规则/现行规则.md）体量 ≤ 30,720 字节；决策记录（规则/决策记录.md）SD-267 起每条必须有「落实：」行，
   写明改了现行规则哪一节（「现行规则 第 X 节」）或哪个检查器（K- 号）。
"""
import os, re, sys
from pathlib import Path

PUB = Path(__file__).resolve().parents[1]
TOOLS = PUB / 'tools'
ARGS = [a for a in sys.argv[1:] if not a.startswith('--')]
NOTES = Path(ARGS[0]).resolve() if ARGS else PUB.parent / 'gh-private'
NO_COPIES, COPIES_ONLY = '--no-copies' in sys.argv, '--copies-only' in sys.argv
HB = Path.home() / 'Desktop/飞行理论笔记整理/Muse交接包'

CURRENT = ['SKILL.md', 'AGENTS.md', 'prompt/AGENT.md', 'prompt/checks.md']
PRIVATE = ['START_HERE.md', 'AGENTS.md', '规则/示例.md', '规则/现行规则.md']
RULES_FILE, DECISIONS = '规则/现行规则.md', '规则/决策记录.md'
RULES_LIMIT = 30720

# 过时说法：(正则, 说明)
STALE = [
    (r'部分机号|一般机号', 'SD-247（2026-11-05）：不用「机号」一词，写「部分 737-NG 飞机」「部分 737-8 飞机」'),
    (r'prompt/rules\.md|整理规则（真源）', '旧整理规则 rules.md 已归档，现行规则是总纲 / 排版规则总表 / 内容校对核心规则'),
    (r'5 章 76 节|第二章\s*机组训练手册|2 机组训练手册|第五章\s*模拟机训练|5 模拟机训练', 'SD-234（2026-10-24）：新框架 5 章 + 附录 72 节，第二章「操作程序」、第五章「模拟机笔记」'),
    (r'4 章 60 节|第四章\s*模拟机训练|4 模拟机训练|第五章\s*技术提示|5 技术提示', '2026-10-23 两本整合：完整版 5 章 76 节——第四章 飞行技术提示、第五章 模拟机训练'),
    (r'第三章\s*运行手册|3 运行手册|第四章\s*飞行技术提示|4 飞行技术提示|第五章\s*模拟机笔记|5 模拟机笔记', '2026-10-25 用户：章名改为 第三章 运行规范 / 第四章 技术提示 / 第五章 故障处置'),
    (r'第四章\s*技术提示|4 技术提示 \||06 技术提示', 'SD-241（2026-10-27）：第四章改名「特殊情况与风险」（7 节）；速查目录 06 段同名'),
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
    # SD-242（2026-10-29）：唯一内容真源 notes_src，速查源只读溯源
    (r'两本源文件同改|改完整版要同步速查|改共有内容两本同改|共有表以速查为准|速查源\.md`?\s*(是|与 notes_src 都是)?唯一内容真源|与 `速查/速查源\.md` 是唯一内容真源', 'SD-242：只改 notes_src，速查源只读溯源、不再两本同改'),
    # 2026-11-05（1.5.0）：App 为主产品、用户本人两型都飞、新条按主题插入
    (r'会给飞行学员学习|给飞行学员学习', '2026-11-04 用户：没有学员，用户本人两型都飞（另有 TestFlight 测试者）'),
    (r'模拟机 / QRH → 第 4 章|第 4 章顺序锁定', 'SD-241 后模拟机 / QRH 在第 5 章；节内按主题插入，条号后移由 App 迁移链处理'),
    (r'「最新文件」只给速查 PDF|最新文件」只给速查', '2026-11 起日常交付是 App，PDF / Word 用户点名才出'),
    (r'或在节末追加|追加在节末', '新条按主题插在最贴切的条目之后，迁移链负责收藏（2026-11-05 审核）'),
    # 2026-11-19 规则矛盾审核（问题 5）：SD-254～264 作废的说法，现行规则里再出现就报
    (r'能用图表达得更简洁的，不用表格|能用图片和思维导图表达更简洁的就不要用表格', '第十轮 11-18 修订（App 规则 〇-3 / P10）：按信息关系选表 / 图，比较类优先用表'),
    (r'只留第六章|只放第六章|两型差异汇总只放第六章|位置、标签、外观类只留第六章', 'SD-258（11-19）：取消第六章，两型差异完全融合进前五章'),
    (r'6\.2 镜像|与 6\.2 逐字一致|与 6\.1 逐字一致|逐字镜像的正式差异汇编表|与其逐字镜像', 'SD-258：第六章取消，K-S22 / K-S23 改为融合后的检查'),
    (r'〈系统名?〉两型差异」条里|K-S23 要求', 'SD-255 / 258：第一章不设节末两型差异条，K-S23 原逐字校验取消'),
    (r'全是数值时整表居中|两型列全是数值\s*/?\s*短值\s*→|两型列\*?\*?全是数值\s*/\s*短值\*?\*?的，两列居中', 'SD-255（11-19）：两型列一律靠左加点，数值也一样，共有格居中'),
    (r'竖屏上下排|竖屏仍显示上下排|照旧上下排', 'SD-255（11-19）：两型对比图横屏、竖屏都并排，字小点图放大'),
    (r'横屏并排版.{0,20}最小字\s*≥\s*11px|看不清不并排', 'SD-255：并排两型对比图豁免 11px 下限，靠点图放大'),
    (r'同表已有加点列时.{0,12}也(靠左)?加点|同一张表里已有加点列时，长句单项列也', 'SD-262（11-19）：单项列一律居中不加点，不因同表别的列加点而跟着加点'),
    (r'第五章\s*5\.3\s*增压相关|5\.3 增压相关|5\.9 燃油、液压与起落架', 'SD-259（11-19）：第五章按 QRH 章分节，一章一节'),
    (r'两型不同处：……」|图底留一行「两型不同处|图下一行「两型不同处」', '11-14 用户：两型不同处不写在图里，进表或注；面板图差异写图底红框图例'),
    (r'一步一框', 'SD-263（11-19）：简化处置导图，一阶段一卡，不需要的不画'),
    (r'第 ?N ?步对应 737-NG 第 ?M ?步」?(写|照写|保留)', 'SD-264（11-19）：不写纯步号对应'),
    (r'OneNote 笔记本同步全书|OneNote（统一同步）|OneNote 只在用户指令时更新|统一同步 OneNote|确认写入 OneNote', 'SD-266（11-19）：OneNote 彻底停更，iPad App 为唯一主产物'),
]
EXEMPT = re.compile(r'作废|不再|已删|不设|已并入|原第|旧|删除|去掉|取消|改名|不收|无前言|不标|没有|已归档|取代')   # 11-19 加「取代」：「（取代原……）」是作废说明

CHECKERS = ['check_layout.py', 'check_struct.py', 'check_src.py', 'check_blocks.py', 'check_quickref.py',
            'check_pair.py', 'check_userpins.py', 'check_linkage.py', 'check_figs.py', 'check_rules.py',
            'verify.py', 'check_splits.py', 'check_coverage.py', 'punct_rule.py', 'fit_fix.py',
            'check_model_symmetry.py']   # K-S37（11-19，原自称 K-S23）

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
lc = (PUB / 'prompt/checks.md').read_text(encoding='utf-8')
sec = lc.split('## 检查器规则号（K-）', 1)
if len(sec) < 2:
    errs.append('prompt/checks.md 找不到「## 检查器规则号（K-）」一节')
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
# 11-19：check_src 的 K-S22 / K-S23（第六章融合后的检查）与 check_model_symmetry 的 K-S37 只以「K-Sxx」形式出现，按实现文件登记
ELSEWHERE.update({'S22': 'check_src.py', 'S23': 'check_src.py', 'S37': 'check_model_symmetry.py'})
# 外部检查（实现在别的仓库，SD-244）：学习系统网页排版检查 AC1～AC10 = study/tools/test_layout.js 的 C1～C10；本机没有 study 库（如噜噜虚拟机）时只提示不报错
EXTERNAL = {'AC%d' % n: PUB.parent / 'study' / 'tools' / 'test_layout.js' for n in range(1, 11)}
EXTERNAL['L9'] = PUB.parent / 'gh-private' / 'figures_src' / 'qrh' / 'build_figs.py'   # 11-12：第五章旁注图最大字号 ≤ 15px（比正文小一号），出图时由 build_figs 检查
EXTERNAL['L10'] = PUB.parent / 'gh-private' / 'figures_src' / 'check_overflow.py'   # 11-12：自绘图文字溢出（Chrome 实测）
for code in sorted(registered) if not COPIES_ONLY else ():
    if code in EXTERNAL:
        if not EXTERNAL[code].is_file() and code == 'AC1':
            print('RULES 提示 外部检查 %s 不在本机（K-AC1～AC10 由学习系统侧执行）' % EXTERNAL[code])
        continue
    if code in ELSEWHERE:
        if not (TOOLS / ELSEWHERE[code]).is_file():
            errs.append('K-%s 登记的实现 %s 不存在' % (code, ELSEWHERE[code]))
        continue
    if not re.search(r'(?<![A-Za-z0-9_-])%s(?![0-9])' % re.escape(code), alltext):
        errs.append('K-%s 在 prompt/checks.md 里登记了，检查脚本里找不到实现' % code)

emitted = set()
for f, t in src_text.items():
    emitted.update(re.findall(r"\b(?:err|sug|rep)\(\s*'([A-Z]+\d*)'", t))
    emitted.update(re.findall(r"(?:errors|warns)\.append\(\s*'([A-Z]+\d+) ", t))
for code in sorted(emitted - registered) if not COPIES_ONLY else ():
    errs.append('检查脚本报出的规则号 %s 没有登记在 prompt/checks.md「检查器规则号」表' % code)

# ---- 3. 规则副本（2026-11-20 起不保留副本）----
for old in ('规则总纲.md', '排版规则总表.md', '内容校对核心规则.md', '长期决策.md', '经验与教训.md', 'B737笔记总规范_单文件.md', '排版检查清单.md'):
    if (NOTES / '规则' / old).exists() and not COPIES_ONLY:
        errs.append('私有库 规则/%s 是瘦身前的旧副本，应移到 规则/历史档_20261120/' % old)

# ---- 4. 现行规则体量与决策记录落实行 ----
if not COPIES_ONLY:
    rules = NOTES / RULES_FILE
    if not rules.is_file():
        errs.append('找不到现行规则 %s' % rules)
    elif rules.stat().st_size > RULES_LIMIT:
        errs.append('现行规则 %d 字节，超过上限 %d（先删过时内容，来历写进决策记录）' % (rules.stat().st_size, RULES_LIMIT))
    dec = NOTES / DECISIONS
    if not dec.is_file():
        errs.append('找不到决策记录 %s' % dec)
    else:
        parts = re.split(r'(?m)^## (SD-\d+)', dec.read_text(encoding='utf-8'))
        seen = []
        for code, body in zip(parts[1::2], parts[2::2]):
            n = int(code[3:])
            seen.append(n)
            if n < 267:
                errs.append('决策记录 %s：SD-1～266 已冻结在历史档，新条从 SD-267 起' % code)
            m = re.search(r'(?m)^\s*(?:[-*]\s*)?落实[：:](.*)$', body)
            if not m or not re.search(r'现行规则\s*第\s*[〇一二三四五六七八九十\d.]+\s*节|K-?[A-Z]*\d+', m.group(1)):
                errs.append('决策记录 %s 缺「落实：」行，或没写改了现行规则哪一节 / 哪个 K- 号（现行规则 第七节）' % code)
        if seen and seen != list(range(267, 267 + len(seen))):
            errs.append('决策记录 SD 号不连续：%s' % seen)

for e in errs:
    print('RULES 错误', e)
print('规则自检：%d 个 K- 号、%d 个检查脚本，问题 %d' % (len(registered), len(src_text), len(errs)))
sys.exit(1 if errs else 0)
