#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""qr_directory.py —— 《机型理论速查笔记》的「速查目录」数据（SD-234，2026-10-24；取代 SD-233 按速查主题生成的版本）
用法：python3 qr_directory.py --repo <gh-private> [--out build/速查目录.json]
供 build_docx.js 的 bodyTopicPage（DOC_TOPICS=本文件、DOC_TOPICS_FRONT=1）排在总目录之后。

规则（新笔记框架定稿，用户 2026-10-24「框架定稿，可以开始重建」）：
  · 8 段：01 系统描述 / 02 滑行与起飞 / 03 爬升与巡航 / 04 下降、进近与着陆 / 05 运行规范 / 06 特殊情况与风险 /
    07 故障处置 / 08 机型差异；直接项就是对应章节的条目（每段的主题见 DIMS），全书每条正文直接出现恰好一次；
  · ↗（灰字）交叉按 XREF 列出：保住按飞行阶段、放行、应急等浏览路径（顾问评估 A 类第 1、2 条）；
  · 目录项写「x.y N|标题|x|小项|」，不带要点行（用户 2026-10-24）；
  · 小项取正文第一张表的首列标签（≥ 3 个）；首列是序号（①②③ / 数字）时取第二列（顾问评估 A 类第 3 条）；
  · 生成后自检：全部条目恰好直接一次、XREF 引用的条目与主题都存在，否则报错退出。"""
import os, re, sys, json, glob
REPO = os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1]) if '--repo' in sys.argv else os.path.expanduser('~/flight-repos/gh-private')
OUT = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else os.path.join(REPO, 'build', '速查目录.json')

# 完整版条目
full, body, order, secs = {}, {}, [], {}
for f in glob.glob(os.path.join(REPO, 'notes_src', '[1-6] *', '[0-9]*.md')):
    t = open(f, encoding='utf-8').read(); sid = re.search(r'(?m)^id: "([\d.]+)"', t).group(1)
    secs[sid] = re.search(r'(?m)^title_cn: "(.*)"', t).group(1)
    for m in re.finditer(r'(?ms)^### (\d+)\. (.*?)\n(.*?)(?=^### |\Z)', t):
        k = '%s %s' % (sid, m.group(1)); full[k] = re.sub(r'\s*\{★+\}\s*$', '', m.group(2)).strip(); body[k] = m.group(3); order.append(k)
skey = lambda k: [int(x) for x in k.split()[0].split('.')] + [int(k.split()[1])]
order.sort(key=skey)
clean = lambda t: re.sub(r'（单位：[^）]*）|【[^】]*】', '', t).strip()
ORD = re.compile(r'^[①-⑳\d]')
CMP = re.compile(r'^[<>≤≥＜＞]')
CALL = re.compile(r"^[A-Z][A-Z0-9 ,.'’!/-]*$")   # 英文喊话（TCAS 音响等）
DASH = {'—', '–', '-', '——'}
wid = lambda t: sum(0.5 if ord(c) < 128 else 1 for c in t)   # 显示宽度：英文数字算半字


def subs_of(k):
    m = re.search(r'<table[^>]*>(.*?)</table>', body[k], re.S)
    if not m: return []
    rows = re.findall(r'<tr(?![^>]*class="(?:hdr|premise|note)")[^>]*>(.*?)</tr>', m.group(1), re.S)
    def cell(row, i):
        cs = re.findall(r'<td[^>]*>(.*?)</td>', row, re.S)
        if len(cs) <= i: return ''
        t = re.split(r'<br\s*/?>|（|\(', cs[i])[0]
        t = re.sub(r'<[^>]+>', '', t).replace('&gt;', '>').replace('&lt;', '<').strip().rstrip('：:。；;，,').strip()
        return re.sub(r'\s*【737[^】]*】\s*', '', t).strip()
    firsts = [cell(r, 0) for r in rows]
    col = 1 if firsts and sum(1 for x in firsts if ORD.match(x)) * 2 >= len(firsts) else 0   # 首列是序号 → 取第二列
    # 2026-10-29 N3（评估会话定）：「—」续行格跳过；比较式（> 13000ft、< 10s）或整句（> 14 字）占一半以上的表整表不取；英文喊话不算长
    cand = [t for t in (cell(r, col) for r in rows) if t and t not in DASH and not ORD.match(t)]
    if cand and sum(1 for t in cand if CMP.match(t)) * 2 >= len(cand): return []
    if cand and sum(1 for t in cand if wid(t) > 14 and not CALL.match(t)) * 2 >= len(cand): return []
    out = []
    for t in cand:
        if CMP.match(t) or len(t) > 16 or t in out: continue
        out.append(t)
    return out if len(out) >= 3 else []


def ks(spec):
    """'1.7 15' / '3.3 24-27' / '5.3 *' → 条目键列表"""
    s, n = spec.split()
    if n == '*': return [k for k in order if k.split()[0] == s]
    a, _, b = n.partition('-')
    return ['%s %d' % (s, i) for i in range(int(a), int(b or a) + 1)]


S = lambda *ids: [{'theme': secs[i], 'secs': [i]} for i in ids]
DIMS = [   # SD-241（2026-10-27）结构重编：第二章 10 节、第三章 14 节、第四章 7 节「特殊情况与风险」、第五章 14 节
    ('01 系统描述', S(*['1.%d' % i for i in range(1, 16)])),
    ('02 滑行与起飞', [{'theme': '地面准备', 'secs': []}] + S('2.1', '2.2', '2.3', '2.4')),
    ('03 爬升与巡航', S('2.5')),
    ('04 下降、进近与着陆', S('2.6', '2.7', '2.8', '2.9', '2.10')),
    ('05 运行规范', S(*['3.%d' % i for i in range(1, 15)]) + [{'theme': '放行与签派', 'secs': []}]),
    ('06 特殊情况与风险', S(*['4.%d' % i for i in range(1, 8)])),
    ('07 故障处置', [{'theme': '处置原则', 'secs': ['5.1']}, {'theme': '各飞行阶段', 'secs': ['5.2']},
                        {'theme': '系统故障', 'secs': ['5.%d' % i for i in range(3, 10)]}, {'theme': '目视进近与单发', 'secs': ['5.10', '5.11']},
                        {'theme': '记忆项目与机动飞行', 'secs': ['5.12', '5.13']}, {'theme': '撤离与水上迫降', 'secs': ['5.14']}]),
    ('08 机型差异', S('6.1', '6.2', '6.3', '6.4') + [{'theme': '只涉及一型的知识点', 'secs': []}]),
]
# ↗ 交叉（主题 → 条目），按列出顺序排在本主题直接项之后；主题名就是节名（或 07 段分组名）
XREF = {
    '引气、空调与增压': ['6.1 1', '5.3 *'], '防冰、排雨': ['6.1 2', '3.5 1-5'], '自动飞行': ['6.1 3'],
    '电气': ['6.1 4', '5.4 *'], '发动机、APU': ['6.1 5-6', '5.5 *'], '防火': ['6.1 7', '5.6 *'], '飞行操纵': ['6.1 8', '5.7 *'],
    '飞行仪表、显示': ['6.1 9', '5.8 *'], '飞行管理、导航': ['6.1 9', '5.8 *'], '燃油': ['6.1 10', '5.9 1-4'], '液压': ['6.1 10', '5.9 5-6'],
    '起落架': ['6.1 11', '5.9 7-11'], '警告系统': ['6.1 12', '4.3 1', '4.3 3'],
    '地面准备': ['1.7 15-16', '1.7 18-21', '3.4 13-16', '3.14 2-3', '5.2 1-5', '6.4 1'],
    '起飞技术与中断起飞': ['1.14 9', '3.4 17-19', '3.7 1', '4.3 2', '6.3 1-2', '5.2 6-12', '6.4 2'],
    '爬升与巡航': ['2.9 4', '4.1 3', '4.2 *', '4.5 *', '3.9 6-8', '3.2 3-9'],
    '下降与进近': ['2.2 6-7', '1.11 19-24', '3.5 6-9', '3.9 1-4', '3.6 3-5', '4.1 4', '5.10 *', '6.4 3'],
    '复飞与中断着陆': ['2.5 3', '1.11 25', '4.1 5', '5.11 *'],
    '着陆与滑跑': ['3.7 1', '3.7 3-4', '5.9 11', '3.14 3'],
    '航空安保与机组通则': ['2.1 1-5'],
    '夏季与高原运行': ['1.2 24'], '平行跑道、RVSM 与安全高度': ['1.10 2'], '紧急遇险、快速离机与就近合适机场': ['5.1 8', '5.14 1'],   # 1024R1 内容逻辑评估（顾问采纳）
    '放行与签派': ['2.1 6', '3.12 1-2', '3.13 1', '3.8 1-3', '3.6 6-7', '2.8 3', '5.8 5', '1.1 6'],
    '颠簸': ['3.9 5'], '气象报文': ['4.2 4'],   # 2026-10-25：3.1 第 7 条（垂直风切变值 S）移作 4.2 第 4 条（用户批注）
    '低能见运行与备降标准': ['新机长附加要求=3.6 3'],   # 2026-10-25 用户：「速查目录里面加一条新机长附加要求的知识点」（3.4 第 3 条表下注）
    '风切变': ['1.15 16', '5.13 4-8'], 'GNSS 信号干扰': ['1.11 3'],
    '处置原则': ['3.10 1', '3.10 3-4', '3.9 8', '6.2 1', '6.2 17'], '各飞行阶段': ['4.3 *'], '系统故障': ['1.1 12', '1.2 28', '4.7 1'],
    '撤离与水上迫降': ['3.10 2'],
    '只涉及一型的知识点': ['%s %s' % tuple(k.split()) for k in order if k.split()[0][0] in '12345' and re.search(r'【737', full[k])],
}

direct, dims, miss = {}, [], []
for dname, themes in DIMS:
    out = []
    for tp in themes:
        items, have = [], set()
        for s in tp['secs']:
            for k in [k for k in order if k.split()[0] == s]:
                assert k not in direct, '重复直接项 ' + k
                direct[k] = (dname, tp['theme']); have.add(k)
                items.append('%s|%s||%s|' % (k, clean(full[k]), '·'.join(subs_of(k))))
        for spec in XREF.get(tp['theme'], []):
            if '=' in spec:   # 「名称=条目号」：条目里的某个知识点（注等）单列一行，名称代替条目标题，指向所在条目
                lab, ref = spec.split('=', 1)
                if ref not in full: miss.append('%s → %s' % (tp['theme'], ref)); continue
                items.append('%s|%s|x||' % (ref, lab)); continue
            for k in ks(spec):
                if k not in full: miss.append('%s → %s' % (tp['theme'], k)); continue
                if k in have: continue
                have.add(k); items.append('%s|%s|x||' % (k, clean(full[k])))
        if items: out.append({'theme': tp['theme'], 'items': items})
    if len(out) == 1: out[0]['nohead'] = True
    dims.append({'dim': dname, 'themes': out})
unknown = [t for t in XREF if not any(t == tp['theme'] for _, th in DIMS for tp in th)]
nodirect = [k for k in order if k not in direct]
if miss or unknown or nodirect:
    raise SystemExit('速查目录自检失败：XREF 找不到条目 %s；XREF 主题不存在 %s；没有直接项 %s' % (miss[:10], unknown, nodirect[:20]))
os.makedirs(os.path.dirname(OUT), exist_ok=True)
json.dump(dims, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
nx = sum(1 for d in dims for t in d['themes'] for x in t['items'] if x.split('|')[2] == 'x')
print('速查目录：%d 段 %d 主题；直接项 %d 条（= 全书 %d 条）、↗ %d 项 → %s' % (len(dims), sum(len(d['themes']) for d in dims), len(direct), len(order), nx, OUT))

# 条目级「相关知识点」（2026-10-29 N4，评估会话定）：真源是私有库根目录《相关知识点源.json》（人工判定），
#   校验后与速查目录同目录输出 相关知识点.json（学习系统快照随 build/ 带上）。有向：A 列 B = 读 A 时该去看 B，消费方不得对称化；
#   同节相邻不列、每条 ≤ 5、why ≤ 12 字；改条号后键 / 目标找不到即报错退出（改号清单要带上这份源文件）。
RSRC = os.path.join(REPO, '相关知识点源.json')
if os.path.exists(RSRC):
    import subprocess
    rel, bad = json.load(open(RSRC, encoding='utf-8'))['rel'], []
    for a, lst in rel.items():
        if a not in full: bad.append('键不存在 ' + a)
        if len(lst) > 5: bad.append('%s 超过 5 条' % a)
        for x in lst:
            k = x['k']
            if k not in full: bad.append('%s → 目标不存在 %s' % (a, k))
            elif a in full and k.split()[0] == a.split()[0] and abs(int(k.split()[1]) - int(a.split()[1])) <= 1: bad.append('%s → 同节相邻 %s' % (a, k))
            if len(x.get('why', '')) > 12: bad.append('%s → %s why 超 12 字' % (a, k))
    if bad: raise SystemExit('相关知识点自检失败：%s' % '；'.join(bad[:10]))
    tag = subprocess.run(['git', '-C', REPO, 'tag', '--points-at', 'HEAD', '--list', 'baseline/*'], capture_output=True, text=True).stdout.split()
    if not tag:
        last = subprocess.run(['git', '-C', REPO, 'describe', '--tags', '--abbrev=0', '--match', 'baseline/*'], capture_output=True, text=True).stdout.strip()
        tag = [last + '+'] if last else ['']
    ROUT = os.path.join(os.path.dirname(OUT), '相关知识点.json')
    json.dump({'baseline': os.environ.get('NOTES_BASELINE') or tag[-1].replace('baseline/', ''),
               '_说明': '有向：A 列 B = 读 A 时该去看 B；消费方不得对称化（不把 B → A 自动补上）。键与 k 为全书条目号「x.y N」；why ≤ 12 字；只列真相关，每条 ≤ 5；同节相邻不列。baseline 带「+」表示在该基线之后、尚未出新基线。',
               'rel': {a: lst for a, lst in rel.items() if lst}}, open(ROUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('相关知识点：%d 条有列表、%d 个关系 → %s' % (sum(1 for v in rel.values() if v), sum(len(v) for v in rel.values()), ROUT))
