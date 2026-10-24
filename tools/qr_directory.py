#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""qr_directory.py —— 《机型理论速查笔记》的「速查目录」数据（SD-234，2026-10-24；取代 SD-233 按速查主题生成的版本）
用法：python3 qr_directory.py --repo <gh-private> [--out build/速查目录.json]
供 build_docx.js 的 bodyTopicPage（DOC_TOPICS=本文件、DOC_TOPICS_FRONT=1）排在总目录之后。

规则（新笔记框架定稿，用户 2026-10-24「框架定稿，可以开始重建」）：
  · 8 段：01 系统描述 / 02 滑行与起飞 / 03 爬升与巡航 / 04 下降、进近与着陆 / 05 运行手册 / 06 飞行技术提示 /
    07 模拟机故障与处置 / 08 机型差异；直接项就是对应章节的条目（每段的主题见 DIMS），全书每条正文直接出现恰好一次；
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
    out = []
    for r in rows:
        t = cell(r, col)
        if not t or ORD.match(t) or len(t) > 16 or t in out: continue
        out.append(t)
    return out if len(out) >= 3 else []


def ks(spec):
    """'1.7 15' / '3.3 24-27' / '5.3 *' → 条目键列表"""
    s, n = spec.split()
    if n == '*': return [k for k in order if k.split()[0] == s]
    a, _, b = n.partition('-')
    return ['%s %d' % (s, i) for i in range(int(a), int(b or a) + 1)]


S = lambda *ids: [{'theme': secs[i], 'secs': [i]} for i in ids]
DIMS = [
    ('01 系统描述', S(*['1.%d' % i for i in range(1, 16)])),
    ('02 滑行与起飞', S('2.1', '2.2')),
    ('03 爬升与巡航', S('2.3')),
    ('04 下降、进近与着陆', S('2.4', '2.5', '2.6', '2.7', '2.8')),
    ('05 运行手册', S(*['3.%d' % i for i in range(1, 13)]) + [{'theme': '放行与签派', 'secs': []}]),
    ('06 飞行技术提示', S(*['4.%d' % i for i in range(1, 11)])),
    ('07 模拟机故障与处置', [{'theme': '处置原则', 'secs': ['5.1', '5.2']}, {'theme': '各飞行阶段', 'secs': ['5.3', '5.4', '5.5', '5.6']},
                        {'theme': '系统故障', 'secs': ['5.%d' % i for i in range(7, 16)]}, {'theme': '目视进近与单发', 'secs': ['5.16', '5.17', '5.18']},
                        {'theme': '记忆项目与机动飞行', 'secs': ['5.19', '5.20']}, {'theme': '撤离与水上迫降', 'secs': ['5.21', '5.22']}]),
    ('08 机型差异', S('6.1', '6.2', '6.3', '6.4', '6.5') + [{'theme': '只涉及一型的知识点', 'secs': []}]),
]
# ↗ 交叉（主题 → 条目），按列出顺序排在本主题直接项之后；主题名就是节名（或 07 段分组名）
XREF = {
    '引气、空调与增压': ['6.1 1', '4.6 6', '5.7 *'], '防冰、排雨': ['6.1 2', '3.3 31-35'], '自动飞行': ['6.1 3'],
    '电气': ['6.1 4', '5.8 *'], '发动机、APU': ['6.1 5-6', '5.9 *'], '防火': ['6.1 7', '5.10 *'], '飞行操纵': ['6.1 8', '5.11 *'],
    '飞行仪表、显示': ['6.1 9', '5.12 *'], '飞行管理、导航': ['6.1 9', '5.12 *'], '燃油': ['6.1 10', '5.13 *'], '液压': ['6.1 10', '5.14 *'],
    '起落架': ['6.1 11', '5.15 *'], '警告系统': ['6.1 12', '4.3 1', '4.3 3'],
    '地面准备': ['1.7 15-16', '1.7 18-21', '3.3 24-27', '3.12 2-3', '5.3 *', '6.5 1'],
    '起飞技术与中断起飞': ['1.14 10', '3.3 28-30', '3.5 1', '4.3 2', '6.3 *', '5.4 *', '5.5 *', '6.5 2'],
    '爬升与巡航': ['2.7 4', '4.1 3', '4.2 *', '4.5 *', '3.7 6-8', '3.2 3-9'],
    '下降与进近': ['2.1 12-13', '1.11 19-24', '3.3 36-39', '3.7 1-4', '3.4 3-5', '4.1 4', '5.16 *', '5.17 *', '6.5 3'],
    '复飞与中断着陆': ['2.3 3', '1.11 25', '4.1 5', '5.18 *'],
    '着陆与滑跑': ['3.5 1', '3.5 3-5', '4.9 *', '3.12 3'],
    '航空安保、机长权利、核心胜任力与 TEM': ['2.1 1-5'],
    '夏季与高原运行': ['1.2 24'], '平行跑道、RVSM 与安全高度': ['1.10 2'], '紧急遇险、快速离机与就近合适机场': ['5.2 7', '5.21 *'],   # 1024R1 内容逻辑评估（顾问采纳）
    '放行与签派': ['2.1 6', '3.10 1-2', '3.11 1', '3.6 1-3', '3.4 6-7', '2.6 4', '4.8 2', '1.1 6'],
    '颠簸': ['3.1 7', '3.7 5'], '风切变': ['1.15 17', '5.6 *', '5.20 4'], '座舱增压': ['1.2 28', '5.7 *'], 'GNSS 信号干扰': ['1.11 3'],
    '处置原则': ['3.8 1', '3.8 3-4', '3.7 8', '6.2 1', '6.2 17'], '各飞行阶段': ['4.3 *'], '系统故障': ['4.6 6', '1.1 12', '4.10 1'],
    '撤离与水上迫降': ['3.8 2'],
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
