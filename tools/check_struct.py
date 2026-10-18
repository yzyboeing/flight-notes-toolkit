#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_struct.py —— 表格结构与版面「源文件级」检查（SD-203，2026-10-17 用户：看速查 12 张截图后
「根据这一次发现的问题，再次完善内容整理和排版规则，然后根据新的规则审核和扫描。还有没有类似的情况？」）

以前漏掉的原因：check_layout 只量成品里「折行」，定宽表撑满页宽却没有折行、首列手动 <br>、表里混入别类行，都量不出来。
本脚本直接读源文件（两本），估算列宽需求，按 layout-checklist 新增的 T18～T24 报：
  T18 定宽表空白：定宽（w-NN）表各列估算所需宽度远小于给定宽度（表收窄或改自动宽度）
  T19 首列手动换行：首列格用 <br> 拆行，合起来不长、表有空位（合并成一行，括注接在后面）
  T20 两型表少数行只有一型：大部分行两型都有，个别行另一型为「—」（移出作注，注前标【机型】）
      另：大部分行只有一型，个别行只有另一型（显示组件失效表「上 DU 失效」）
  T21 两型数值表里的跨列共有行：数值对照表中夹着两型共有的限制 / 说明行（移出作注）
  T22 数值表里的判据 / 说明行：「项目 ｜ 数值」类表中，个别行的数值格是一整句判据 / 程序（移出作注）
  T23 表内注行有多条：<tr class="note"> 里有 ≥ 2 条（CAUTION / 注 / 说明混放），拆成表下注并分清性质
  T25 多列表取值列仍加点（SD-204）；T26 文字两型表只有一行差异（SD-204）
  T24 句末标点规则 A：punct_rule.py 高把握违例（错误）；存疑只计数
用法：python3 check_struct.py [--repo 笔记库] [--only T18,T20]   只读。退出码：有 T24 错误为 1。"""
import io, os, re, glob, sys, json
REPO = os.path.abspath(os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else '~/flight-repos/gh-private'))
ONLY = set(sys.argv[sys.argv.index('--only') + 1].split(',')) if '--only' in sys.argv else None
TOTAL = 16838 - 2 * 1440            # 横版版心（DXA），与 build_docx 同口径的近似
CJK_W, LAT_W, PAD, DOT = 180, 95, 300, 160   # 9pt：汉字 180、西文约 95；格内边距；圆点悬挂

def plain(s): return re.sub(r'<[^>]+>', '', s).replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').strip()
def tw(s):  # 估算一行文字宽（DXA）
    p = plain(s)
    return sum(CJK_W if ord(c) > 0x2000 else LAT_W for c in p)
def items(cell): return [x for x in re.split(r'<br\s*/?>', cell) if plain(x)]
def ttype(hdr): return [plain(h) for h in hdr]
OUT = []
import importlib.util as _iu
_sp = _iu.spec_from_file_location('fixw_fit', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fixw_fit.py')); FF = _iu.module_from_spec(_sp); _sp.loader.exec_module(FF)
def rep(rule, f, title, msg): OUT.append((rule, f, title, msg))

def tables(f):
    t = io.open(f, encoding='utf-8').read(); cur = ''
    pos = 0
    for m in re.finditer(r'(?m)^(#{3,4} [^\n]*)$|<table[^>]*>.*?</table>', t, re.S):
        if m.group(1): cur = m.group(1).lstrip('# ').strip(); continue
        tb = m.group(0)
        prev = t[:m.start()].rstrip().rsplit('\n', 1)[-1]
        yield cur, tb, prev

FILES = sorted(glob.glob(os.path.join(REPO, 'notes_src', '[1-5]*', '[0-9]*.md'))) + [os.path.join(REPO, '速查', '速查源.md')]
for f in FILES:
    fn = '速查' if f.endswith('速查源.md') else os.path.basename(f).split(' ')[0]
    for title, tb, prev in tables(f):
        hm = re.search(r'<tr class="hdr">(.*?)</tr>', tb, re.S)
        if not hm: continue
        ths = re.findall(r'<th([^>]*)>(.*?)</th>', hm.group(1), re.S)
        hdr = [plain(x[1]) for x in ths]; ncol = len(ths)
        rows = []
        for r in re.finditer(r'<tr(?![^>]*class="(?:hdr|premise)")([^>]*)>(.*?)</tr>', tb[hm.end():], re.S):
            cells = re.findall(r'<td([^>]*)>(.*?)</td>', r.group(2), re.S)
            rows.append((r.group(1), cells))
        body = [c for a, c in rows if 'note' not in a]
        figside = prev.startswith(('%%FIGSIDE%%', '%%FIGRIGHT%%'))
        # T18 定宽 / 图旁表列宽可优化：就是 fixw_fit 会改的表（检查与修复同口径）
        if re.search(r'\bw-\d+', hm.group(1)):
            pr = FF.propose(tb, prev)
            if pr: rep('T18', fn, title, '%s——python3 fixw_fit.py --apply' % pr['info'])
        # T19 首列手动换行
        for cells in body:
            if not cells: continue
            c0 = cells[0][1]
            row_tall = any(re.search(r'<br\s*/?>', c) or tw(c) > 0.5 * TOTAL for cc in body for _, c in cc[1:])   # SD-204：表里别的列有多行 / 长句，首列分行让宽度给内容列，不报
            if re.search(r'<br\s*/?>', c0) and not re.match(r'^\s*[①-⑳]', plain(c0)) and not row_tall:
                one = tw(re.sub(r'<br\s*/?>', '', c0)) + PAD
                if one <= 0.25 * TOTAL and ncol <= 3:
                    rep('T19', fn, title, '首列「%s」手动换行，合成一行约 %.0f%% 宽' % (plain(re.sub(r'<br\s*/?>', '｜', c0))[:30], 100 * one / TOTAL))
        # T20 两型表少数行只有一型
        if ncol >= 3 and re.search(r'737-NG', hdr[1] if ncol > 1 else '') and re.search(r'737-8', hdr[2] if ncol > 2 else ''):
            only8 = [cells for cells in body if len(cells) >= 3 and plain(cells[1][1]) == '—' and plain(cells[2][1]) not in ('—', '')]
            onlyN = [cells for cells in body if len(cells) >= 3 and plain(cells[2][1]) == '—' and plain(cells[1][1]) not in ('—', '')]
            n = len(body)
            for small, big, who in ((onlyN, only8, '737-NG'), (only8, onlyN, '737-8')):
                if '差异' in title: break                                   # 两型差异汇编表：一型条目本来就是主体
                if len(big) == 1 and len(small) == 1: continue               # 两边各一行（反推使用限制）：表本身是两型并列
                if 1 <= len(small) <= 2 and n >= 4 and (len(big) >= 2 or len(small) / n <= 0.25):
                    for cells in small:
                        if len(plain(cells[1 if who == '737-NG' else 2][1])) <= 20: continue   # 短数值 / 短语占不了多少空白，不动（2026-10-17 逐条审定）
                        rep('T20', fn, title, '「%s」行只有 %s（共 %d 行，另一型独有 %d 行）——移出作注「注：【%s】……」' % (plain(cells[0][1])[:20], who, n, len(big), who))
        # T21 两型数值表里的跨列共有行
        if ncol == 3 and re.search(r'737-NG', hdr[1]) and re.search(r'737-8', hdr[2]) and not re.search(r'no-share', ths[0][0]):
            shared = [cells for cells in body if any('colspan="2"' in a for a, c in cells)]
            numeric = [cells for cells in body if len(cells) == 3 and all(re.fullmatch(r'[\d.,\s%~～±≥≤<>（）()A-Za-z/：:在地面空中；]*', plain(c)) for a, c in cells[1:])]
            if shared and len(numeric) >= 2 and len(shared) <= len(body) / 2 and '差异' not in title:
                for cells in shared:
                    if len(plain(cells[-1][1])) <= 10: continue          # 共有的单个短数值（极限 N1 104%）不动
                    rep('T21', fn, title, '两型数值对照表中的共有行「%s」——移出作注' % plain(cells[0][1])[:24])
        # T22 数值表里的判据 / 说明行
        if ncol == 2 and re.search(r'数值|参数|值$', hdr[1] + hdr[0]):
            short = [cells for cells in body if len(cells) == 2 and len(plain(cells[1][1])) <= 16]
            for cells in body:
                if len(cells) == 2 and (re.search(r'[：:]\s*<br', cells[1][1]) or len(re.findall(r'[\u4e00-\u9fff]', plain(cells[1][1]))) >= 25) and len(short) >= 0.6 * len(body):   # 条件清单或长句才报（2026-10-17 逐条审定）
                    rep('T22', fn, title, '数值表中「%s」一行是判据 / 说明句——移出作注' % plain(cells[0][1])[:24])
        # T23 表内注行多条
        for a, cells in rows:
            if 'note' in a and cells and len(items(cells[0][1])) >= 2:
                kinds = sorted(set(re.findall(r'^(CAUTION|WARNING|注|说明)', plain(x)) and re.findall(r'^(CAUTION|WARNING|注|说明)', plain(x))[0] or '无标签' for x in items(cells[0][1])))
                rep('T23', fn, title, '表内注行有 %d 条（%s）——拆成表下注，CAUTION / 限制 / 说明分清' % (len(items(cells[0][1])), '、'.join(kinds)))
        # T25 多列表里的取值列仍加点（SD-204 C24）；T26 文字两型表只有一行差异（SD-204 C26）
        if ncol >= 3:
            for k in range(1, ncol):
                a_, h_ = ths[k]; hp = plain(h_)
                if 'col-bullet' not in a_ or re.search(r'(说明|内容|措施|处置|要点|逻辑|描述|现象|原因|作用|功能|备注|定义|含义|动作|程序|操作|职责|737)', hp): continue
                cs = [c for cells in body if len(cells) == ncol for a2, c in [cells[k]] if 'colspan' not in a2]
                its = [plain(x) for c in cs for x in re.split(r'<br\s*/?>', c) if plain(x)]
                if len(cs) >= 2 and its and all(len(x) <= 14 and re.search(r'[\d≥≤<>±]|^(好|中|差|无|有|是|否|亮|灭)', x) for x in its):
                    rep('T25', fn, title, '「%s」列是取值（%s…）却加点——整列居中不加点' % (hp, ' / '.join(its[:3])))
        if ncol == 3 and re.search(r'737-NG', hdr[1]) and re.search(r'737-8', hdr[2]) and '差异' not in title:
            diff = [c for c in body if len(c) == 3]; shared = [c for c in body if len(c) == 2 and 'colspan="2"' in c[1][0]]
            numeric = all(re.fullmatch(r'[\d.,\s%~～±≥≤<>（）()A-Za-z/：:]*', plain(c)) for r_ in diff for a2, c in r_[1:])
            if len(diff) <= 1 and len(shared) >= 2 and not numeric:
                rep('T26', fn, title, '两型表只有 %d 行有差异、%d 行共有——改「项目 ｜ 说明」，差异写注' % (len(diff), len(shared)))
# T24 句末标点
sys.argv_bak = sys.argv
try:
    import importlib.util
    sp = importlib.util.spec_from_file_location('punct_rule', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'punct_rule.py'))
    tmp = '/tmp/_punct_%d.json' % os.getpid()
    sys.argv = ['punct_rule.py', '--repo', REPO, '--report', tmp]
    import contextlib
    with contextlib.redirect_stdout(io.StringIO()):
        mod = importlib.util.module_from_spec(sp); sp.loader.exec_module(mod); mod.main()
    pr = json.load(open(tmp)); os.remove(tmp)
    for r in pr['改']: rep('T24', r[0].replace('.md', ''), r[1], '句末标点：「…%s」%s → %s' % (r[2][-20:], r[4], r[5]))
    N_DOUBT = len(pr['疑'])
finally:
    sys.argv = sys.argv_bak
if ONLY: OUT[:] = [x for x in OUT if x[0] in ONLY]
from collections import Counter
c = Counter(x[0] for x in OUT)
print('check_struct：%s；句末标点存疑 %d 处（不算错误）' % ('、'.join('%s %d' % kv for kv in sorted(c.items())) or '无', N_DOUBT))
for r, f, t, m in OUT: print('  [%s] %s「%s」：%s' % (r, f, t[:30], m))
sys.exit(1 if c.get('T24') else 0)
