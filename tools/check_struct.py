#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_struct.py —— 表格结构与版面「源文件级」检查（SD-203，2026-10-17 用户：看速查 12 张截图后
「根据这一次发现的问题，再次完善内容整理和排版规则，然后根据新的规则审核和扫描。还有没有类似的情况？」）

以前漏掉的原因：check_layout 只量成品里「折行」，定宽表撑满页宽却没有折行、首列手动 <br>、表里混入别类行，都量不出来。
本脚本直接读源文件（两本），估算列宽需求，按 layout-checklist 新增的 T18～T24 报：
  T18 定宽表空白：定宽（w-NN）表各列估算所需宽度远小于给定宽度（表收窄或改自动宽度）
  T19 首列手动换行不省列宽（SD-207 改，错误）：术语格、括注另行除外；合成一行不超过本列其他格最宽一行的 → 合成一行\n  T29 首列括注另行（C29，错误）；T30 首列被单个长格撑宽（错误）；T31 注内两组以上「A＝B」并列判读（错误，SD-208）
  T20 两型表少数行只有一型：大部分行两型都有，个别行另一型为「—」（移出作注，注前标【机型】）
      另：大部分行只有一型，个别行只有另一型（显示组件失效表「上 DU 失效」）
  T21 两型数值表里的跨列共有行：数值对照表中夹着两型共有的限制 / 说明行（移出作注）
  T22 数值表里的判据 / 说明行：「项目 ｜ 数值」类表中，个别行的数值格是一整句判据 / 程序（移出作注）
  T23 表内注行有多条：<tr class="note"> 里有 ≥ 2 条（CAUTION / 注 / 说明混放），拆成表下注并分清性质
  T28 术语格两行 / 一行（SD-206）
  T27 标注与成品不符（SD-205，读成品 docx：显式居中列没居中、首列出现圆点）
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
_GRIDS = {}   # (本, 表头首格文字, 列数) → 成品列宽（P06 用）
def _docx_grid(path):
    import zipfile, html as _h
    if not os.path.exists(path): return []
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8'); out = []
    for m in re.finditer(r'<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>', x, re.S):
        tb = m.group(0); grid = [int(v) for v in re.findall(r'<w:gridCol w:w="(\d+)"', tb)]
        hdr0 = ''; c0 = []; hdrall = ''
        for ri, r in enumerate(re.findall(r'<w:tr[ >].*?</w:tr>', tb, re.S)):
            tcs = re.findall(r'<w:tc>.*?</w:tc>', r, re.S)
            if not tcs: continue
            c = tcs[0]
            if '<w:gridSpan' in c or ('vMerge' in c and 'restart' not in c): continue
            ps = [_h.unescape(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))) for p in re.findall(r'<w:p[ >].*?</w:p>', c, re.S)]
            ps = [p.strip() for p in ps if p.strip()]
            if ri == 0:
                hdr0 = ''.join(ps); hdrall = re.sub(r'\s', '', _h.unescape(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r)))); continue
            if ps:
                rowp = [len([p for p in re.findall(r'<w:p[ >].*?</w:p>', cc, re.S) if re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)]) for cc in tcs]
                c0.append((ps, max(rowp)))
        out.append((grid, hdr0, c0, hdrall))
    return out
for _bk0, _p0 in (('速查', 'B737机型理论速查.docx'), ('全书', 'B737机型理论速查笔记.docx')):
    for _g, _h0, _c0, _ha in _docx_grid(os.path.join(REPO, 'build', _p0)):
        _GRIDS.setdefault((_bk0, _ha, len(_g)), _g)

def tables(f):
    t = io.open(f, encoding='utf-8').read(); cur = ''
    pos = 0
    for m in re.finditer(r'(?m)^(#{3,4} [^\n]*)$|<table[^>]*>.*?</table>', t, re.S):
        if m.group(1): cur = m.group(1).lstrip('# ').strip(); continue
        tb = m.group(0)
        prev = t[:m.start()].rstrip().rsplit('\n', 1)[-1]
        yield cur, tb, prev

FILES = sorted(glob.glob(os.path.join(REPO, 'notes_src', '[1-6]*', '[0-9]*.md'))) + [os.path.join(REPO, '速查', '速查源.md')]
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
        if re.search(r'\bw-\d+', hm.group(1)) and not re.search(r'w-lock|keep-wide|col-eq', hm.group(1)):   # SD-209：实测锁定 / 用户点名满宽 / 等宽的表不再按估算报
            pr = FF.propose(tb, prev)
            if pr: rep('T18', fn, title, '%s——python3 fixw_fit.py --apply' % pr['info'])
        elif figside and not re.search(r'\bw-\d+', hm.group(1)) and '定' not in [x.strip() for x in prev.split('|')[3:]]:   # P06（SD-209）：没有定宽的图旁表按成品实测列宽作起始比例，只在能收图时报（标「定」的不报）
            g = _GRIDS.get(('速查' if '速查' in f else '全书', re.sub(r'\s', '', ''.join(hdr)), ncol))
            pr = FF.propose(tb, prev, pct0=[100.0 * x / sum(g) for x in g]) if g else None
            if pr and pr.get('fig') and pr['fig'][0] != pr['fig'][1]: rep('T18', fn, title, '图旁表 %s——按 P06 收图（源文件 %%%%FIGSIDE%%%% 宽度）' % pr['info'])
        # T19 首列手动换行（SD-207 改：按「这一断能不能让首列变窄」判断；2026-10-17 用户 344「为什么要另外放在第二行」）
        #   术语格（SD-206）、括注另行（C29）不算；其余首列 <br>，合成一行后不超过本列其他格最宽一行的，断了也不省宽度 → 报
        def _first_cells(rows_):   # 首列格（被上方 rowspan 盖住的行，第一个 <td> 不是首列）
            out_, rem_ = [], 0
            for a_, cells_ in rows_:
                if rem_ > 0: rem_ -= 1; continue
                if not cells_: continue
                m_ = re.search(r'rowspan="(\d+)"', cells_[0][0]); rem_ = int(m_.group(1)) - 1 if m_ else 0
                out_.append((a_, cells_))
            return out_
        FIRST = _first_cells(rows)
        c0s = [cells[0][1] for a_, cells in FIRST if 'note' not in a_ and 'colspan' not in cells[0][0]]
        def _lines_w(c): return max([tw(x) for x in items(c)] or [0])
        for k0, c0 in enumerate(c0s):
            if not re.search(r'<br\s*/?>', c0) or re.match(r'^\s*[①-⑳]', plain(c0)): continue
            segs = items(c0)
            if len(segs) == 2 and re.fullmatch(r'（[^（）]+）', plain(segs[1])): continue                       # C29 括注另行
            if len(segs) == 2 and re.fullmatch(r'[A-Z][A-Za-z0-9/\-]{1,11}', plain(segs[0])): continue          # SD-206 术语格
            if all(plain(s).endswith(('。', '；')) for s in segs): continue                                        # 首列本身是句子列表（4.3 上座前后流程）
            joined = tw(re.sub(r'<br\s*/?>', '', c0))
            others = [_lines_w(c) for j, c in enumerate(c0s) if j != k0]
            if others and joined <= max(others):
                rep('T19', fn, title, '首列「%s」手动换行不省列宽（本列另有更宽的格）——合成一行' % plain(re.sub(r'<br\s*/?>', '｜', c0))[:34])
        # T29 首列括注另行（C29，SD-207 用户 2026-10-17）：首列「名称（括注）」一行写完，但同一行别的格已多行 → 括注放第二行
        if 'parallel' not in tb[:tb.find('>')]:
            for a0, cells in FIRST:
                if 'note' in a0 or len(cells) < 2 or 'colspan' in cells[0][0]: continue
                c0 = cells[0][1]; p0 = plain(c0)
                if '<br' in c0: continue
                m29 = re.search(r'（([^（）]*)）$', p0)
                if not m29 or len(m29.group(1)) <= 3 or len(p0[:m29.start()].strip()) < 2: continue
                if re.fullmatch(r'[约≈]?\s*[\d.,～~\-]+\s*[A-Za-z%℃℉°]+', m29.group(1).strip()): continue   # 单位换算（30m（100ft））不拆
                if any('<br' in c for _, c in cells[1:]) or 'rowspan' in cells[0][0]:
                    rep('T29', fn, title, '首列「%s」：本行已多行，括注应另起一行（C29）' % p0[:30])
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
# T27 标注与成品不符（SD-205）：读成品 docx——源文件表头显式 col-center（非 737 机型列）的列是否真居中；首列是否出现圆点（parallel 除外）
def _docx_tables(path):
    import zipfile
    if not os.path.exists(path): return []
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8'); out = []
    for m in re.finditer(r'<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>', x, re.S):
        rows = []
        for r in re.findall(r'<w:tr[ >].*?</w:tr>', m.group(0), re.S):
            cells = []
            for c in re.findall(r'<w:tc>.*?</w:tc>', r, re.S):
                span = int((re.search(r'<w:gridSpan w:val="(\d+)"', c) or [0, 1])[1])
                paras = re.findall(r'<w:p[ >].*?</w:p>', c, re.S)
                import html as _h; txt = _h.unescape(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', c)))
                jcs = [(re.search(r'<w:jc w:val="(\w+)"', p_) or [0, 'left'])[1] for p_ in paras if re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p_)]
                cells.append((span, txt, jcs, 'vMerge' in c and 'restart' not in c))
            rows.append(cells)
        out.append(rows)
    return out
_DOCX = {'速查': _docx_tables(os.path.join(REPO, 'build', 'B737机型理论速查.docx')), '全书': _docx_tables(os.path.join(REPO, 'build', 'B737机型理论速查笔记.docx'))}
def _norm(s): return re.sub(r'[\s\u200b\u2060\u00ad●•–]', '', s)
for f in FILES:
    fn = '速查' if f.endswith('速查源.md') else os.path.basename(f).split(' ')[0]
    D = _DOCX['速查' if fn == '速查' else '全书']
    if not D: continue
    for title, tb, prev in tables(f):
        hm = re.search(r'<tr class="hdr">(.*?)</tr>', tb, re.S)
        if not hm: continue
        ths = re.findall(r'<th([^>]*)>(.*?)</th>', hm.group(1), re.S)
        hdr = [_norm(plain(h)) for _, h in ths]
        ctr = []; col = 0
        for a_, h_ in ths:
            cs = int((re.search(r'colspan="(\d+)"', a_) or [0, 1])[1])
            if re.search(r'(^|\s)col-center(\s|$)', (re.search(r'class="([^"]*)"', a_) or [0, ''])[1]) and '737' not in plain(h_): ctr += list(range(col, col + cs))
            col += cs
        par = 'parallel' in tb[:80]
        first_row = _norm(plain((re.findall(r'<tr(?![^>]*class="(?:hdr|premise|note)")[^>]*>(.*?)</tr>', tb[hm.end():], re.S) or [''])[0]))[:12]
        for rows in D:
            hrow = next((r for r in rows if [_norm(c[1]) for c in r] == hdr), None)
            if not hrow: continue
            hi = rows.index(hrow)
            if first_row and first_row[:6] not in _norm(''.join(c[1] for r in rows[hi + 1:hi + 3] for c in r)): continue
            for r in rows[hi + 1:]:
                k = 0
                for span, txt, jcs, cont in r:
                    tt = _norm(txt)
                    if k in ctr and span == 1 and tt and tt not in ('—', '-') and not cont and any(j != 'center' for j in jcs):
                        rep('T27', fn, title, '表头标了 col-center 的「%s」列，成品里「%s」没居中——查规则优先级（layout-checklist 第零节 D）' % (plain(ths[min(k, len(ths) - 1)][1])[:12], txt[:12]))
                    if k == 0 and not par and txt.strip().startswith(('●', '•')):
                        rep('T27', fn, title, '首列「%s」成品里带圆点（C25 首列不加点）' % txt[:12])
                    k += span
# T30 首列被单个长格撑宽（SD-207，用户 2026-10-17 速查 405「第 1 列的空白空间为什么这么大？最新的规则为什么没有排查出来」）：
#   读成品 docx 的实际列宽：首列宽 ≥ 版心 25%，只有一格撑满（最宽一段 ≥ 列宽 90%），其余格最宽也不到列宽 60% → 这一格折行 / 括注另行，首列收窄
#   以前漏掉：T14 量的是「列宽比最长一行宽多少」，被一格撑满的列最长一行正好等于列宽，量出来空白为 0
for _bk, _p in (('速查', 'B737机型理论速查.docx'), ('全书', 'B737机型理论速查笔记.docx')):
    for grid, h0, c0, _ha in _docx_grid(os.path.join(REPO, 'build', _p)):
        if len(grid) < 2 or len(c0) < 3: continue
        W = grid[0]
        if W < 0.25 * TOTAL: continue
        ws = sorted(((min(W, max(tw(p) for p in ps) + PAD), ps, rp) for ps, rp in c0), key=lambda x: -x[0])
        tall = ws[0][2] >= 2 or max(tw(p) for p in ws[0][1]) + PAD > 1.05 * W   # 撑宽那一格所在行已多行（或本格已折行）：分行不增行
        if ''.join(ws[0][1]).count('、') >= 4: continue   # 首列这一格本身是「、」列举的内容（4.4「故障｜处置」），不是标签
        if tall and ws[0][0] >= 0.9 * W and ws[1][0] <= 0.6 * W and re.search(r'[，、 （：]', ''.join(ws[0][1])):
            rep('T30', _bk, h0, '首列（宽 %d%%）被「%s」一格撑宽，其余格最宽只占列宽 %d%%——这一格折行 / 括注另行，首列收窄' % (100 * W / TOTAL, ''.join(ws[0][1])[:24], 100 * ws[1][0] / W))
# T31 注内并列判读（SD-208，用户 2026-10-18 速查 21 AUTO FAIL：「注释太密集，可否换一种方式，比如做成表格前面加①②③?」）：
#   一条「注：」里有两组以上「A＝B」（全角等号＝判读 / 含义；半角 = 是公式，不算）并用「；」分开 → 改成表格（SD-173 条款优先用表格）
for f in FILES:
    fn = '速查' if f.endswith('速查源.md') else os.path.basename(f).split(' ')[0]; cur = ''
    for l in io.open(f, encoding='utf-8'):
        if l.startswith('#'): cur = l.lstrip('# ').strip(); continue
        if l.startswith('注：') and plain(l).count('＝') >= 2 and '；' in l:
            rep('T31', fn, cur, '注「%s」里有 %d 组并列判读——改成表格（指示 ｜ 含义）' % (plain(l)[:24], plain(l).count('＝')))
# ---- SD-209（2026-10-19 整体优化方案批 1）新增：T32 注密集、T33 条款未成表、T34 颜色、T35 缩写注缺中文 ----
def _entries(f):
    t = io.open(f, encoding='utf-8').read()
    return [(m.group(1).strip(), m.group(2)) for m in re.finditer(r'(?ms)^### (.+?)\n(.*?)(?=^### |^## |^%%PART%%|\Z)', t)]
_SRCS = sorted(glob.glob(os.path.join(REPO, 'notes_src', '[1-6]*', '*.md'))) + [os.path.join(REPO, '速查', '速查源.md')]
for f in _SRCS:
    fn = os.path.basename(f).replace('.md', '') if 'notes_src' in f else '速查'
    for title, body in _entries(f):
        notes = [l for l in body.split('\n') if l.startswith('注：') and not re.match(r'注：(?:缩写[：——]+)?[A-Z][A-Za-z/\-]{1,12}：[A-Z]', l)]   # 缩写注不计（批 3；1019R6 写法「注：缩写——X：…」）
        # 2026-10-23（用户，速查第 207 条：「改成注释放在表格下面」）：B37 改——注不再改「注 ｜ 说明」注表，一律「注：」行紧跟主表；T32 改查残留的注表
        if (ONLY is None or 'T32' in ONLY) and re.search(r'<tr class="hdr"><th[^>]*>注</th>', body):
            rep('T32', fn, title, '有「注 ｜ 说明」注表（B37：注一律写成「注：」行紧跟主表，能对应表格行的并回该行）')
        if ONLY is None or 'T33' in ONLY:
            for l in re.sub(r'<table.*?</table>', '', body, flags=re.S).split('\n'):
                if l.startswith(('<', '注：', '公司差异', '%%', '#')) or not l.strip(): continue
                if plain(l).count('；') >= 3 and re.search(r'任一|同时满足|以下|包括|分为', plain(l)):
                    rep('T33', fn, title, '段落「%s…」有 %d 个并列分句——条款改表格（B14）' % (plain(l)[:20], plain(l).count('；') + 1))
        if ONLY is None or 'T34' in ONLY:
            for c in re.findall(r'<td[^>]*>(.*?)</td>', body, re.S):
                # 2026-10-26（1026R2，优化意见第 5 条）：W02 原文「几个数值界限可并存；文字红与数值红并存时文字红降为黑粗」——
                # 只报 ≥ 2 处文字红，或文字红与数值红并存；几处数值红并存合规（原口径一格 ≥ 2 处红一律报，327 处多为合规的并列数值）
                ems = [plain(e) for e in re.findall(r'<em>(.*?)</em>', c, re.S)]
                txt = [e for e in ems if not re.search(r'\d', e)]; num = len(ems) - len(txt)
                if len(txt) >= 2 or (txt and num):
                    rep('T34', fn, title, '一格文字红 %d 处%s「%s…」（W02：文字红一格一处；与数值红并存时文字红降为黑粗）' % (len(txt), '、数值红 %d 处' % num if num else '', '／'.join(txt)[:24]))
        if ONLY is None or 'T35' in ONLY:
            for l in [l for l in body.split('\n') if l.startswith('注：')]:   # 1019R6：缩写注不计入 T32，但 T35 照查
                for seg in re.split(r'[；。]', plain(l)):
                    m = re.match(r'\s*(?:注：)?([A-Z][A-Za-z/\-]{1,12})：([A-Z][A-Za-z][A-Za-z /\-]*)(.*)$', seg)
                    _pb = plain(body)
                    if m:   # 1019R6：表格里缩写与中文名同格（「MEA<br>最低航路飞行高度」）或同一行有短中文名格（「IRS 显示组件 | 1 个 | ISDU」）也算条目里已有中文名
                        _ab = m.group(1); _hit = False
                        for _row in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
                            _cs = [plain(c).strip() for c in re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', _row, re.S)]
                            if any(c == _ab for c in _cs) and any(c != _ab and re.search(r'[\u4e00-\u9fff]', c) and len(c) <= 14 for c in _cs): _hit = True
                            if any(c.startswith(_ab) and re.fullmatch(_ab + r'\s*[\u4e00-\u9fff /、]{2,14}', c) for c in _cs): _hit = True
                        if _hit: continue
                    if m and not re.search(r'[\u4e00-\u9fff]', m.group(3)) and not re.search(re.escape(m.group(1)) + r'\s*[（(][^）)]*[\u4e00-\u9fff]|[\u4e00-\u9fff][）)]?\s*[（(]\s*' + re.escape(m.group(1)) + r'\s*[）)]', _pb):   # 英文全称后面没有中文；W06（1019R6）：条目里已有「缩写（中文）/ 中文（缩写）」的只写英文全称
                        rep('T35', fn, title, '缩写注「%s：%s」缺中文（W06：缩写：英文全称（中文））' % (m.group(1), m.group(2)[:30]))
# T40 序号表编号（1020R1，用户 2026-10-20：「③和⑩是哪里来的」）：首列 ①②… 的表必须从 ① 起连续；图表并排（图中标号对照）的表除外
_C = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
for f in _SRCS:
    fn = os.path.basename(f).replace('.md', '') if 'notes_src' in f else '速查'
    t = io.open(f, encoding='utf-8').read()
    for m in re.finditer(r'<table[^>]*>(.*?)</table>', t, re.S):
        pre = t[:m.start()].rstrip().rsplit('\n', 1)[-1]
        if pre.startswith(('%%FIGSIDE%%', '%%FIGRIGHT%%')): continue
        body = m.group(1)
        if 'rowspan' in body: continue
        firsts = re.findall(r'<tr(?![^>]*class="(?:hdr|premise|note|warn)")[^>]*>\s*<td[^>]*>\s*(?:<strong>)?([①-⑳])', body)
        if len(firsts) < 2: continue
        idx = [_C.index(x) for x in firsts]
        if idx[0] != 0 or any(b != a + 1 for a, b in zip(idx, idx[1:])):
            ttl = re.findall(r'(?m)^### (.+)$', t[:m.start()])
            if ONLY is None or 'T40' in ONLY: rep('T40', fn, (ttl[-1] if ttl else '')[:30], '序号表编号 %s 不从 ① 起连续——按本表重新编号，或改用「项目」列' % ''.join(firsts))
# T28 术语格（SD-206）：缩写定义格在内容列有多行的表里须写「缩写<br>中文」；内容列都一行的表须写「缩写（中文）」
try:
    _TERMS = set(json.load(open(os.path.join(REPO, '缩写术语表.json'), encoding='utf-8'))['缩写'])
except Exception:
    _TERMS = set()
if _TERMS:
    for f in FILES:
        fn = '速查' if f.endswith('速查源.md') else os.path.basename(f).split(' ')[0]
        for title, tb, prev in tables(f):
            rows = [re.findall(r'<td([^>]*)>(.*?)</td>', r, re.S) for r in re.findall(r'<tr(?![^>]*class="(?:hdr|premise|note)")[^>]*>(.*?)</tr>', tb, re.S)]
            tall = any(re.search(r'<br\s*/?>', c) or tw(c) > 0.45 * TOTAL for cc in rows for _, c in cc[1:])
            for cc in rows:
                if not cc: continue
                c0 = cc[0][1]; p0 = plain(c0)
                m1 = re.fullmatch(r'([A-Z][A-Za-z0-9/\-]{1,11})（[^）]{2,16}）', p0)
                m2 = re.match(r'^\s*(?:<[^>]+>)*([A-Z][A-Za-z0-9/\-]{1,11})(?:</[^>]+>)*<br\s*/?>', c0)
                if tall and m1 and m1.group(1) in _TERMS: rep('T28', fn, title, '术语格「%s」应写成两行（上缩写、下中文）' % p0)
                if not tall and m2 and m2.group(1) in _TERMS: rep('T28', fn, title, '内容列都一行，术语格「%s」应写成一行「缩写（中文）」' % p0)
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
# T32 / T35 为错误级（SD-209），但存量（70 / 34 处）要在批 2～3 清零后才加入退出码；sync 输出里照常计数
sys.exit(1 if any(c.get(k) for k in ('T19', 'T24', 'T27', 'T28', 'T29', 'T30', 'T31')) else 0)
