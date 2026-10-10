# LibreOffice HTML → OneNote Graph 接口 HTML
import re,sys,html
from html.parser import HTMLParser
import os as _os
# 2026-10-09 用户：「主要排版规则不变的情况下可以适当加宽页面以让表格内容不那么拥挤」——正文框 700 → 860、最宽表格 670 → 830（表 ≈ 正文框 − 30，文字与表格右缘大致对齐）；ONENOTE_PAGEW 可调
PAGEW=int(_os.environ.get('ONENOTE_PAGEW', '860')); TABW=PAGEW-30
FONT="Songti SC"; SRCW=1020; DOT_PT=6.5; DOT_COLOR="#7f7f7f"
# 2026-10-09 用户：圆点改用 OneNote 自带列表「Small Solid Square」（接口写 list-style-type:square，读回与用户手设的一致）；
# 短线：接口不支持 OneNote 的「–」列表符号（自定义符号被丢弃，none 会跳出列表），ONENOTE_CHILD 选子项写法——
#   dash（默认）：保留文字「– 」，作为父项列表项下的缩进段（读回留在 <ul> 内）；circle：OneNote 二级空心圆列表
LISTS=_os.environ.get('ONENOTE_LISTS','1')!='0'; CHILD=_os.environ.get('ONENOTE_CHILD','dash')
BUL='\x01B\x01'   # 3 个字符：fit_widths 按约 21px 估列表缩进
K=TABW/SRCW
def px(v): return max(30,round(int(v)*K))
class C(HTMLParser):
    def __init__(s):
        super().__init__(convert_charrefs=True)
        s.out=[]; s.inp=0; s.st=[{}]; s.inbody=False; s.sup=0; s.blk=None; s.skip=0
    def cur(s): return s.st[-1]
    def push(s,**kw):
        d=dict(s.cur()); d.update({k:v for k,v in kw.items() if v is not None}); s.st.append(d)
    def handle_starttag(s,t,a):
        a=dict(a); sty=a.get('style','')
        if t=='body': s.inbody=True; return
        if not s.inbody: return
        if t in('style','title','script'): s.skip+=1; return
        if t=='font':
            sz=re.search(r'font-size:\s*([\d.]+)pt',sty)
            s.push(color=a.get('color'),size=sz.group(1) if sz else None); return
        if t=='b': s.push(bold=True); return
        if t=='u': s.push(u=True); return
        if t=='i': s.push(it=True); return
        if t=='sup': s.sup+=1; s.push(); return
        if t=='span':
            s.push(bold=False if 'font-weight: normal' in sty else None); return
        if t in('p','h1','h2','h3','h4'):
            al=a.get('align') or (re.search(r'text-align:\s*(\w+)',sty) or [None,None])[1]
            al={'center':'center','right':'right'}.get(al,'left')
            s.out.append('<p style="text-align:%s;margin-top:0;margin-bottom:0">'%al)
            s.push(bold=True if t[0]=='h' else None); s.blk=t; s.inp=1; s.lead=True; return
        if t=='table':
            w=a.get('width'); 
            s.out.append('<table border="1" style="border-collapse:collapse%s">'%(';width:%dpx'%px(w) if w else '')); return
        if t=='tr': s.out.append('<tr>'); return
        if t in('td','th'):
            st=['border:1px solid #bfbfbf']
            bg=a.get('bgcolor') or (re.search(r'background:\s*(#\w+)',sty) or [None,None])[1]
            if bg and bg.lower() not in('#ffffff','transparent'): st.append('background-color:%s'%bg)
            if a.get('width'): st.append('width:%dpx'%px(a['width']))
            span=''.join(' %s="%s"'%(k,a[k]) for k in('rowspan','colspan') if k in a)
            s.out.append('<td style="%s"%s>'%(';'.join(st),span)); return
        if t=='br': s.out.append('<br/>'); return
    def handle_endtag(s,t):
        if not s.inbody: return
        if t in('style','title','script'): s.skip-=1; return
        if t in('font','b','u','i','span','sup'):
            if len(s.st)>1: s.st.pop()
            if t=='sup': s.sup-=1
            return
        if t in('p','h1','h2','h3','h4'):
            if len(s.st)>1: s.st.pop()
            s.inp=0; s.out.append('</p>'); return
        if t=='table': s.out.append('</table>'); return
        if t=='tr': s.out.append('</tr>'); return
        if t in('td','th'): s.out.append('</td>'); return
    def handle_data(s,d):
        if not s.inbody or s.skip: return
        if not s.inp: return
        d=re.sub(r'\s*\n\s*',' ',d)
        if s.lead: d=d.lstrip()
        if not d: return
        s.lead=False
        c=s.cur()
        if s.sup and '●' in d:
            s.out.append(BUL if LISTS else '<span style="font-family:%s;font-size:%spt;color:%s">•&nbsp;</span>'%(FONT,DOT_PT,DOT_COLOR)); return
        st=['font-family:%s'%FONT,'font-size:%spt'%(c.get('size') or '10')]
        if c.get('color') and c['color'].lower()!='#000000': st.append('color:%s'%c['color'])
        if c.get('bold'): st.append('font-weight:bold')
        if c.get('u'): st.append('text-decoration:underline')
        if c.get('it'): st.append('font-style:italic')
        s.out.append('<span style="%s">%s</span>'%(';'.join(st),html.escape(d,quote=False)))

SRC_STATS={'del':0,'keep_xref':0}
def drop_sources(h):
    def f(m):
        p=m.group(0)
        t=re.sub(r'\s+',' ',html.unescape(re.sub(r'<[^>]+>','',p))).strip()
        if not re.match(r'(来源|出处)：',t): return p
        if '详见' in t:
            SRC_STATS['keep_xref']+=1
            x=t[t.find('详见'):]
            return '<p style="text-align:left;margin-top:0;margin-bottom:0"><span style="font-family:%s;font-size:9pt;color:#7f7f7f">%s</span></p>'%(FONT,html.escape(x,quote=False))
        SRC_STATS['del']+=1
        return ''
    return re.sub(r'<p\b[^>]*>.*?</p>',f,h,flags=re.S)

XREF_STATS={'line':0,'note_del':0,'note_strip':0}
def drop_xref(h):
    """OneNote 版不写「详见」（用户 2026-10-05「掉 OneNote 里面的详见和来源」）：
    独立的「详见 x.y」行删除；注「注：标题——实际内容——详见 x.y」只删「——详见 x.y」保留内容；其余指路注整条删。"""
    def f(m):
        p = m.group(0)
        t = re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', '', p))).strip()
        if '详见' not in t: return p
        if t.startswith('详见'): XREF_STATS['line'] += 1; return ''
        body = re.sub(r'^注[：:]\s*', '', t)
        mm = re.match(r'^[^—]+——(.+)——\s*详见[^—；;]*$', body)
        if mm and body.count('详见') == 1 and '，' in mm.group(1):      # 「标题——带逗号的实际内容——详见 x.y」才保留内容
            k = p.rfind('——详见')
            if k > 0: XREF_STATS['note_strip'] += 1; return p[:k] + '。</span></p>'
        XREF_STATS['note_del'] += 1; return ''
    return re.sub(r'<p\b[^>]*>.*?</p>', f, h, flags=re.S)

NOTE_BG = '#eef4fb'   # SD-130：表后「注：」淡蓝底整块；连续几条注连成一块
def note_blocks(h):
    out, i = [], 0
    items = list(re.finditer(r'<table\b.*?</table>|<p\b[^>]*>.*?</p>', h, re.S))
    buf, last = [], 0
    def flush():
        if buf:
            out.append('<table><tr><td style="background-color:%s;width:%dpx">%s</td></tr></table>' % (NOTE_BG, TABW, ''.join(buf)))
            buf.clear()
    for m in items:
        g = m.group(0)
        gap = h[last:m.start()]
        if gap.strip(): flush(); out.append(gap)
        t = html.unescape(re.sub(r'<[^>]+>', '', g)).strip()
        if g.startswith('<p') and re.match(r'注[：:]', t): buf.append(g)
        else: flush(); out.append(g)
        last = m.end()
    flush(); out.append(h[last:])
    return ''.join(out)

# 2026-10-06 用户：「OneNote 所有排版规则和要求都跟目前的最新规则一致」——
# w-NN 定宽表（SD-153）按成品列宽比例（docx 宽度 × K，与 PDF 同占比）；col-eq 列（SD-152）等宽
import glob as _glob, os as _os
_NS = _os.path.expanduser('~/flight-repos/gh-private/notes_src/')
def _hkey(cells): return ''.join(re.sub(r'[\s\u200b\u2060\u00ad]+', '', html.unescape(re.sub(r'<[^>]+>', '', c))) for c in cells)
def _sources():
    """定宽 / 等宽标记的来源：环境变量 ONENOTE_SRC_REF（如 baseline/20261006-3）给了就读该基线，避免混进别的会话正在改的源文件"""
    ref = _os.environ.get('ONENOTE_SRC_REF')
    if _os.environ.get('ONENOTE_SPEC') == 'quickref':      # 速查版（onenote_qr.py）：定宽 / 等宽标记读 速查/速查源.md
        repo = _os.path.dirname(_NS.rstrip('/')); f = '速查/速查源.md'
        if not ref: return [open(_os.path.join(repo, f), encoding='utf-8').read()]
        import subprocess
        return [subprocess.run(['git', '-C', repo, 'show', '%s:%s' % (ref, f)], capture_output=True, text=True, check=True).stdout]
    if not ref:
        return [open(f, encoding='utf-8').read() for f in _glob.glob(_NS + '*/*.md')]
    import subprocess
    repo = _os.path.dirname(_NS.rstrip('/'))
    names = subprocess.run(['git', '-C', repo, '-c', 'core.quotepath=off', 'ls-tree', '-r', '--name-only', ref, 'notes_src'], capture_output=True, text=True, check=True).stdout.split('\n')
    return [subprocess.run(['git', '-C', repo, 'show', '%s:%s' % (ref, n)], capture_output=True, text=True).stdout for n in names if n.endswith('.md')]
def load_specs():
    fixed, eq = set(), {}
    for text in _sources():
        for m in re.finditer(r'<tr class="hdr">(.*?)</tr>', text, re.S):
            ths = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1), re.S)
            k = _hkey([t[1] for t in ths]); cls = [(re.search(r'class="([^"]*)"', a) or [0, ''])[1] for a, _ in ths]
            if any(re.search(r'\bw-\d', c) for c in cls): fixed.add(k)
            if any('col-eq' in c for c in cls):
                col, idx = 0, []
                for (a, _), c in zip(ths, cls):
                    cs = int((re.search(r'colspan="(\d+)"', a) or [0, 1])[1])
                    if 'col-eq' in c: idx += list(range(col, col + cs))
                    col += cs
                eq[k] = idx
    return fixed, eq
FIXED, EQ = load_specs()

CJK_W, ASC_W, PAD = 13.5, 7.0, 18      # 9pt 宋体-简在 OneNote 中的近似字宽（px）与格子左右内边距
SHORT_MAX, LONG_MIN = 240, 140
ORPHAN = 4 * CJK_W                     # 尾行不超过 4 个汉字宽视为短字，要消除         # SD-35③：短列一行排下；长句列不少于约 10 个汉字
def _seg_w(t):
    return sum(CJK_W if ord(ch) > 0x2E80 or 0x2190 <= ord(ch) <= 0x22FF or ch in '±×' else ASC_W for ch in t)   # 箭头、≤ ≥ ≈ × ± 在宋体里是全角（2026-10-05 实测）
def fit_widths(h):
    """按原笔记 SD-35③ / SD-85 重新定列宽：短列（标签、数值、序号）按最长一行排下；长句列分剩余宽度。"""
    def one(m):
        t = m.group(0)
        rows = re.findall(r'<tr>(.*?)</tr>', t, re.S)
        grid, cells = {}, []          # (行, 列) 占位；cells: (行, 起列, 跨列, 每行文字段, 原宽)
        for r, row in enumerate(rows):
            c = 0
            for cm in re.finditer(r'<td\b([^>]*)>(.*?)</td>', row, re.S):
                while (r, c) in grid: c += 1
                attrs, inner = cm.group(1), cm.group(2)
                cs = int((re.search(r'colspan="(\d+)"', attrs) or [0, 1])[1]); rs = int((re.search(r'rowspan="(\d+)"', attrs) or [0, 1])[1])
                w0 = int((re.search(r'width:(\d+)px', attrs) or [0, 0])[1])
                segs = [html.unescape(re.sub(r'<[^>]+>', '', x)).strip() for x in re.split(r'<br/>|</p>\s*<p[^>]*>', inner)]
                cells.append((r, c, cs, segs, w0))
                for dr in range(rs):
                    for dc in range(cs): grid[(r + dr, c + dc)] = 1
                c += cs
        n = max((c + cs for _, c, cs, _, _ in cells), default=0)
        if n < 2: return t
        hkey = ''.join(''.join(segs) for r, c, cs, segs, w0 in cells if r == 0).replace(' ', '')
        hkey = re.sub(r'[\s\u200b\u2060\u00ad]+', '', hkey)
        need, orig = [0] * n, [0] * n
        for r, c, cs, segs, w0 in cells:
            if cs == 1:
                need[c] = max(need[c], max((_seg_w(x) for x in segs), default=0) + PAD)
                orig[c] = max(orig[c], w0)
        total = max(sum(orig), 1)
        short = [i for i in range(n) if need[i] <= SHORT_MAX]
        long_ = [i for i in range(n) if i not in short]
        W = [0] * n
        for i in short: W[i] = max(int(need[i] + 0.5), 30)
        if long_:
            # 2026-10-05 用户：「明明右边还是有空位的，可以让表格尽量宽一点，这样行数就会少一点」——
            # 长句列只要会折行，表格就加宽，直到长句都一行排下或到满宽 TABW 为止（不再锁在原表宽的等比例上）
            target = min(TABW, max(total, sum(W[i] for i in short) + sum(max(need[i], LONG_MIN) for i in long_)))
            rest = max(target - sum(W[i] for i in short), LONG_MIN * len(long_))
            # 按原宽比例分给长句列；某列分到的超过它一行排下所需，多出的再分给其余长句列
            def fill(capped, rest=rest):
                V, todo, left = list(W), list(long_), rest
                while todo:
                    lo = sum(orig[i] or 1 for i in todo)
                    cap = [i for i in todo if capped and left * (orig[i] or 1) / lo >= need[i]]
                    if not cap:
                        for i in todo: V[i] = max(LONG_MIN, int(left * (orig[i] or 1) / lo))
                        break
                    for i in cap: V[i] = max(LONG_MIN, int(need[i] + 0.5)); left -= V[i]; todo.remove(i)
                return V
            t0 = min(TABW, max(total, sum(W[i] for i in short) + LONG_MIN * len(long_)))   # 旧做法：表宽按原表比例
            cands = [fill(True), fill(False), fill(False, max(t0 - sum(W[i] for i in short), LONG_MIN * len(long_)))]   # 几种分法都做完后处理，取总行数少的（同样行数取先者，即更宽的）
        else:
            cands = [[max(W[i], orig[i]) for i in range(n)]]
        def nlines(V):
            return sum(max(1, -(-int(_seg_w(x)) // max(int(sum(V[c:c + cs]) - PAD), 1)))
                       for r, c, cs, segs, w0 in cells for x in segs if x)
        def wraps(W):
            return any(_seg_w(x) > sum(W[c:c + cs]) - PAD for r, c, cs, segs, w0 in cells for x in segs if x)
        def balance(W):
            """SD-151：表宽不超 TABW；超了就从加行最少的列收；有折行且有空位时，把空位给最省行的列。"""
            W = list(W)
            while sum(W) > TABW:
                ex = sum(W) - TABW
                best = None
                for i in range(n):
                    if W[i] <= 40: continue
                    d = min(ex, 4); V = list(W); V[i] -= d
                    cost = (nlines(V), -W[i])
                    if best is None or cost < best[0]: best = (cost, i, d)
                if not best: break
                W[best[1]] -= best[2]
            while True:
                room = TABW - sum(W); best = None
                if room <= 0: break
                for r, c, cs, segs, w0 in cells:
                    cw = sum(W[c:c + cs]) - PAD
                    for x in segs:
                        if not x: continue
                        L = max(1, -(-int(_seg_w(x)) // max(int(cw), 1)))
                        if L < 2: continue
                        d = -(-int(_seg_w(x)) // (L - 1)) - cw
                        if 0 < d <= room:
                            V = list(W); V[c + cs - 1] += d
                            gain = nlines(W) - nlines(V)
                            if gain > 0 and (best is None or d / gain < best[0]): best = (d / gain, c + cs - 1, d)
                if not best: break
                W[best[1]] += best[2]
            # SD-151：「哪怕减少不了行数，也可以把空白区域利用起来」——仍有折行就把余下空位给折行最多的列
            room = TABW - sum(W)
            if room > 0 and wraps(W):
                cnt = [0] * n
                for r, c, cs, segs, w0 in cells:
                    cw = sum(W[c:c + cs]) - PAD
                    for x in segs:
                        if x: cnt[c + cs - 1] += max(1, -(-int(_seg_w(x)) // max(int(cw), 1))) - 1
                W[max(range(n), key=lambda i: cnt[i])] += room
            return W
        done = []
        for W in cands:
            # 跨列格（colspan）里的长句：所跨各列之和不够一行排下时，把差额匀给这几列，表宽以 TABW 为限
            for r, c, cs, segs, w0 in cells:
                if cs < 2: continue
                nw = max((_seg_w(x) for x in segs), default=0) + PAD
                gap = min(nw - sum(W[c:c + cs]), TABW - sum(W))
                if gap > 0:
                    for k in range(cs): W[c + k] += int(gap / cs) + (1 if k < gap % cs else 0)
            # SD-85：消除尾行短字——某格最后一行只剩 ≤ ORPHAN 宽度的字时，给该列加宽，从不会因此多折一行的列匀出宽度
            segs_by_col = [[] for _ in range(n)]
            for r, c, cs, segs, w0 in cells:
                if cs == 1: segs_by_col[c] += [_seg_w(x) for x in segs if x]
            def lines(w, cw): return max(1, -(-int(w) // max(int(cw), 1)))
            def demand(i):
                cw = W[i] - PAD; best = 0
                for w in segs_by_col[i]:
                    L = lines(w, cw)
                    if L >= 2 and w - (L - 1) * cw <= ORPHAN:
                        best = max(best, int(w / (L - 1) - cw) + 2)
                return best
            def slack(i):
                cw = W[i] - PAD; sl = W[i] - 30
                for w in segs_by_col[i]:
                    L = lines(w, cw); sl = min(sl, int(cw - w / L))
                return max(sl, 0)
            for _ in range(3 * n):
                changed = False
                for i in range(n):
                    d = demand(i)
                    if not d or d > 120: continue
                    room = max(0, TABW - sum(W))
                    donors = sorted([j for j in range(n) if j != i], key=lambda j: -slack(j))
                    give = min(d, room); takes = []
                    for j in donors:
                        if give >= d: break
                        tk = min(slack(j), d - give)
                        if tk > 0: takes.append((j, tk)); give += tk
                    if give >= d:
                        for j, tk in takes: W[j] -= tk
                        W[i] += d; changed = True
                if not changed: break
            done.append(W)
        W = min(done, key=nlines)
        if hkey in FIXED and all(orig):
            # SD-153 定宽表：列间比例照成品（orig 已是 docx 宽度 × K）；OneNote 字相对更大，
            # 表宽从成品占比起按同一比例放大，取行数降到最少的最小表宽（SD-151 表宽以文字成行为准），以 TABW 为限
            base = [max(30, o) for o in orig]; fmax = TABW / sum(base)
            fs = [1 + k * 0.02 for k in range(int((fmax - 1) / 0.02) + 1)] + [fmax] if fmax > 1 else [fmax]
            sc = lambda f: [max(30, int(o * f)) for o in base]
            best = min(nlines(sc(f)) for f in fs)
            W = sc(next(f for f in fs if nlines(sc(f)) == best))
            if wraps(W): W = sc(fmax)                           # 仍有折行：照比例拉满（「有空位就别折行」）
        else:
            W = balance(W)
        if hkey in EQ:                                         # SD-152 col-eq：这几列等宽
            ix = [i for i in EQ[hkey] if i < n]
            if ix:
                v = sum(W[i] for i in ix) // len(ix)
                for i in ix: W[i] = v
        # 写回：每格宽度 = 所跨各列之和
        k = [0]
        def td(cm):
            r, c, cs, segs, w0 = cells[k[0]]; k[0] += 1
            w = sum(W[c:c + cs])
            a = re.sub(r'width:\d+px', 'width:%dpx' % w, cm.group(1)) if 'width:' in cm.group(1) else cm.group(1).replace('style="', 'style="width:%dpx;' % w, 1)
            return '<td%s>%s</td>' % (a, cm.group(2))
        t2 = re.sub(r'<td\b([^>]*)>(.*?)</td>', td, t, flags=re.S)
        return re.sub(r'(<table border="1" style="[^"]*?)width:\d+px', r'\1width:%dpx' % sum(W), t2, count=1)
    return re.sub(r'<table border="1".*?</table>', one, h, flags=re.S)

def unshrink(h):
    """PDF 为整表同页把部分表压到 8～8.5pt（C9）；OneNote 不分页，统一恢复 9pt，表内各字号按同一比例放大。"""
    def one(m):
        t = m.group(0)
        sizes = [float(x) for x in re.findall(r'font-size:([\d.]+)pt', t)]
        if not sizes: return t
        base = max(set(sizes), key=sizes.count)
        if base >= 9: return t
        k = 9.0 / base
        return re.sub(r'font-size:([\d.]+)pt', lambda q: 'font-size:%gpt' % (round(float(q.group(1)) * k * 2) / 2), t)
    return re.sub(r'<table border="1".*?</table>', one, h, flags=re.S)
def unmerge(h):
    """OneNote 不支持合并格：自己拆开时会把整格宽度记到第一列，表被撑宽（2026-10-06 读回发现，三种复飞方式 668 → 1456px）。
    这里先拆成单格：内容放在左上格，其余补空格；每格写本列宽度，底色沿用。合并由用户按《OneNote合并单元格清单》手动做。"""
    def one(m):
        t = m.group(0)
        if 'colspan=' not in t and 'rowspan=' not in t: return t
        rows = re.findall(r'<tr>(.*?)</tr>', t, re.S)
        grid, cells = {}, []
        for r, row in enumerate(rows):
            c = 0
            for cm in re.finditer(r'<td\b([^>]*)>(.*?)</td>', row, re.S):
                while (r, c) in grid: c += 1
                a = cm.group(1)
                cs = int((re.search(r'colspan="(\d+)"', a) or [0, 1])[1]); rs = int((re.search(r'rowspan="(\d+)"', a) or [0, 1])[1])
                w = int((re.search(r'width:(\d+)px', a) or [0, 0])[1])
                cells.append((r, c, cs, rs, a, cm.group(2), w))
                for dr in range(rs):
                    for dc in range(cs): grid[(r + dr, c + dc)] = (r, c)
                c += cs
        n = max(c + cs for r, c, cs, rs, a, x, w in cells)
        W = [0] * n
        for r, c, cs, rs, a, x, w in cells:
            if cs == 1: W[c] = max(W[c], w)
        for r, c, cs, rs, a, x, w in sorted(cells, key=lambda z: z[2]):      # 没有单列格的列：均分跨列格的余量
            miss = [k for k in range(c, c + cs) if not W[k]]
            if miss:
                left = max(w - sum(W[c:c + cs]), 30 * len(miss))
                for k in miss: W[k] = left // len(miss)
        out = []
        for r in range(len(rows)):
            tds = []
            for c in range(n):
                o = grid.get((r, c))
                if not o: continue
                cell = next(z for z in cells if z[0] == o[0] and z[1] == o[1])
                a = re.sub(r'\s(rowspan|colspan)="\d+"', '', cell[4])
                a = re.sub(r'width:\d+px', 'width:%dpx' % W[c], a)
                tds.append('<td%s>%s</td>' % (a, cell[5] if (r, c) == o else ''))
            out.append('<tr>%s</tr>' % ''.join(tds))
        head = re.match(r'<table[^>]*>', t).group(0)
        head = re.sub(r'width:\d+px', 'width:%dpx' % sum(W), head)
        return head + ''.join(out) + '</table>'
    return re.sub(r'<table border="1".*?</table>', one, h, flags=re.S)

def unwrap_figside(h):
    """2026-10-09：Word 版「左图右表」（%%FIGSIDE%%）是一张两格排版表：左格图 + 图注、右格正文表。OneNote 不带图片，
    外框只剩图注，正文表被挤在右格里（1009R4 速查 06 页读回表宽 1559 / 对齐错）。拆掉外框，只留正文表；图由补图步骤另插。"""
    out, i = [], 0
    while True:
        s = h.find('<table border="1"', i)
        if s < 0: out.append(h[i:]); break
        m = re.match(r'<table border="1"[^>]*><tr><td[^>]*>((?:(?!<td|<table).)*?)</td><td[^>]*>(<table border="1")', h[s:], re.S)
        if not m: out.append(h[i:s + 1]); i = s + 1; continue
        inner_s = s + m.start(2); depth, k = 0, inner_s
        while True:
            a, b = h.find('<table', k), h.find('</table>', k)
            if a != -1 and a < b: depth += 1; k = a + 6
            else:
                depth -= 1; k = b + 8
                if depth == 0: break
        inner = h[inner_s:k]
        end = h.find('</table>', k) + 8          # 外框的 </td></tr></table>
        out.append(h[i:s]); out.append(inner); i = end
    return ''.join(out)
def convert(src):
    p=C(); p.feed(src); h=''.join(p.out)
    h=re.sub(r'<p style="[^"]*">(\s|<br/>)*<br/>(\s|<br/>)*</p>','<p style="margin-top:0;margin-bottom:0"><span style="font-size:4pt">&#160;</span></p>',h)  # 表间分隔段
    h=re.sub(r'<p style="[^"]*">(\s|<span[^>]*>\s*</span>)*</p>','',h)  # 去空段
    h=drop_sources(h)
    h=drop_xref(h)
    h=unwrap_figside(h)
    h=unshrink(h)
    h=fit_widths(h)
    h=unmerge(h)
    h=note_blocks(h)
    if LISTS: h=to_lists(h)
    return h
_PARA=re.compile(r'<p\b([^>]*)>(.*?)</p>',re.S)
def _is_dash(inner):
    t=html.unescape(re.sub(r'<[^>]+>','',inner)).lstrip()
    return t.startswith('– ') or t.startswith('–\u00a0')
def _undash(inner):
    return re.sub(r'^((?:<span[^>]*>)?)\s*–[\s\u00a0]','\\1',inner,count=1)
def to_lists(h):
    """父项（生成器的「●」）→ <ul><li style="list-style-type:square">；紧跟的「– 」子项挂在该父项下（见 CHILD）。
    只把相邻（中间只有空白）的段落连成一组；不跟在父项后的「– 」段落保持原样。"""
    out=[]; i=0; items=list(_PARA.finditer(h)); k=0
    while k<len(items):
        m=items[k]
        if BUL not in m.group(2):
            if CHILD=='circle' and _is_dash(m.group(2)):          # 跟在引语后的独立短线组 → 一级空心圆列表
                grp=[m]; j=k+1
                while j<len(items) and not h[grp[-1].end():items[j].start()].strip() and _is_dash(items[j].group(2)):
                    grp.append(items[j]); j+=1
                out.append(h[i:m.start()])
                out.append('<ul>'+''.join('<li style="list-style-type:circle"><p%s>%s</p></li>'%(g.group(1),_undash(g.group(2))) for g in grp)+'</ul>')
                i=grp[-1].end(); k=j; continue
            k+=1; continue
        grp=[m]; j=k+1
        while j<len(items) and not h[grp[-1].end():items[j].start()].strip() and (BUL in items[j].group(2) or _is_dash(items[j].group(2))):
            grp.append(items[j]); j+=1
        out.append(h[i:m.start()])
        lis=[]; cur=None
        for g in grp:
            a,inner=g.group(1),g.group(2)
            if BUL in inner:
                if cur is not None: lis.append(cur)
                cur=['<p%s>%s</p>'%(a,inner.replace(BUL,'',1)),[]]
            else:
                cur[1].append((a,inner))
        lis.append(cur)
        u=[]
        for par,kids in lis:
            body=par
            if kids:
                if CHILD=='circle':
                    body+='<ul>'+''.join('<li style="list-style-type:circle"><p%s>%s</p></li>'%(a,_undash(x)) for a,x in kids)+'</ul>'
                else:
                    body+=''.join('<p%s>%s</p>'%(a,x) for a,x in kids)
            u.append('<li style="list-style-type:square">%s</li>'%body)
        out.append('<ul>'+''.join(u)+'</ul>')
        i=grp[-1].end(); k=j
    out.append(h[i:])
    h=''.join(out)
    # 2026-10-09 实测：同一格 / 同一段里「列表 → 普通段落 → 列表」，OneNote 导入会把后一组并进前一组当子项、中间段挤到最后（内容顺序变了）；
    # 加 div、空段、改 ol 都不行——这种相邻的几组一律退回文字圆点（与 1009R4 及以前写法相同）
    SEQ=re.compile(r'<ul>(?:(?!</?ul>).)*</ul>(?:(?:\s*<p\b[^>]*>(?:(?!</p>).)*</p>)+\s*<ul>(?:(?!</?ul>).)*</ul>)+',re.S)
    h=SEQ.sub(lambda m:_unlist(m.group(0)),h)
    return h.replace(BUL,'')
DOT='<span style="font-family:%s;font-size:%spt;color:%s">•&nbsp;</span>'%(FONT,DOT_PT,DOT_COLOR)
def _unlist(x):
    x=re.sub(r'<li style="list-style-type:circle"><p([^>]*)>(<span[^>]*>)?','\\n<p\\1>\\2– ',x)
    x=re.sub(r'<li style="list-style-type:square"><p([^>]*)>(<span[^>]*>)?',lambda m:'<p%s>%s%s'%(m.group(1),DOT,m.group(2) or ''),x)
    return re.sub(r'</?(?:ul|li)\b[^>]*>','',x)
def flat_lists(h):
    """核对用：把列表还原成段落（方块 → 段首「•」，空心圆 → 段首「–」），读回与期望同口径比较对齐和圆点。"""
    mk=lambda a:'•' if 'square' in a else ('–' if 'circle' in a else '')
    h=re.sub(r'<li\b([^>]*)>\s*<p\b([^>]*)>',lambda m:'<p%s>%s'%(m.group(2),mk(m.group(1))),h)
    h=re.sub(r'<li\b([^>]*)>(.*?)(?=</li>|<ul\b|<p\b)',lambda m:'<p>%s%s</p>'%(mk(m.group(1)),m.group(2)),h,flags=re.S)
    return re.sub(r'</?(?:ul|ol|li)\b[^>]*>','',h)
def strip_index(s):
    m=re.search(r'<h\d[^>]*>(?:(?!</h\d>).)*块索引(?:(?!</h\d>).)*</h\d>\s*',s,re.S)
    if not m: return s,False
    rest=s[m.end():]
    t=re.match(r'<table\b.*?</table>\s*',rest,re.S)
    assert t, '块索引后不是表格'
    assert '<table' not in t.group(0)[6:], '嵌套表格'
    return s[:m.start()]+rest[t.end():],True
def page(title,body,header=''):
    hd=('<div style="position:absolute;left:36px;top:110px;width:%dpx"><p style="margin-top:0;margin-bottom:0"><span style="font-family:%s;font-size:8pt;color:#8c8c8c">%s</span></p></div>'%(PAGEW,FONT,html.escape(header))) if header else ''
    return ('<!DOCTYPE html><html><head><title>%s</title><meta charset="utf-8"/></head>'
            '<body data-absolute-enabled="true">%s<div style="position:absolute;left:36px;top:140px;width:%dpx">%s</div></body></html>')%(html.escape(title),hd,PAGEW,body)
if __name__=='__main__':
    src=open(sys.argv[1]).read(); out=page(sys.argv[2],convert(src)); open(sys.argv[3],'w').write(out)
    print(len(src)//1024,'KB ->',len(out)//1024,'KB')
