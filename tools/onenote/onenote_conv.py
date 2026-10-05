# LibreOffice HTML → OneNote Graph 接口 HTML
import re,sys,html
from html.parser import HTMLParser
FONT="Songti SC"; PAGEW=700; SRCW=1020; TABW=670; DOT_PT=6.5; DOT_COLOR="#7f7f7f"
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
            s.out.append('<span style="font-family:%s;font-size:%spt;color:%s">•&nbsp;</span>'%(FONT,DOT_PT,DOT_COLOR)); return
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

CJK_W, ASC_W, PAD = 13.5, 7.0, 18      # 9pt 宋体-简在 OneNote 中的近似字宽（px）与格子左右内边距
SHORT_MAX, LONG_MIN = 240, 140
ORPHAN = 4 * CJK_W                     # 尾行不超过 4 个汉字宽视为短字，要消除         # SD-35③：短列一行排下；长句列不少于约 10 个汉字
def _seg_w(t):
    return sum(CJK_W if ord(ch) > 0x2E80 else ASC_W for ch in t)
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
            target = min(TABW, max(total, sum(W[i] for i in short) + LONG_MIN * len(long_)))
            rest = max(target - sum(W[i] for i in short), LONG_MIN * len(long_))
            lo = sum(orig[i] for i in long_) or len(long_)
            for i in long_: W[i] = max(LONG_MIN, int(rest * (orig[i] or 1) / lo))
        else:
            W = [max(W[i], orig[i]) for i in range(n)]
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
def convert(src):
    p=C(); p.feed(src); h=''.join(p.out)
    h=re.sub(r'<p style="[^"]*">(\s|<br/>)*<br/>(\s|<br/>)*</p>','<p style="margin-top:0;margin-bottom:0"><span style="font-size:4pt">&#160;</span></p>',h)  # 表间分隔段
    h=re.sub(r'<p style="[^"]*">(\s|<span[^>]*>\s*</span>)*</p>','',h)  # 去空段
    h=drop_sources(h)
    h=drop_xref(h)
    h=unshrink(h)
    h=fit_widths(h)
    h=note_blocks(h)
    return h
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
