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
def convert(src):
    p=C(); p.feed(src); h=''.join(p.out)
    h=re.sub(r'<p style="[^"]*">(\s|<br/>)*<br/>(\s|<br/>)*</p>','<p style="margin-top:0;margin-bottom:0"><span style="font-size:4pt">&#160;</span></p>',h)  # 表间分隔段
    h=re.sub(r'<p style="[^"]*">(\s|<span[^>]*>\s*</span>)*</p>','',h)  # 去空段
    h=drop_sources(h)
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
