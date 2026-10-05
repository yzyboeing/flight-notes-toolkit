# 表格列宽排查：对比旧版（/tmp/old_conv.py）与当前 onenote_conv 的表宽和估算行数。用法：python3 tab_audit.py <基线docx>
import sys, re, html, importlib.util
sys.path.insert(0,'.')
from onenote_build import split
import onenote_conv as new
sys.path.insert(0,'/tmp'); import old_conv as old
pages=split(sys.argv[1])
def stats(h, mod):
    out=[]
    for t in re.findall(r'<table border="1".*?</table>', h, re.S):
        if 'eef4fb' in t.lower(): continue   # 注块
        tw=int((re.search(r'width:(\d+)px', t) or [0,0])[1]); lines=0; wrap=False
        for a,inner in re.findall(r'<td\b([^>]*)>(.*?)</td>', t, re.S):
            w=int((re.search(r'width:(\d+)px',a) or [0,0])[1]); cw=max(w-mod.PAD,1)
            for x in re.split(r'<br/>|</p>\s*<p[^>]*>', inner):
                x=html.unescape(re.sub(r'<[^>]+>','',x)).strip()
                if not x: continue
                L=max(1,-(-int(mod._seg_w(x))//cw)); lines+=L; wrap|=L>1
        out.append((tw,lines,wrap,re.sub(r'<[^>]+>|\s','',t)[:16]))
    return out
nb=nw=0; tot=0; still=0; rows=[]
for p in pages:
    a=stats(old.convert(p['html']),old); b=stats(new.convert(p['html']),new)
    for (tw0,l0,w0,k),(tw1,l1,w1,_) in zip(a,b):
        tot+=1
        if w0 and tw0<new.TABW-10: nb+=1
        if w1 and tw1<new.TABW-10: still+=1; rows.append(('仍有空位',p['title'],k,tw1))
        if l1>l0: rows.append(('增行',p['title'],k,tw0,tw1,l0,l1))
        if l1<l0: rows.append(('减行',p['title'],k,tw0,tw1,l0,l1))
print('表格',tot,'；旧算法「有折行但没到满宽」',nb,'；新算法',still,'；减少行数的表',sum(1 for r in rows if r[0]=='减行'),'共减',sum(r[5]-r[6] for r in rows if r[0]=='减行'),'行')
for r in rows[:400]: print(*r)
