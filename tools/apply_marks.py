#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""apply_marks.py —— 按清单给速查区条目加 / 改强调标记（remark_quickref.py 之后的逐条例外）

用法：python3 tools/apply_marks.py <清单.json> [--src notes_src] [--write]
清单每项：[条目号, 原文片段, "em" | "strong", (次数)] 包裹；或 [条目号, 旧串, "raw" | "raw_all", 新串] 直接替换。
只动标记：写入前后去标签全文必须一致。
"""
import re,io,sys,json,os
import glob
SRC=sys.argv[sys.argv.index('--src')+1] if '--src' in sys.argv else 'notes_src'
F=glob.glob(os.path.join(SRC,'0 *','0 *.md'))[0]
t=io.open(F,encoding='utf-8').read()
ops=json.load(io.open(sys.argv[1],encoding='utf-8'))
bad=0
for n,s,tag,*rest in ops:
    cnt=rest[0] if rest else 1
    m=re.search(r'(?m)^### %d\. '%n,t); st=m.start()
    e=re.search(r'(?m)^##',t[m.end():]); en=m.end()+e.start() if e else len(t)
    sec=t[st:en]
    if tag=='raw_all':
        c=sec.count(s)
        if not c: print('FAIL raw_all',n,repr(s)); bad+=1; continue
        sec=sec.replace(s,rest[0]); t=t[:st]+sec+t[en:]; continue
    if tag=='raw':
        if sec.count(s)!=1: print('FAIL raw',n,repr(s),sec.count(s)); bad+=1; continue
        sec=sec.replace(s,rest[0]); t=t[:st]+sec+t[en:]; continue
    # 只替换未被同类标签直接包住的出现
    pat=re.compile(r'(?<!<%s>)%s'%(tag,re.escape(s)))
    found=[x for x in pat.finditer(sec)]
    # 排除落在标签内部（属性）或已在 <em> 里的情况：检查左侧最近的 <em>/</em>
    ok=[]
    for x in found:
        left=sec[:x.start()]
        if tag=='em' and left.rfind('<em>')>left.rfind('</em>'): continue
        if tag=='strong' and left.rfind('<strong>')>left.rfind('</strong>'): continue
        ok.append(x)
    if (cnt!='all' and len(ok)!=cnt) or not ok:
        print('FAIL',n,repr(s),tag,'found',len(ok)); bad+=1; continue
    for x in reversed(ok):
        sec=sec[:x.start()]+'<%s>%s</%s>'%(tag,s,tag)+sec[x.end():]
    t=t[:st]+sec+t[en:]
plain=lambda x: re.sub(r'</?(em|strong)>','',x)
assert plain(t)==plain(io.open(F,encoding='utf-8').read()), '文字发生变化，中止'
if '--write' in sys.argv and not bad: io.open(F,'w',encoding='utf-8').write(t); print('written',len(ops))
elif '--write' in sys.argv: 
    io.open(F,'w',encoding='utf-8').write(t); print('written with',bad,'failures skipped')
else: print('dry',len(ops),'bad',bad)
