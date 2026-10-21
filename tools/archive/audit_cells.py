import json,sys,re
B,C,P,Q,N='','','','',''
def show(t): return t.replace(B,'•').replace(C,'–').replace(P,'[P]').replace(Q,'[Q]').replace(N,'[N]')
seen=set(); out={'introQ':[], 'mixcell':[], 'mixcol':[], 'introB':[]}
for ln in open(sys.argv[1],encoding='utf-8'):
    tb=json.loads(ln); key=json.dumps(tb,ensure_ascii=False)
    if key in seen: continue
    seen.add(key); h='/'.join(tb['hdr'])[:30]
    cols={}
    for r in tb['rows']:
        if re.search('hdr|note|premise|warn',r['cls']): continue
        for c in r['cells']:
            if c['head'] or c['col']==0 or c['span']!=1: continue
            t=c['t']; ls=[x for x in t.split('\n') if x.strip()]
            cols.setdefault(c['col'],[]).append(t)
            for i,x in enumerate(ls[:-1]):
                if x.startswith(B) and re.search('[：:]$',x.strip()):
                    nx=ls[i+1]
                    if nx.startswith(Q): out['introQ'].append((h,show(x)[:40],show(nx)[:30]))
                    elif nx.startswith(B): out['introB'].append((h,show(x)[:40],show(nx)[:30]))
            if any(x[0] in B+C+P+Q+N for x in ls) and any(x[0] not in B+C+P+Q+N for x in ls):
                out['mixcell'].append((h,show(t)[:90].replace('\n',' ⏎ ')))
    for k,cs in cols.items():
        b=[t for t in cs if B in t]; nb=[t for t in cs if B not in t and len(t.strip())>=8 and not re.fullmatch(r'[—－\-–/／无空×✕✓√?？…（）()\s]*',t)]
        if b and nb: out['mixcol'].append((h,k,len(b),len(nb),show(nb[0])[:50].replace('\n',' ⏎ ')))
for k,v in out.items():
    print('==',k,len(v))
    for x in v[:int(sys.argv[2]) if len(sys.argv)>2 else 200]: print('  ',x)
