# -*- coding: utf-8 -*-
"""找出 $变量 紧跟非 ASCII 字符的地方——bash 3.2（macOS 自带）会把后面的字节
   并进变量名，于是报 unbound variable。必须写成 ${变量}。"""
import re, sys, io, os
tot = 0
for p in sys.argv[1:]:
    if not os.path.isfile(p): continue
    try: s = io.open(p, encoding='utf-8').read()
    except Exception: continue
    bad = []
    for m in re.finditer(r'\$[A-Za-z_][A-Za-z0-9_]*', s):
        nxt = s[m.end():m.end()+1]
        if nxt and ord(nxt) > 127:
            bad.append((s[:m.start()].count('\n')+1, m.group(0)+nxt))
    if bad:
        print('  !! %s' % p)
        for ln, t in bad: print('       行 %-4d %s' % (ln, t)); 
        tot += len(bad)
print('共 %d 处' % tot)
