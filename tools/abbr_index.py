#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""abbr_index.py —— 附录「缩写与术语」数据（SD-209 批 2 L3）：
从两本源文件的「注：缩写：英文全称（中文）」（或「，中文」）收集缩写，合并《缩写术语表.json》，按字母序写 build/缩写表.json，
生成器（DOC_ABBR）在两本末尾排成「缩写 ｜ 英文全称 ｜ 中文」附录。sync 全量时自动跑。
用法：python3 abbr_index.py --repo <gh-private> [--out build/缩写表.json]"""
import re, io, os, glob, json, html, sys, collections
REPO = os.path.abspath(os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else '.'))
OUT = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else os.path.join(REPO, 'build', '缩写表.json')
P = lambda s: html.unescape(re.sub(r'<[^>]+>', '', s))
found = collections.OrderedDict()
def add(ab, en, zh):
    ab = ab.strip(); en = en.strip(); zh = zh.strip().rstrip('。；')
    if ab not in found or (zh and not found[ab][1]): found[ab] = (en, zh)
files = sorted(glob.glob(os.path.join(REPO, 'notes_src', '[1-5]*', '[0-9]*.md'))) + [os.path.join(REPO, '速查', '速查源.md')]
for f in files:
    for l in io.open(f, encoding='utf-8').read().split('\n'):
        if not l.startswith('注：'): continue
        for seg in re.split(r'[；。]', P(l)):
            m = re.match(r'\s*(?:注：)?(?:缩写：)?([A-Z][A-Za-z/\-]{1,12})：([A-Z][A-Za-z][A-Za-z /\-\'’]*?)\s*[，,]?\s*(（([^）]*)）|([一-鿿][^（）]*))?\s*$', seg)
            if m: add(m.group(1), m.group(2), m.group(4) or m.group(5) or '')
tf = os.path.join(REPO, '缩写术语表.json')
if os.path.exists(tf):
    try:
        d = json.load(io.open(tf, encoding='utf-8'))
        for ab, v in (d.get('条目', d) if isinstance(d, dict) else {}).items():
            if isinstance(v, dict): add(ab, v.get('英文', v.get('en', '')), v.get('中文', v.get('zh', '')))
            elif isinstance(v, str): add(ab, '', v)
    except Exception as e: print('缩写术语表.json 读取跳过：%s' % e)
rows = [{'abbr': k, 'en': v[0], 'zh': v[1]} for k, v in sorted(found.items(), key=lambda kv: kv[0].upper())]
os.makedirs(os.path.dirname(OUT), exist_ok=True)
io.open(OUT, 'w', encoding='utf-8').write(json.dumps(rows, ensure_ascii=False, indent=1) + '\n')
print('缩写表：%d 条（缺中文 %d）→ %s' % (len(rows), sum(1 for r in rows if not r['zh']), OUT))
