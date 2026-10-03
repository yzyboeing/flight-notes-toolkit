#!/usr/bin/env python3
"""生成《笔记结构索引》（给噜噜判断知识点放哪里用，SD-111）。

用法：python3 notes_index.py <notes_src 目录> <全书 PDF> <输出 .md> [基线号]
内容：每章 → 每节（一句简介）→ 块与条目（标题 + 全书页码）；速查区每条 → 正文「详见」位置。
页码按条目标题在全书 PDF 中首次出现的页（找不到写「—」）。每次出新基线时重跑。
"""
import os, re, sys, glob

src, pdf, out = sys.argv[1], sys.argv[2], sys.argv[3]
base = sys.argv[4] if len(sys.argv) > 4 else ''

pages = []
try:
    import fitz
    d = fitz.open(pdf)
    pages = [re.sub(r'\s+', '', p.get_text()) for p in d]
except Exception as e:
    print('!! PDF 读取失败，页码留空：', e)

def page_of(title, start=0):
    key = re.sub(r'\s+', '', title)
    for i in range(start, len(pages)):
        if key in pages[i]:
            return i + 1
    return None

strip = lambda t: re.sub(r'<[^>]+>', '', t).replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&').strip()
L = [f'# 笔记结构索引（基线 {base}）', '',
     '> 自动生成，每次出新基线时更新。用途：判断新知识点放在哪一章、哪一节、哪一块，以及速查区和正文的对应关系。',
     '> 页码是全书 PDF 页码。块（A、B、C）下的条目（A-1、A-2）是节内地址；第四、五章不分块，条目写「N. 标题」。', '']

chap_dirs = sorted([p for p in glob.glob(os.path.join(src, '[0-9] *')) if os.path.isdir(p)])
last = 0
for cd in chap_dirs:
    cname = os.path.basename(cd)
    files = [f for f in glob.glob(os.path.join(cd, '*.md')) if not os.path.basename(f).startswith('_')]
    key = lambda f: [int(x) for x in re.findall(r'^\d+(?:\.\d+)?', os.path.basename(f))[0].split('.')]
    files.sort(key=key)
    L += [f'## {cname}', '']
    for f in files:
        s = open(f, encoding='utf-8').read()
        body = s.split('\n---\n', 1)[1] if s.startswith('---') else s
        h1 = re.search(r'^# (.+)$', body, re.M)
        sec = h1.group(1).strip() if h1 else os.path.basename(f)[:-3]
        intro = ''
        m = re.search(r'^# .+\n+([^#<\n].+)$', body, re.M)
        if m: intro = strip(m.group(1))[:80]
        quick = cname.startswith('0')
        if not quick:
            L += [f'### {sec}' + (f'（p{page_of(sec.replace("　", ""), max(0, last - 1)) or "—"}）' if pages else ''), '']
            if intro: L += [f'{intro}', '']
        heads = re.findall(r'^###\s+(.+)$', body, re.M)
        rows = []
        for h in heads:
            h = strip(h)
            if h == '块索引': continue
            if quick and not re.match(r'^\d+\.', h):
                continue
            p = page_of(h.replace("　", ""), max(0, last - 1)) if pages else None
            if p: last = p
            if quick:
                # 速查条目的「详见」
                i = body.find('### ' + h) if ('### ' + h) in body else body.find(h)
                seg = body[i:i + 400]
                dm = re.search(r'<!--\s*详见\s*(.+?)-->', seg)
                det = ''
                if dm:
                    det = '；'.join(x.split('|')[-1] for x in re.findall(r'\[\[([^\]]+)\]\]', dm.group(1)))
                rows.append(f'| {h} | {det or "—"} | {p or "—"} |')
            else:
                rows.append(f'| {h} | {p or "—"} |')
        if quick:
            # 速查区按「## 分组」输出
            groups = re.split(r'^## (.+)$', body, flags=re.M)
            L += ['| 速查条目 | 详见（正文位置） | 全书页 |', '|---|---|---|'] + rows + ['']
        else:
            L += ['| 条目 | 全书页 |', '|---|---|'] + rows + ['']
open(out, 'w', encoding='utf-8').write('\n'.join(L))
print('written', out, len(L), 'lines')
