#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""book_pdf.py —— 给全书 PDF（iPad 竖版）加三级书签：章 → 节 → 块（第 4、5 章为条目）

用法：python3 tools/book_pdf.py <全书.pdf> [--src notes_src] [--out 输出.pdf]
章名取 MANIFEST 里各章目录名（「第一章　系统理论」等在 PDF 中的标题行），节取各节源文件名，
块取节内「### A　块名」。按页面文字行精确匹配定位，找不到的项跳过并报数。
"""
import io, os, re, sys, glob
try:
    import pymupdf
except ImportError:  # PyMuPDF 旧版模块名
    import fitz as pymupdf

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
PDF = sys.argv[1]
SRC = os.path.abspath(_arg('--src', 'notes_src'))
OUT = _arg('--out', PDF[:-4] + '_书签.pdf')
CN = '零一二三四五'

d = pymupdf.open(PDF)
lines = [[re.sub(r'\s+', ' ', l.replace('　', ' ')).strip() for l in p.get_text().split('\n')] for p in d]
def find(text, start):
    key = re.sub(r'\s+', ' ', text.replace('　', ' ')).strip()
    for i in range(start, len(d)):
        if key in lines[i]: return i
    return None

toc, miss, cur = [], 0, 2
chapters = sorted({os.path.basename(os.path.dirname(f)) for f in glob.glob(os.path.join(SRC, '[0-5] *', '*.md'))})
for ch in chapters:
    n = int(ch.split(' ')[0]); name = ch.split(' ', 1)[1]
    pg = find('第%s章 %s' % (CN[n], name), cur)
    if pg is None: miss += 1; continue
    toc.append([1, '第%s章　%s' % (CN[n], name), pg + 1]); cur = pg
    files = sorted((f for f in glob.glob(os.path.join(SRC, ch, '*.md')) if not os.path.basename(f).startswith('_')),
                   key=lambda f: [int(x) for x in re.findall(r'\d+', os.path.basename(f).split(' ')[0])])
    for f in files:
        sid, sname = os.path.basename(f)[:-3].split(' ', 1)
        if n == 0: continue                       # 速查区不分节
        src = io.open(f, encoding='utf-8').read()
        h1 = re.search(r'(?m)^# (.+)$', src)          # 以源文件 H1 为准（文件名里的「·」在标题里可能写作「 / 」）
        title = h1.group(1).strip() if h1 else '%s %s' % (sid, sname)
        if not title.startswith(sid): title = '%s %s' % (sid, title)
        spg = find(title, cur)
        if spg is None: miss += 1; print('  找不到：', title); continue
        toc.append([2, re.sub(r'^(\S+)\s+', r'\1　', title.replace('　', ' ')), spg + 1]); cur = spg
        for bl, bn in re.findall(r'(?m)^### ([A-Z])　(.+)$', src):
            bpg = find('%s %s' % (bl, bn.strip()), cur)
            if bpg is not None: toc.append([3, '%s　%s' % (bl, bn.strip()), bpg + 1])
        for no, tt in re.findall(r'(?m)^### (\d+)\. (.+)$', src):     # 第 4、5 章不分块：第三级直接挂条目（SD-30）
            tt = tt.split('　【')[0].strip()                     # 标题后的出处注长了会折行，只用正文标题定位
            ipg = find('%s. %s' % (no, tt), cur)
            if ipg is None: ipg = next((k for k in range(cur, len(lines)) if any(l.startswith('%s. %s' % (no, tt[:12])) for l in lines[k])), None)
            if ipg is not None: toc.append([3, '%s. %s' % (no, tt), ipg + 1]); cur = ipg
d.set_toc(toc)
io.open(OUT, 'wb').write(d.tobytes(garbage=3, deflate=True))
print('写入 %s：%d 页，书签 %d 个（找不到 %d 个章 / 节标题）' % (os.path.basename(OUT), len(d), len(toc), miss))
