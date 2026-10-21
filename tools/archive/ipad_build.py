#!/usr/bin/env python3
"""ipad_build.py —— 速查区 iPad 竖版：排版预检 + 生成（SD-23 / SD-24）

用法：python3 ../pub/tools/ipad_build.py build/mod0.md build/名称   （在 gh-private 下运行）
流程：竖版 docx → LibreOffice 转 PDF → 查出被拆到两页的条目 → 给这些条目标题加「另起一页」
      → 重排，最多 4 轮；一页放不下的条目（本来就要跨页）不加。最后交给 quickref_pdf.py 加书签和索引页。
"""
import os, re, sys, subprocess, glob, io
try:
    import pymupdf
except ImportError:  # PyMuPDF 旧版模块名
    import fitz as pymupdf

MD, BASE = sys.argv[1], sys.argv[2]
TOOLS = os.path.dirname(os.path.abspath(__file__))
src = io.open(glob.glob('notes_src/0 *' + '/0 *.md')[0], encoding='utf-8').read()
BLOCKS = set(m.strip() for m in re.findall(r'(?m)^## (.+)$', src))
TITLE = re.compile(r'^(\d{1,3})\. \S')

def build(breaks, tag):
    env = dict(os.environ, DOC_PORTRAIT='1', BREAK_BEFORE=','.join(sorted(breaks, key=int)))
    docx = '%s_%s.docx' % (BASE, tag)
    subprocess.run(['node', os.path.join(TOOLS, 'build_docx.js'), MD, docx], env=env, check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(docx) or '.', docx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=170)
    return docx[:-5] + '.pdf'

def split_items(pdf):
    """返回 {条目号: 是否一页放得下}：条目内容越过了页边界的条目"""
    d = pymupdf.open(pdf); res = {}; cur = None; cur_top = False; started = False
    for p in d:
        lines = [l.strip() for l in p.get_text().split('\n') if l.strip() and not re.match(r'^第\s*\d+\s*页$', l.strip())]
        if not lines: continue
        first = lines[0]
        if started and cur and not TITLE.match(first) and first not in BLOCKS:
            res[cur] = not cur_top            # 标题已在页顶还跨页 → 一页放不下
        for i, l in enumerate(lines):
            if l in BLOCKS: cur = None          # 进入新块（如末尾「按物理量索引」），不再算上一条
            m = TITLE.match(l)
            if m:
                started = True; cur = m.group(1); cur_top = (i == 0 or (i == 1 and lines[0] in BLOCKS))
    return res

breaks, giveup = set(), set()
for rnd in range(1, 5):
    pdf = build(breaks, 'r%d' % rnd)
    s = split_items(pdf)
    new = {n for n, fits in s.items() if fits and n not in breaks}
    giveup |= {n for n, fits in s.items() if not fits}
    print('第 %d 轮：跨页条目 %s；新加另起一页 %s' % (rnd, sorted(s, key=int), sorted(new, key=int)))
    breaks -= giveup                         # 一页放不下的，另起一页也没用
    if not new: break
    breaks |= new
print('另起一页的条目：', ','.join(sorted(breaks, key=int)) or '无', '｜一页放不下而跨页：', ','.join(sorted(giveup, key=int)) or '无')
print('PDF:', pdf)
