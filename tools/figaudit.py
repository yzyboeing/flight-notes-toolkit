# -*- coding: utf-8 -*-
"""1020R1 插图逐张审查：找出 PDF 里每张插图所在页，量图宽 / 图高 / 本页空白，并把这些页拼成缩略图（4 × 3）。
报：图宽 < 版心 35% 且本页下方空白 > 25% 版心高；并排图与表高度差 > 40%。用法：figaudit.py <pdf> <输出目录>"""
import sys, os, re, subprocess, shutil
import pymupdf
pdf, out = sys.argv[1], os.path.abspath(sys.argv[2])
os.makedirs(out, exist_ok=True)
d = pymupdf.open(pdf)
rows, pages = [], []
for pn, p in enumerate(d):
    ims = [b for b in p.get_text('dict')['blocks'] if b['type'] == 1 and (b['bbox'][2] - b['bbox'][0]) > 40]
    if not ims: continue
    W, H = p.rect.width, p.rect.height
    txt = p.get_text('blocks')
    body = [b for b in txt if b[1] > 40 and b[3] < H - 40]
    bottom = max([b[3] for b in body] + [im['bbox'][3] for im in ims])
    caps = [b[4].strip().replace('\n', '') for b in txt if re.match(r'\s*图 \d', b[4])]
    tabs = p.find_tables().tables
    for im in ims:
        x0, y0, x1, y1 = im['bbox']
        side = [t for t in tabs if t.bbox[1] < y1 and t.bbox[3] > y0 and (t.bbox[0] > x1 - 5 or t.bbox[2] < x0 + 5)]
        flag = []
        if (x1 - x0) < 0.35 * (W - 72) and (H - 50 - bottom) > 0.25 * (H - 90): flag.append('小图且下方空白多')
        if side:
            th = max(t.bbox[3] for t in side) - min(t.bbox[1] for t in side)
            if th and abs(th - (y1 - y0)) / th > 0.4: flag.append('并排高度差 %.0f%%' % (100 * ((y1 - y0) - th) / th))
            gap = (min(t.bbox[0] for t in side) - x1) if min(t.bbox[0] for t in side) > x1 else (x0 - max(t.bbox[2] for t in side))
            if gap > 40: flag.append('图表间空 %.0fpt' % gap)
        rows.append((pn + 1, round(x1 - x0), round(y1 - y0), round(H - 50 - bottom), 'side' if side else '', ' '.join(flag), caps[0][:30] if caps else ''))
    pages.append(pn + 1)
with open(os.path.join(out, 'figs.tsv'), 'w') as f:
    for r in rows: f.write('\t'.join(map(str, r)) + '\n')
for r in rows:
    if r[5]: print(*r, sep='\t')
print('插图页', len(pages), '张图', len(rows), '有标记', sum(1 for r in rows if r[5]))
tmp = os.path.join(out, '_p'); shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
for pg in pages: subprocess.run(['pdftoppm', '-f', str(pg), '-l', str(pg), '-r', '45', '-png', pdf, os.path.join(tmp, 'p%04d' % pg)], check=True)
import glob
pngs = sorted(glob.glob(os.path.join(tmp, '*.png')))
CH = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
for k in range(0, len(pngs), 12):
    grp = pngs[k:k + 12]
    cells = ''.join('<div class="c"><img src="file://%s"><span>%s</span></div>' % (p, int(os.path.basename(p)[1:5])) for p in grp)
    html = ('<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0;background:#888;font:bold 18px Helvetica}'
            '.g{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;padding:6px}.c{position:relative;background:#fff}'
            '.c img{width:100%;display:block;outline:1px solid #555}.c span{position:absolute;right:4px;bottom:2px;color:#c00}</style></head>'
            '<body><div class="g">' + cells + '</div></body></html>')
    hp = os.path.join(out, 'f%03d.html' % (k // 12)); open(hp, 'w').write(html)
    subprocess.run([CH, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--allow-file-access-from-files',
                    '--window-size=2400,1300', '--screenshot=' + hp[:-5] + '.png', 'file://' + hp], capture_output=True)
print('拼版', (len(pngs) + 11) // 12, '张')
