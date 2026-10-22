#!/usr/bin/env python3
"""contact_sheet.py —— 出基线前整本缩略图过目（SD-221，2026-10-21 用户：「以后不要再出现这种情况，应该是你自己审核的时候就发现并优化」）。
起因：速查 1021R1 第 15 页只剩页脚的空白页，检查器漏报、交付前也没整本看过，被用户先看到。

用法：python3 contact_sheet.py --repo <gh-private>            # 两本 PDF 每页缩成小图拼成总览，写 build/整本过目/；列出存疑页
      python3 contact_sheet.py --repo <gh-private> --done     # 逐张看完、存疑页处理完后登记（release.py 前置检查要求登记与当前成品一致）
总览：每张 8 × 5 = 40 页，页码标在缩略图左上角。存疑页（自动列出，须逐页放大看）：
  · 正文为空（只剩页眉页脚）；· 正文墨迹高度不到版心 15%（半空页：多半是分页留白或孤行）。"""
import sys, os, re, json, hashlib, datetime
import pymupdf

REPO = sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else os.getcwd()
B = os.path.join(REPO, 'build')
OUT = os.path.join(B, '整本过目')
REC = os.path.join(B, '整本过目.json')
STEMS = ('B737机型理论知识笔记', 'B737机型理论知识速查')
COLS, ROWS, TW = 8, 5, 200   # 每张 8 × 5 页，缩略图宽 200px


def sha(p):
    h = hashlib.sha256()
    with open(p, 'rb') as f:
        for blk in iter(lambda: f.read(1 << 20), b''): h.update(blk)
    return h.hexdigest()[:16]


def edition():
    import subprocess
    return subprocess.run(['git', 'config', '-f', os.path.join(REPO, '成品设置.gitconfig'), '--get', 'notes.docEdition'],
                          capture_output=True, text=True).stdout.strip()


if '--done' in sys.argv:
    rec = {'版本': edition(), '日期': datetime.date.today().isoformat(),
           '成品': {s: sha(os.path.join(B, s + '.pdf')) for s in STEMS}}
    json.dump(rec, open(REC, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('已登记整本过目：%s' % json.dumps(rec, ensure_ascii=False))
    sys.exit(0)

os.makedirs(OUT, exist_ok=True)
for f in os.listdir(OUT):
    if f.endswith('.png'): os.remove(os.path.join(OUT, f))
for s in STEMS:
    d = pymupdf.open(os.path.join(B, s + '.pdf'))
    W, H = d[0].rect.width, d[0].rect.height
    sc = TW / W; th = int(H * sc)
    sus = []
    for i, p in enumerate(d):
        if i == 0: continue   # 封面
        t = p.get_text()
        body = re.sub(r'(?m)^.*｜\s*第\s*\d+\s*页\s*$', '', t)
        body = re.sub(r'第 \d+ 页', '', body).strip()
        big = [x['bbox'] for x in p.get_image_info() if (x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 20000]
        if not body and not big: sus.append((i + 1, '空白页（只剩页眉页脚）')); continue
        bl = [b for b in p.get_text('blocks') if b[4].strip() and not re.search(r'第\s*\d+\s*页', b[4])] + [tuple(x) + ('img',) for x in big]
        ys = [b[3] for b in bl]
        top = min([b[1] for b in bl] + [H])
        if ys and (max(ys) - top) < 0.15 * (H - 100): sus.append((i + 1, '半空页（正文高度 < 15%%，止于 %.0f%%）' % (100 * max(ys) / H)))
    per = COLS * ROWS
    mat = pymupdf.Matrix(sc, sc)
    for k in range(0, len(d), per):
        # 逐页直接渲染小图再拼（show_pdf_page 嵌整页矢量太慢：两本 5 分钟没跑完）
        SW, SH = COLS * (TW + 6) + 6, ROWS * (th + 6) + 6
        big = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, SW, SH), False)
        big.clear_with(215)
        pos = []
        for j, i in enumerate(range(k, min(k + per, len(d)))):
            pm = d[i].get_pixmap(matrix=mat, alpha=False)
            x, y = 6 + (j % COLS) * (TW + 6), 6 + (j // COLS) * (th + 6)
            pm.set_origin(x, y); big.copy(pm, pm.irect); pos.append((x, y, i + 1))
        # 页码标签：拼好的位图放进一页 PDF 再写字、整页渲染一次
        sh = pymupdf.open(); pg = sh.new_page(width=SW, height=SH)
        pg.insert_image(pg.rect, pixmap=big)
        for x, y, n in pos:
            pg.draw_rect(pymupdf.Rect(x, y, x + 26, y + 13), color=None, fill=(0.8, 0, 0))
            pg.insert_text((x + 3, y + 10), str(n), fontsize=9, color=(1, 1, 1))
        pg.get_pixmap(alpha=False).save(os.path.join(OUT, '%s_%03d-%03d.png' % (s[-2:], k + 1, min(k + per, len(d)))))
    print('%s：%d 页 → %d 张总览；存疑 %d 页%s' % (s, len(d), (len(d) + per - 1) // per, len(sus),
          ''.join('\n  · 第 %d 页 %s' % x for x in sus)))
print('总览在 %s，逐张看完、存疑页处理完后跑 --done 登记' % OUT)
