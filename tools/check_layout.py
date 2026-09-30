#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_layout.py —— 排版与整理规则总检查器（SD-79，2026-09-30 用户）
用法（任意目录）：
    python3 ~/flight-repos/pub/tools/check_layout.py [--repo ~/flight-repos/gh-private] [--no-src] [--out 报告.md]
一次扫完所有能机器判定的规则，按「错误 / 建议」两级列出，每条带规则号（对应 pub/prompt/layout-checklist.md）
和页码。错误 = 违反已定规则，必须改；建议 = 可能需要优化，由 AI 看成品页后按清单里的语义标准裁定，
不同意要写明理由。退出码：有错误为 1，否则为 0。只读，不改任何文件。

规则来源：standing-decisions.md SD-51 / SD-66 / SD-71～SD-79，AI交接/04 经验，用户历次排版反馈。
依赖：PyMuPDF；读 build/ 下的成品 PDF，所以先跑 `./sync.sh --full --no-push "…"` 再检查。"""
import sys, os, re, glob, subprocess, collections
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
REPO = os.path.abspath(os.path.expanduser(arg('--repo', '~/flight-repos/gh-private')))
T = os.path.dirname(os.path.abspath(__file__))
BOOK = arg('--book') or os.path.join(REPO, 'build', 'B737机型理论知识笔记.pdf')
QREF = arg('--qref') or os.path.join(REPO, 'build', 'B737机型理论基础知识速查.pdf')
HEADER = 'B737-NG / B737-8 机型理论知识笔记'
def git_cfg(k):
    return subprocess.run(['git', '-C', REPO, 'config', '--get', k], capture_output=True, text=True).stdout.strip()
EDITION, NOTICE, SIGN = git_cfg('notes.docEdition'), git_cfg('notes.docNotice'), git_cfg('notes.docPrefaceSignature')

ERR, SUG = [], []
def err(rule, msg): ERR.append((rule, msg))
def sug(rule, msg): SUG.append((rule, msg))
nosp = lambda s: re.sub(r'\s+', '', s)
def vis(s):   # 视觉宽度：汉字 2，西文 1
    return sum(2 if re.match(r'[⺀-鿿＀-￯]', ch) else 1 for ch in s)

# ---------- S 源头校验（沿用三件套） ----------
if '--no-src' not in sys.argv:
    for rule, tool, extra in (('S1', 'check_src.py', ['--quiet']), ('S2', 'check_blocks.py', []), ('S3', 'check_quickref.py', [])):
        r = subprocess.run([sys.executable, os.path.join(T, tool), '--src', 'notes_src'] + extra, cwd=REPO, capture_output=True, text=True)
        if r.returncode:
            tail = [l for l in (r.stdout + r.stderr).strip().splitlines() if l.strip()][-6:]
            err(rule, '%s 未通过：\n      ' % tool + '\n      '.join(tail))
    if not os.path.isdir(os.path.join(REPO, 'notes_src', '3 运行手册')):
        err('S4', '第三章文件夹应为 notes_src/3 运行手册/（SD-73）')

# ---------- F 成品文件 ----------
for p in (BOOK, QREF, BOOK[:-4] + '.docx', QREF[:-4] + '.docx'):
    if not os.path.exists(p): err('F1', '缺成品：' + os.path.relpath(p, REPO))
for p in glob.glob(os.path.join(REPO, 'build', '*竖版*')):
    err('F2', '不应再有竖版成品（SD-71）：' + os.path.relpath(p, REPO))
src_m = max((os.path.getmtime(f) for f in glob.glob(os.path.join(REPO, 'notes_src', '*', '*.md'))), default=0)
for p in (BOOK, QREF):
    if os.path.exists(p) and os.path.getmtime(p) < src_m:
        err('F3', '%s 比 notes_src 旧，先重新 sync 再检查' % os.path.basename(p))
if not (os.path.exists(BOOK) and os.path.exists(QREF)):
    print('成品不全，先跑 sync.sh --full'); ERR and [print(' ', r, m) for r, m in ERR]; sys.exit(1)

# ---------- 逐页分析（两本都查） ----------
def body_lines(pg):
    """本页正文的视觉行：[(x0, y0, x1, y1, text)]，去掉页眉页脚。"""
    out = []
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans'])
            if t.strip(): out.append((*l['bbox'], t))
    return out

def cell_lines(lines, bb):
    """落在格子里的视觉行（按基线聚类），返回 [(x0, x1, text)]。"""
    got = [l for l in lines if bb[0] - 1 <= (l[0] + l[2]) / 2 <= bb[2] + 1 and bb[1] - 1 <= (l[1] + l[3]) / 2 <= bb[3] + 1]
    rows = collections.OrderedDict()
    for l in sorted(got, key=lambda z: (round(z[1] / 3), z[0])):
        k = round(((l[1] + l[3]) / 2) / 3)
        r = rows.setdefault(k, [l[0], l[2], ''])
        r[0] = min(r[0], l[0]); r[1] = max(r[1], l[2]); r[2] += l[4]
    return [tuple(v) for v in rows.values()]

PAD = 12   # 单元格左右内边距合计（pt）
def scan(pdf, name, header):
    d = pymupdf.open(pdf)
    W, H = d[0].rect.width, d[0].rect.height
    # L1 只出横版
    port = [i + 1 for i, p in enumerate(d) if p.rect.width < p.rect.height]
    if port: err('L1', '%s 有竖版页：%s' % (name, port[:10]))
    # L2 字体统一（SD-76）
    bad = collections.Counter()
    for p in d:
        for f in p.get_fonts():
            base = f[3].split('+')[-1]
            if 'Songti' not in base: bad[base] += 1
    for f, n in bad.items(): err('L2', '%s 出现非宋体字体 %s（%d 页）——字体须统一 Songti SC（SD-76）' % (name, f, n))
    # L3 书签栏（SD-77）
    if not d.get_toc(): err('L3', name + ' 没有书签')
    if 'UseOutlines' not in (d.pdf_catalog() and d.xref_get_key(d.pdf_catalog(), 'PageMode')[1] or ''):
        err('L3', name + ' 打开时未展开书签栏（PageMode 应为 UseOutlines）')
    # C 封面（SD-73 / SD-74）
    cov = d[0].get_text('blocks')
    ct = nosp(''.join(b[4] for b in cov))
    for want, what in ((EDITION and '版本号' + EDITION, '「版本号 %s」' % EDITION), (nosp(NOTICE), '特别提示全文')):
        if want and want not in ct: err('C1', '%s 封面缺%s' % (name, what))
    if EDITION and not any(re.search(r'版本号 %s(?!\d)' % EDITION, b[4]) for b in cov):
        err('C1', '%s 封面「版本号」与数字之间应只有一个半角空格（SD-74）' % name)
    yv = [b[1] for b in cov if '版本号' in b[4]]; yn = [b[1] for b in cov if '特别提示' in b[4]]
    if yv and yn and not yn[0] < yv[0]: err('C2', '%s 封面顺序应为：特别提示在上、版本号在下（SD-74）' % name)
    if re.search(r'版次|第\s*\d+\s*版', ''.join(b[4] for b in cov)): err('C1', name + ' 封面不写「版次」，直接写版本号（SD-73）')
    # 页眉（SD-73；单册不设页眉）
    miss = []
    for i in range(1, len(d)):
        tops = sorted((b for b in d[i].get_text('blocks') if b[4].strip()), key=lambda b: b[1])
        has = bool(tops) and tops[0][3] < 0.08 * H and nosp(tops[0][4]) == nosp(header) if header else False
        if header and not has: miss.append(i + 1)
        if not header and tops and tops[0][3] < 0.08 * H and '机型' in tops[0][4] and '第' not in tops[0][4]:
            miss.append(i + 1)
    if miss: err('L4', '%s 页眉不符（应为「%s」）：第 %s 页' % (name, header or '无页眉', miss[:12]))
    # 版心 / 页边距（SD-73：720 DXA = 36pt）
    over = []
    for i, p in enumerate(d):
        for l in body_lines(p):
            if l[0] < 36 - 4 or l[2] > W - 36 + 4: over.append(i + 1); break
    if over: sug('L5', '%s 文字伸进右页边距（多为长英文串撑出行尾），看是否需要断开：第 %s 页' % (name, sorted(set(over))[:12]))
    # 表格逐格检查
    CW = W - 72
    for i, p in enumerate(d):
        try: tabs = p.find_tables().tables
        except Exception: continue
        lines = body_lines(p)
        for t in tabs:
            if t.bbox[3] - t.bbox[1] < 8 or t.bbox[1] < 0.08 * H: continue
            tw = t.bbox[2] - t.bbox[0]
            # T5 表内字号底线 8pt（SD-80）
            small = [sp['size'] for b in p.get_text('dict', clip=t.bbox)['blocks'] for l in b.get('lines', []) for sp in l['spans']
                     if sp['text'].strip() and sp['size'] < 7.9]
            if small: err('T5', '%s 第 %d 页：表内有 %.1fpt 的字，低于 8pt 底线' % (name, i + 1, min(small)))
            if tw > CW + 4: err('T1', '%s 第 %d 页：表格宽 %.0fpt 超过版心 %.0fpt' % (name, i + 1, tw, CW))
            rows = [[c for c in r.cells] for r in t.rows]
            if not rows: continue
            ncol = max(len(r) for r in rows)
            info = [[(c, cell_lines(lines, c)) if c else None for c in r] for r in rows]
            # 每列：宽度与最长一行的占用
            colw, colused = [0] * ncol, [0] * ncol
            for r in info:
                for k, x in enumerate(r):
                    if not x: continue
                    c, ls = x
                    if k + 1 < len(r) and r[k + 1] is None: continue      # 跨列格不代表单列宽
                    colw[k] = max(colw[k], c[2] - c[0])
                    colused[k] = max(colused[k], max((l[1] - l[0] for l in ls), default=0))
            spare = [max(0, colw[k] - colused[k] - PAD) for k in range(ncol)]
            # T2 有剩余宽度却折行（SD-72）：表头、以及 20 个字宽以内的短格
            for ri, r in enumerate(info):
                for k, x in enumerate(r):
                    if not x: continue
                    c, ls = x
                    if len(ls) < 2: continue
                    txt = ''.join(l[2] for l in ls).strip()
                    if ri > 0 and vis(txt) > 20: continue
                    if re.match(r'[•–▪①-⑳]', txt) or '•' in txt: continue
                    # 自然折行时首行会接近撑满格宽；首行明显短于格宽说明是原文 <br> 主动分行（SD-72 不管）
                    if (ls[0][1] - ls[0][0]) < (c[2] - c[0]) - PAD - 18: continue
                    need = sum(l[1] - l[0] for l in ls) + PAD - (c[2] - c[0])
                    have = sum(spare[j] for j in range(ncol) if j != k) + max(0, CW - tw)
                    if 0 < need and need + 6 <= have:
                        sug('T2', '%s 第 %d 页：%s「%s」折成 %d 行，表内还有约 %.0fpt 空余可让它一行放下' % (
                            name, i + 1, '表头' if ri == 0 else '短格', txt[:24], len(ls), have))
            # T3 整列都是一行短内容却没居中（SD-75 / SD-78），第 2 列起、至少 2 格
            for k in range(1, ncol):
                cells = [r[k] for r in info[1:] if k < len(r) and r[k] and (k + 1 >= len(r) or r[k + 1] is not None)]
                cells = [x for x in cells if x[1]]
                if len(cells) < 2 or any(len(x[1]) != 1 for x in cells): continue
                left = [x for x in cells if abs(((x[1][0][0] + x[1][0][1]) / 2) - ((x[0][0] + x[0][2]) / 2)) > 4
                        and (x[0][2] - x[0][0]) - (x[1][0][1] - x[1][0][0]) > PAD + 8]
                if len(left) == len(cells):
                    sug('T3', '%s 第 %d 页：第 %d 列每格都是一行短内容却左对齐（如「%s」），考虑整列居中' % (
                        name, i + 1, k + 1, left[0][1][0][2].strip()[:20]))
            # T4 并列长句未分条（SD-78）：一格内 ≥2 个「；」、各分句 ≥ 20 字宽、却没有 •
            for r in info[1:]:
                for x in r:
                    if not x: continue
                    txt = ''.join(l[2] for l in x[1]).strip()
                    if re.search(r'[•–▪①-⑳]', txt) or re.match(r'(注|解释|出处|公司差异)[：:]', txt): continue
                    parts = [s for s in re.split(r'；', txt) if s.strip()]
                    if len(parts) >= 3 and all(vis(s) >= 20 for s in parts[:-1]):
                        sug('T4', '%s 第 %d 页：「%s…」含 %d 个并列长分句，考虑按语义分条加「•」' % (name, i + 1, txt[:20], len(parts)))
    return d

book = scan(BOOK, '全书', HEADER)
scan(QREF, '单册', '机型基础知识速查')   # 单册页眉为册名（DOC_HEADER 留空时渲染器的默认）

# P 前言（SD-73 / SD-74）与总目录
pre = next((i for i in range(1, 5) if nosp(book[i].get_text()).find('前言') >= 0), None)
if pre is None: err('P1', '全书第 2～5 页找不到前言')
else:
    t = book[pre].get_text()
    if SIGN and SIGN not in t: err('P1', '前言缺署名 %s（SD-74）' % SIGN)
    if re.search(r'以\s*(SOP|FCOM)[^。]{0,6}为准', t): err('P1', '前言不应再有「以 SOP / FCOM 为准」一句（SD-73）')
toc = ''.join(book[i].get_text() for i in range(1, 6))
if '运行规范' in re.sub(r'运行规范\s*C\d+', '', toc): err('P2', '总目录仍出现「运行规范」，第三章名应为「运行手册」（SD-73）')
if '速查主题清单' not in toc: err('P2', '总目录缺「速查主题清单」入口（SD-73）')

# V 成品通用校验（verify.py：空白页、标签泄漏、异常项目符号、front matter）
for pdf in (BOOK, QREF):
    r = subprocess.run([sys.executable, os.path.join(T, 'verify.py'), pdf], capture_output=True, text=True)
    if r.returncode: err('V1', '%s verify.py 未通过：%s' % (os.path.basename(pdf), ' / '.join(l for l in r.stdout.splitlines() if '[]' not in l)))

# B 断表 / 孤行 / 标题孤立（SD-71 / SD-77 / SD-79）
for pdf in (BOOK, QREF):
    r = subprocess.run([sys.executable, os.path.join(T, 'check_splits.py'), pdf], capture_output=True, text=True)
    for l in r.stdout.splitlines()[1:]:
        l = l.strip()
        rule = 'B1' if '本可整页' in l else 'B2' if '孤行' in l else 'B3'
        err(rule, '%s %s' % ('全书' if pdf == BOOK else '单册', l))

# Z 压缩表清单（SD-80）：fit_fix.py 为「略超一页」的表选的压缩级别，供人工抽看是否仍清晰易读
for kf in sorted(glob.glob(os.path.join(REPO, 'build', 'keep_force_*.txt'))):
    book_name = os.path.basename(kf)[11:-4]
    for l in open(kf, encoding='utf-8'):
        l = l.strip()
        if l.startswith('~'):
            sug('Z1', '%s：「%s…」按第 %s 级压缩后整表同页（%s）——抽看是否清晰' % (
                book_name, l[3:23], l[1], {'1': '9pt、行距收紧', '2': '8.5pt', '3': '8pt'}.get(l[1], '?')))
        elif l.startswith('!'):
            sug('Z2', '%s：「%s…」压到 8pt 仍放不下，照常分页——看能否精简内容或拆表' % (book_name, l[1:21]))

# ---------- 报告 ----------
out = ['# 排版与规则检查报告', '', '全书 %d 页。错误 %d 条（必须改），建议 %d 条（AI 看成品页后按 layout-checklist.md 裁定）。' % (len(book), len(ERR), len(SUG)), '']
for title, items in (('## 错误', ERR), ('## 建议', SUG)):
    out.append(title)
    if not items: out.append('（无）')
    for rule, msg in sorted(items, key=lambda x: x[0]): out.append('- [%s] %s' % (rule, msg))
    out.append('')
txt = '\n'.join(out)
print(txt)
if arg('--out'): open(arg('--out'), 'w', encoding='utf-8').write(txt + '\n')
sys.exit(1 if ERR else 0)
