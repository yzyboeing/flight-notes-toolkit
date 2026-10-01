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

# S5 速查区与正文同一张表的表头标注要一致（2026-09-30 用户：「以上的所有更新都同步更新到其它章节了吗」）
#   按速查每条下的「详见 x.y A-n」找到正文条目，列数相同、表头至多差一格的表视为同一张表，col-center / col-bullet / col-plain / col-left 必须相同
def _hdr_tables(txt):
    out = []
    for m in re.finditer(r'<tr class="hdr">(.*?)</tr>', txt):
        cells = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1))
        out.append(([re.sub(r'<[^>]+>|\s', '', c[1]) for c in cells], [(re.search(r'class="([^"]+)"', c[0]) or [None, ''])[1] for c in cells]))
    return out
try:
    _q = open(os.path.join(REPO, 'notes_src', '0 基础知识速查区', '0 基础知识速查区.md'), encoding='utf-8').read()
    _files = {}
    for _f in glob.glob(os.path.join(REPO, 'notes_src', '[1-5]*', '*.md')):
        _m = re.match(r'(\d\.\d+)', os.path.basename(_f))
        if _m: _files[_m.group(1)] = _f
    for _it in re.split(r'(?m)^(?=### \d+\. )', _q):
        _m = re.match(r'### (\d+)\. (.+)', _it)
        if not _m: continue
        for _sec, _addr in re.findall(r'\|(\d\.\d+) ([A-Z]-\d+)\]\]', _it):
            if _sec not in _files: continue
            _s = open(_files[_sec], encoding='utf-8').read()
            _a = re.search(r'(?m)^### ' + re.escape(_addr) + r'[\s\u3000]', _s)
            if not _a: continue
            _nx = re.search(r'(?m)^### ', _s[_a.end():]); _b = _s[_a.start(): _a.end() + (_nx.start() if _nx else len(_s))]
            for _hq, _cq in _hdr_tables(_it):
                for _hb, _cb in _hdr_tables(_b):
                    if len(_hb) == len(_hq) and sum(x == y for x, y in zip(_hb, _hq)) >= max(1, len(_hq) - 1) and (any(_cq) or any(_cb)) and _cq != _cb:
                        err('S5', '速查第 %s 条「%s」与正文 %s %s 同一张表的表头标注不一致：速查 %s，正文 %s' % (_m.group(1), _m.group(2)[:12], _sec, _addr, _cq, _cb))
except Exception as _e:
    sug('S5', '速查 / 正文标注对照未能完成：%s' % _e)

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
    # SD-97：页脚一行（章名 / 节名 / 页码）整行去掉——以页码所在行为准，同一行及以下都算页脚
    fy = max((l[1] for l in out if re.search(r'第\s*\d+\s*页\s*$', l[4])), default=None)
    return [l for l in out if fy is None or l[1] < fy - 2]

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
# 源文件里 <br> 主动换行的位置（T2 / T8 不把作者有意的换行当成问题）：build/book.md 去标签后只留字母数字，<br> 记为「|」
_bk = os.path.join(REPO, 'build', 'book.md')
BRTEXT = ''
if os.path.exists(_bk):
    _t = re.sub(r'<br\s*/?>', '\x01', open(_bk, encoding='utf-8').read())
    _t = re.sub(r'<[^>]+>', '', _t)
    BRTEXT = ''.join(ch if (ch.isalnum() or ch == '\x01') else '' for ch in _t).replace('\x01', '|')
def explicit_br(first_line):
    k = ''.join(ch for ch in first_line if ch.isalnum())
    return bool(k) and (k + '|') in BRTEXT
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
    # 页脚导航（SD-97：顶端不放页眉；右下角「章名　节名　　第 X 页」，章首页 / 前言 / 总目录只有章名；单册「册名　块名」）
    miss = []
    for i in range(1, len(d)):
        bl = [b for b in d[i].get_text('blocks') if b[4].strip()]
        fy = max((b[1] for b in bl if re.search(r'第\s*\d+\s*页\s*$', b[4])), default=None)
        ftxt = nosp(''.join(b[4] for b in sorted((b for b in bl if fy is not None and b[1] >= fy - 2), key=lambda b: b[0])))
        toptxt = [b for b in bl if b[3] < 0.06 * H]
        has = bool(ftxt) and re.match(header, ftxt) is not None and 'Error' not in ftxt and '§' not in ftxt and not toptxt if header else False
        if header and not has: miss.append(i + 1)
        if not header and tops and tops[0][3] < 0.08 * H and '机型' in tops[0][4] and '第' not in tops[0][4]:
            miss.append(i + 1)
    if miss: err('L4', '%s 页脚导航不符（应以「%s」开头、顶端不放页眉）：第 %s 页' % (name, header or '无', miss[:12]))
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
                    if (ls[0][1] - ls[0][0]) < (c[2] - c[0]) - PAD - 18 or explicit_br(ls[0][2]): continue
                    need = sum(l[1] - l[0] for l in ls) + PAD - (c[2] - c[0])
                    have = sum(spare[j] for j in range(ncol) if j != k) + max(0, CW - tw)
                    if 0 < need and need + 6 <= have:
                        sug('T2', '%s 第 %d 页：%s「%s」折成 %d 行，表内还有约 %.0fpt 空余可让它一行放下' % (
                            name, i + 1, '表头' if ri == 0 else '短格', txt[:24], len(ls), have))
            # T3 整列都是一行短内容却没居中（SD-75 / SD-78），第 2 列起、至少 2 格；序号表（首列 ①②③ / 1 2 3）按 SD-90 左对齐，不查
            first = [''.join(l[2] for l in r[0][1]).strip() for r in info[1:] if r and r[0] and r[0][1]]
            serial = len(first) >= 2 and all(re.match(r'^([①-⑳]|\d{1,2}[.、]?)$', x) for x in first)
            for k in (range(1, ncol) if not serial else []):
                cells = [r[k] for r in info[1:] if k < len(r) and r[k] and (k + 1 >= len(r) or r[k + 1] is not None)]
                cells = [x for x in cells if x[1]]
                if len(cells) < 2 or any(len(x[1]) != 1 for x in cells): continue
                if any(x[1][0][2].lstrip().startswith('•') for x in cells): continue   # 按同列统一规则加点的列本就左齐（表可能跨页，只看到一半）
                left = [x for x in cells if abs(((x[1][0][0] + x[1][0][1]) / 2) - ((x[0][0] + x[0][2]) / 2)) > 4
                        and (x[0][2] - x[0][0]) - (x[1][0][1] - x[1][0][0]) > PAD + 8]
                if len(left) == len(cells):
                    sug('T3', '%s 第 %d 页：第 %d 列每格都是一行短内容却左对齐（如「%s」），考虑整列居中' % (
                        name, i + 1, k + 1, left[0][1][0][2].strip()[:20]))
            # T6 / T7 / T8 同列统一与末行孤字（SD-84，2026-09-30 用户速查区第 15～55 条反馈）
            PH = re.compile(r'^[—－\-–/／无空×✕✓√?？…\s]*$')
            for k in range(ncol):
                cells = [r[k] for r in info[1:] if k < len(r) and r[k] and (k + 1 >= len(r) or r[k + 1] is not None) and r[k][1]]
                cells = [x for x in cells if not PH.match(''.join(l[2] for l in x[1]))]
                if len(cells) < 2: continue
                bul = [x for x in cells if x[1][0][2].lstrip().startswith('•')]
                # 没加点的格只算「句子」：≥ 12 字或以句号结尾；短值（持续 / ≥ 800m）和 ①② 步骤格不算不统一
                def sentence(x):
                    t0 = ''.join(l[2] for l in x[1]).strip()
                    return not re.match(r'[①-⑳]', t0) and '①' not in t0 and (len(re.sub(r'\s', '', t0)) >= 12 or t0.endswith('。'))
                nbs = [x for x in cells if x not in bul and sentence(x)]
                # 取值列（有「持续」「≥ 800m」这类不带句号的短值格）：长的并列格加点、短值居中不加点，是允许的差异（用户第 52 条）
                valcol = any(len(re.sub(r'\s', '', ''.join(l[2] for l in x[1]))) < 12 and not ''.join(l[2] for l in x[1]).strip().endswith('。')
                             for x in cells if x not in bul)
                if k > 0 and bul and nbs and not valcol:
                    other = nbs[0]
                    sug('T6', '%s 第 %d 页：第 %d 列有 %d 格分条加点、%d 格句子没有（如「%s」）——整列加点（col-bullet）或整列不加（col-center）' % (
                        name, i + 1, k + 1, len(bul), len(nbs), other[1][0][2].strip()[:18]))
                def centered(x):
                    cx = (x[0][0] + x[0][2]) / 2
                    return all(abs((l[0] + l[1]) / 2 - cx) <= 4 for l in x[1]) and any((x[0][2] - x[0][0]) - (l[1] - l[0]) > PAD + 8 for l in x[1])
                def lefty(x):
                    return all(l[0] - x[0][0] < PAD + 2 for l in x[1]) and any((x[0][2] - x[0][0]) - (l[1] - l[0]) > PAD + 8 for l in x[1])
                nb = [x for x in cells if x not in bul]
                cen = [x for x in nb if centered(x)]; lef = [x for x in nb if lefty(x) and not centered(x)]
                if cen and lef:
                    sug('T7', '%s 第 %d 页：第 %d 列 %d 格居中、%d 格左对齐（如「%s」）——统一对齐' % (
                        name, i + 1, k + 1, len(cen), len(lef), lef[0][1][0][2].strip()[:18]))
            # T10 引语后的子项被排成同级（2026-09-30 用户，速查第 6 条）：「• ……：」后面紧跟的仍是「•」
            for r in info[1:]:
                for x in r:
                    if not x: continue
                    ls = [l[2].strip() for l in x[1]]
                    for a0, b0 in zip(ls, ls[1:]):
                        if a0.startswith('•') and re.search(r'[：:]$', a0) and b0.startswith('•'):
                            sug('T10', '%s 第 %d 页：「%s」是引语，后面的子项应为「–」，现为「•」' % (name, i + 1, a0[:24])); break
            # T9 两型对照列格式不一致（2026-09-30 全书复审，速查第 120 条）：表头有 737-NG 与 737-8 两列，一列有「•」、另一列同类句子格没有
            try:
                hdrtxt = [''.join(l[2] for l in x[1]) if x else '' for x in info[0]]
                ng = [k for k, h in enumerate(hdrtxt) if re.search(r'737-NG', h) and '737-8' not in h]
                m8 = [k for k, h in enumerate(hdrtxt) if re.search(r'737-8', h) and '737-NG' not in h]
                if ng and m8:
                    def colbul(k):
                        cs = [r[k] for r in info[1:] if k < len(r) and r[k] and r[k][1]]
                        cs = [x for x in cs if len(re.sub(r'\s', '', ''.join(l[2] for l in x[1]))) >= 12]
                        return any(x[1][0][2].lstrip().startswith('•') for x in cs), any(not x[1][0][2].lstrip().startswith('•') for x in cs)
                    a, b = colbul(ng[0]), colbul(m8[0])
                    if (a[0] and not b[0] and b[1]) or (b[0] and not a[0] and a[1]):
                        sug('T9', '%s 第 %d 页：737-NG / 737-8 对照两列一列分条加点、另一列没有——两列格式应一致（表头标 col-bullet 或 col-center）' % (name, i + 1))
            except Exception: pass
            for r in info:
                for x in r:
                    if not x or len(x[1]) < 2 or len(x[1]) > 3 or x[1][0][2].lstrip().startswith('•'): continue   # 只管短格；长段落末行一两个字属正常
                    last = re.sub(r'[\s，。；：、（）()「」.,;:]', '', x[1][-1][2])
                    prev = x[1][-2]   # 上一行接近撑满格宽才是自然折行；原文 <br> 主动换行（如「（被动）」）不算
                    if (prev[1] - prev[0]) < (x[0][2] - x[0][0]) - PAD - 18 or re.match(r'\s*[（(]', x[1][-1][2]) or explicit_br(x[1][-2][2]): continue
                    if 0 < len(last) <= 2 and not re.match(r'[•–▪]', x[1][-1][2].strip()):
                        sug('T8', '%s 第 %d 页：「%s…」折行后末行只剩「%s」——调列宽让它少折一行' % (name, i + 1, x[1][0][2].strip()[:14], last))
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

book = scan(BOOK, '全书', r'(第[零一二三四五六七八九]章|前言|总目录)')   # SD-96 页眉左侧为章名
scan(QREF, '单册', 'B737机型理论基础知识速查')   # 单册页眉左侧为册名（2026-09-30 用户定；SD-96 右侧为块名）

# B4 一页只有一两行（2026-09-30 用户：「尽量避免在一页中只有一两行的情况」）：正文（去页眉页脚）不超过 2 行的页
for nm, pdf in (('全书', BOOK), ('单册', QREF)):
    dd = pymupdf.open(pdf); Hh = dd[0].rect.height
    for i in range(1, len(dd) - 1):
        ln = [l for l in body_lines(dd[i]) if l[1] > 0.06 * Hh]
        if 0 < len(ln) <= 2:
            err('B4', '%s 第 %d 页只有 %d 行（「%s」）——调整上一页间距或内容，避免孤页' % (nm, i + 1, len(ln), ln[0][4].strip()[:20]))

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
