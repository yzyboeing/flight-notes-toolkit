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
import sys, os, re, glob, subprocess, collections, json
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
REPO = os.path.abspath(os.path.expanduser(arg('--repo', '~/flight-repos/gh-private')))
T = os.path.dirname(os.path.abspath(__file__))
BOOK = arg('--book') or os.path.join(REPO, 'build', 'B737机型理论知识笔记.pdf')
QREF = arg('--qref') or os.path.join(REPO, 'build', 'B737机型理论知识速查.pdf')
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
    for rule, tool, extra in (('S1', 'check_src.py', ['--quiet']), ('S2', 'check_blocks.py', [])) + ((('S3', 'check_quickref.py', []),) if os.path.exists(os.path.join(REPO, '速查', '速查源.md')) else ()):   # SD-146 速查版独立成册：源在 速查/速查源.md
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
_QMD = os.path.join(REPO, 'notes_src', '0 基础知识速查区', '0 基础知识速查区.md')   # SD-139 已删第零章：无此文件时 S5 不查
try:
    if not os.path.exists(_QMD): raise StopIteration
    _q = open(_QMD, encoding='utf-8').read()
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
except StopIteration:
    pass
except Exception as _e:
    sug('S5', '速查 / 正文标注对照未能完成：%s' % _e)

# ---------- F 成品文件 ----------
QON = os.path.exists(QREF)   # SD-139：速查单册停出，有单册文件时才检查
for p in (BOOK, BOOK[:-4] + '.docx') + ((QREF, QREF[:-4] + '.docx') if QON else ()):
    if not os.path.exists(p): err('F1', '缺成品：' + os.path.relpath(p, REPO))
for p in glob.glob(os.path.join(REPO, 'build', '*竖版*')):
    err('F2', '不应再有竖版成品（SD-71）：' + os.path.relpath(p, REPO))
src_m = max((os.path.getmtime(f) for f in glob.glob(os.path.join(REPO, 'notes_src', '*', '*.md'))), default=0)
for p in (BOOK,) + ((QREF,) if QON else ()):
    if os.path.exists(p) and os.path.getmtime(p) < src_m:
        err('F3', '%s 比 notes_src 旧，先重新 sync 再检查' % os.path.basename(p))
if not os.path.exists(BOOK):
    print('成品不全，先跑 sync.sh --full'); ERR and [print(' ', r, m) for r, m in ERR]; sys.exit(1)

# ---------- 逐页分析（两本都查） ----------
DESC_H = re.compile(r'^(说明|具体说明|条件|触发条件|限制|限值|条件 / 限值|定义|处置|处置流程|措施|要求|具体要求|内容|具体内容|描述|工作逻辑|控制逻辑|功能|作用|现象|结果|数值|标准|备注|原因|原理|工作原理|要点|注意事项|适用范围|逻辑|方法|做法|操作|动作|含义|影响|后果)$')
def _explicit(cls_re):
    out = set()
    try:
        bk = open(os.path.join(REPO, 'build', 'book.md'), encoding='utf-8').read()
        for m in re.finditer(r'<th class="([^"]*)">(.*?)</th>', bk):
            if re.search(cls_re, m.group(1)): out.add(re.sub(r'<[^>]+>', '', m.group(2)).strip())
    except Exception: pass
    return out
EXPLICIT_BULLET = _explicit(r'col-bullet')
EXPLICIT_NOBULLET = _explicit(r'col-(center|plain|left)')

def body_lines(pg):
    """本页正文的视觉行：[(x0, y0, x1, y1, text)]，去掉页眉页脚。"""
    out = []
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans']).replace('●', '•')   # SD-107：分条圆点排成缩小的「●」
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

PAD = 13   # 单元格左右内边距合计（pt）：左 170 + 右 90 DXA ≈ 13pt（2026-10-03 悬挂圆点后）
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
# R6（SD-152）：源文件标 one-page 的表（允许 7.5pt）——取表头文字作识别签名
ONEPAGE_SIGS = []
try:
    _bm = open(os.path.join(REPO, 'build', 'book.md'), encoding='utf-8').read()
    for _m in re.finditer(r'<table class="[^"]*\bone-page\b[^"]*">(?:(?!</table>).)*?<tr class="hdr">(.*?)</tr>', _bm, re.S):
        ONEPAGE_SIGS.append(nosp(re.sub(r'<[^>]+>', '', _m.group(1)))[:12])
    FIXW_SIGS = [nosp(re.sub(r'<[^>]+>', '', _m.group(1)))[:12] for _m in re.finditer(r'<tr class="hdr">((?:(?!</tr>).)*\bw-\d+(?:(?!</tr>).)*)</tr>', _bm, re.S)]
except Exception:
    FIXW_SIGS = []
NWL = set()   # fit_fix 实测证明「再收就多行」的列（keep_force 的 NWL: 记录）：已收到「不增行」的极限，T13 不报
try:
    for _kf in ('keep_force_B737机型理论知识笔记.txt', 'keep_force_B737机型理论知识速查.txt'):   # 全书与速查各有一份记录
        if not os.path.exists(os.path.join(REPO, 'build', _kf)): continue
        for _l in open(os.path.join(REPO, 'build', _kf), encoding='utf-8'):
            if _l.startswith('NWL:'):
                _k, _sg = _l.strip()[4:].split(':', 1); NWL.add((_sg[:20], int(_k)))
except Exception:
    pass

def scan(pdf, name, header):
    d = pymupdf.open(pdf)
    W, H = d[0].rect.width, d[0].rect.height
    # L1 只出横版
    port = [i + 1 for i, p in enumerate(d) if p.rect.width < p.rect.height]
    if port: err('L1', '%s 有竖版页：%s' % (name, port[:10]))
    # L2 字体统一（SD-107 思源宋体）：逐段核对，只允许 SourceHanSerifCN-Medium（正文）/ -Bold（粗体）
    bad = collections.Counter(); badx = {}
    for i, p in enumerate(d):
        for b in p.get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                for sp in l['spans']:
                    base = sp['font'].split('+')[-1]
                    if base.startswith('SourceHanSerifCN-'): continue
                    t = ''.join(ch['c'] for ch in sp['chars'])
                    bad[base] += 1; badx.setdefault(base, (i + 1, t[:12]))
    for f, n in bad.items(): err('L2', '%s 出现非思源宋体字体 %s（%d 段，如第 %d 页「%s」）——字体须统一思源宋体（SD-107）；多半是 LibreOffice profile 里没有字体文件' % (name, f, n, badx[f][0], badx[f][1]))
    # L3 书签栏（SD-77）
    if not d.get_toc(): err('L3', name + ' 没有书签')
    ut = [x[1] for x in d.get_toc() if '（单位：' in x[1]]
    if ut: err('L3', '%s 书签目录里有单位括注 %d 条（如「%s」）——目录不写「（单位：…）」（2026-10-01 用户）' % (name, len(ut), ut[0][:30]))
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
                     if sp['text'].strip() and sp['text'].strip() != '●' and sp['size'] < 7.9]   # 分条圆点「●」按半号排，不算小字
            if small and min(small) >= 6.9 and any(sig in nosp(p.get_text()) for sig in ONEPAGE_SIGS): small = []   # R6：源文件标 one-page 的超大表允许 7.5pt
            if small: err('T5', '%s 第 %d 页：表内有 %.1fpt 的字，低于 8pt 底线' % (name, i + 1, min(small)))
            if tw > CW + 4: err('T1', '%s 第 %d 页：表格宽 %.0fpt 超过版心 %.0fpt' % (name, i + 1, tw, CW))
            rows = [[c for c in r.cells] for r in t.rows]
            if not rows: continue
            ncol = max(len(r) for r in rows)
            info = [[(c, cell_lines(lines, c)) if c else None for c in r] for r in rows]
            # 表后淡蓝底注、公司差异、警告条紧贴表格时会被识别成表格的末几行：去掉，免得序号表判不出（2026-10-05）
            _nt = lambda r: r and r[0] and re.match(r'\s*(注[：:]|公司差异|警告[：:])', ''.join(l[2] for l in r[0][1]))
            while len(info) > 1 and _nt(info[-1]): info.pop(); rows.pop()
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
            # T13 改由 layout_measure 实测统一检查（见 scan 之后，跨页表按各段最小余量），这里不再单页判断
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
            first = [''.join(l[2] for l in r[0][1]).strip() for r in info[1:] if r and r[0] and r[0][1] and not (len(r) > 1 and all(x is None for x in r[1:]))]   # 跳过整行合并的说明行（SD-129 表后说明进表格末行）
            _sr = lambda x: re.match(r'^([①-⑳]|\d{1,2}[.、]?)$', x)
            serial = len(first) >= 2 and all(_sr(x) for x in (first if _sr(first[-1]) else first[:-1])) and len(first) - (0 if _sr(first[-1]) else 1) >= 2   # 末行可以是表后说明行（SD-129）
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
                # SD-119 修订（用户 2026-10-03「除了父子关系之外，去掉所有单独的句子之前的小圆点」）：单句格不加点是对的，
                # 只有没加点、格内却有多句（中间有「；」或「。」）的才算不统一
                multi = lambda x: len([p for p in re.split(r'[；]', ''.join(l[2] for l in x[1]).strip()) if len(p.strip()) >= 4]) >= 2
                nbs = [x for x in cells if x not in bul and sentence(x)]   # 2026-10-03 用户退回：一列要加点就整列都加（单句也加）
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
                # 表前有通栏说明行时，info[1] 是表头：最上面一格是 ≤ 8 字的短词（「步骤」「自动动作」），按表头处理，不算居中数据格
                hdrlike = lambda x: x is cells[0] and len(re.sub(r'\s', '', ''.join(l[2] for l in x[1]))) <= 8
                # 列宽收窄到刚好放下（2026-10-05）后，填满格子的文字左右留白都很小，看不出是居中还是靠左：两边都不算
                full = lambda x: all((x[0][2] - x[0][0]) - (l[1] - l[0]) <= PAD + 24 for l in x[1])
                cen = [x for x in nb if centered(x) and not hdrlike(x) and not full(x)]; lef = [x for x in nb if lefty(x) and not centered(x) and not full(x)]
                if cen and lef:
                    sug('T7', '%s 第 %d 页：第 %d 列 %d 格居中、%d 格左对齐（如「%s」）——统一对齐' % (
                        name, i + 1, k + 1, len(cen), len(lef), lef[0][1][0][2].strip()[:18]))
            # T13 首列靠左（SD-104，2026-10-01 用户：「一般第一列都是居中」）：首列是标签 / 短句却左对齐。首列是长句 / 问句 / 条件句（≥ 约 30 字或带句号）的允许左对齐，不报
            try:
                fc = [r[0] for r in info[1:] if r and r[0] and r[0][1] and (len(r) < 2 or r[1] is not None)]
                fc = [x for x in fc if not PH.match(''.join(l[2] for l in x[1]))]
                def fct(x): return ''.join(l[2] for l in x[1]).strip()
                def fleft(x):
                    cx = (x[0][0] + x[0][2]) / 2
                    return any(abs((l[0] + l[1]) / 2 - cx) > 4 and l[0] - x[0][0] < PAD + 2 for l in x[1]) and any((x[0][2] - x[0][0]) - (l[1] - l[0]) > PAD + 8 for l in x[1])
                lf = [x for x in fc if fleft(x)]
                longs = [x for x in fc if vis(re.sub(r'[（(][^）)]*[）)]', '', fct(x))) > 60 or '。' in fct(x) or fct(x).startswith('•')]
                if len(fc) >= 2 and len(lf) >= max(2, len(fc) // 2) and len(longs) < 0.4 * len(fc):
                    sug('T13', '%s 第 %d 页：首列「%s」等 %d 格左对齐——首列一般居中（标签 / 短句），表头标 col-center；首列是长句 / 问句的可保留左对齐' % (
                        name, i + 1, fct(lf[0])[:16], len(lf)))
            except Exception: pass
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
            # T8 末行孤字：改在 scan 之后统一用 layout_measure.measure() 实测（与 fit_fix 同一双眼睛，2026-10-05），这里不再自己另测
            # SD-102：序号表（首列全是 ①② / 1、2）——T4 不报（序号表不分条）；T12 报自动加点；非序号表的说明类句子列没加点报 T11
            try:
                ctext = lambda x: ''.join(l[2] for l in x[1]).strip() if x else ''
                # 序号列不一定在首列（2026-10-06 用户，2.3 B-1「类别｜序号｜条件」）：首列或表头为 序号 / # / 编号 / 步骤 / 条款 的列，与生成器 serialColOf 同口径
                _hd0 = [ctext(x) for x in info[0]] if info else []
                serial, sc = False, -1
                for _sc in range(min(3, len(_hd0))):
                    if _sc > 0 and not re.fullmatch(r'(序号|#|编号|步骤|条款)', _hd0[_sc].replace(' ', '')): continue
                    firsts = [ctext(r[_sc]) for r in info[1:] if r and _sc < len(r) and r[_sc]]
                    firsts = [f[0] if re.match(r'[\u2460-\u2473]\s*\S', f) else f for f in firsts]   # 「① 加标签文字」首列按序号表（2026-10-06 用户）
                    serial = len(firsts) >= 2 and all(re.fullmatch(r'([\u2460-\u2473]|\d{1,2}[.、]?|[（(]\s*(\d{1,2}|[a-zA-Z])\s*[)）]|[a-zA-Z][.、)）])', f) for f in firsts)
                    def _sv(t):
                        t = re.sub(r'[\s（()）.、]', '', t)
                        return ord(t) - 0x245F if re.fullmatch(r'[\u2460-\u2473]', t) else int(t) if t.isdigit() else ord(t.lower()) - 96 if re.fullmatch(r'[a-zA-Z]', t) else -1
                    if not serial and len(firsts) >= 4:   # 末尾 1～2 行非编号补充行（如「特殊情况」），与生成器同口径
                        for tn in (1, 2):
                            hd = firsts[:-tn]
                            if len(hd) >= 3 and all(re.fullmatch(r'([\u2460-\u2473]|\d{1,2}[.、]?|[（(]\s*(\d{1,2}|[a-zA-Z])\s*[)）]|[a-zA-Z][.、)）])', f) for f in hd) and [_sv(f) for f in hd] == list(range(1, len(hd) + 1)):
                                serial = True; firsts = hd; break
                    else:
                        serial = serial and [_sv(f) for f in firsts] == list(range(1, len(firsts) + 1))   # 从 1 开始的连续编号才算（襟翼位置 10 / 15 / 25 不算）
                    if serial: sc = _sc; break
                hdrs = [ctext(x) for x in info[0]] if info else []
                for k, h in enumerate(hdrs):
                    if k == 0 or not h: continue
                    ser_k = serial and k > sc   # 序号列右边的列才按序号表查
                    cs = [ctext(r[k]) for r in info[1:] if k < len(r) and r[k]]
                    cs = [c for c in cs if c and not re.fullmatch(r'[—\-–/无\s]+', c)]
                    if len(cs) < 2: continue
                    bul = sum(1 for c in cs if c.startswith('•') and '–' not in c)   # 「• 引语：」+「– 子项」是 F2 层级写法，不算自动加点
                    single_b = sum(1 for c in cs if c.count('•') == 1 and '–' not in c and len(c) < 80)
                    multi_b = any(c.count('•') >= 2 for c in cs)   # 2026-10-03 用户：一列要加点就整列都加——序号表这一列有多条加点时，单句加点是对的
                    if ser_k and any('•' in c or '●' in c for c in cs):   # SD-140（2026-10-05 用户）：「有了序号就不要加小圆点」，无例外，报错
                        err('T12', '%s 第 %d 页：序号表的「%s」列有小圆点——序号表其余列不加「•」（SD-140）' % (name, i + 1, h))
                    if ser_k and not re.fullmatch(r'(时机|宣布时机|总则|类别)', h.replace(' ', '')):   # SD-140「后面的一列都是靠左」（类别等居中表头除外，2026-10-06 用户）：居中的格（左右留白相等且明显大于内边距）报错（M18-L015，2026-10-06）
                        ctr = 0
                        pad = (collections.Counter(round(l[0] - x[0][0]) for r in info[1:] for x in r if x for l in x[1] if l[2].strip()).most_common(1) or [(6, 0)])[0][0]   # 本表实际左内边距：取众数（悬挂缩进的「–」子项会更靠左，不能取最小）
                        for r in info[1:]:
                            x = r[k] if k < len(r) else None
                            if not x or (k + 1 < len(r) and r[k + 1] is None): continue
                            c, ls = x
                            for l in ls:
                                if re.fullmatch(r'[—\-–/无\s]+', l[2].strip()): continue
                                lg, rg = l[0] - c[0], c[2] - l[1]
                                if lg > pad + 2.5 and abs(lg - rg) < 3: ctr += 1
                        if ctr: err('T12', '%s 第 %d 页：序号表的「%s」列有 %d 行居中——序号表其余列一律靠左（SD-140）' % (name, i + 1, h, ctr))
                    if any(re.match(r'([\u2460-\u2473]|\d+[.、])', c) for c in cs): continue   # 格内本身是 ①② / 1. 编号条目的列不报
                    sent = ([c for c in cs if vis(c) >= 12 or re.search(r'[，。；、]', c)] if re.search(r'说明$', h)
                            else [c for c in cs if vis(c) >= 36 or re.search(r'[，。；]', c)])   # 与生成器同口径，门槛略放宽避免临界误报
                    # SD-119 修订：单句不加点；只有格内有多句（「；」「。」分开 ≥ 2 句）却整列没加点时才提示
                    multi = [c for c in cs if len([p for p in re.split(r'[；]', c) if len(p.strip()) >= 4]) >= 2]
                    if (not serial and bul == 0 and len(sent) >= 0.6 * len(cs) and h not in EXPLICIT_NOBULLET
                            and (DESC_H.match(h) or re.search(r'(条件|要求|说明|逻辑|措施|处置|要点|内容)$', h))):
                        sug('T11', '%s 第 %d 页：说明类句子列「%s」没有加「•」——整列左对齐加点（SD-102）' % (name, i + 1, h))
            except Exception: serial = False
            # T4 并列长句未分条（SD-78）：一格内 ≥2 个「；」、各分句 ≥ 20 字宽、却没有 •
            for r in ([] if serial else info[1:]):
                for x in r:
                    if not x: continue
                    txt = ''.join(l[2] for l in x[1]).strip()
                    if re.search(r'[•–▪①-⑳]', txt) or re.match(r'(注|解释|出处|公司差异)[：:]', txt): continue
                    parts = [s for s in re.split(r'；', txt) if s.strip()]
                    if len(parts) >= 3 and all(vis(s) >= 20 for s in parts[:-1]):
                        sug('T4', '%s 第 %d 页：「%s…」含 %d 个并列长分句，考虑按语义分条加「•」' % (name, i + 1, txt[:20], len(parts)))
    return d

book = scan(BOOK, '全书', r'(第[零一二三四五六七八九]章|前言|总目录|目录|按主题查)')   # SD-96 页眉左侧为章名
if QON: scan(QREF, '单册', 'B737机型理论知识速查')   # SD-146 册名   # SD-139 单册停出：只有新近生成的单册才检查   # 单册页眉左侧为册名（2026-09-30 用户定；SD-96 右侧为块名）
# T8 末行孤字（SD-84 / SD-85，错误级）：用 layout_measure.measure() 实测——与 fit_fix 的自动修复看同一批格子，
# 检查器不再维护第二套测量（2026-10-05，APU 火警「1s」教训：两套眼睛必然漏）
try:
    from layout_measure import measure as _lm_measure
    for _nm, _pdf in (('全书', BOOK),) + ((('单册', QREF),) if QON else ()):
        _res = _lm_measure(_pdf)
        # T13 列宽多余（SD-151，错误级）：与 fit_fix 实测收窄同口径（严格余量：本格行数不增）；跨页的表按各段最小余量（页底段 + 下页页顶续段，列数、表宽相同算同一张）
        _segs = collections.OrderedDict()
        for _x in _res:
            if _x.get('nw') and _x.get('col', -1) >= 0 and 'tb' in _x:
                _segs.setdefault((_x['page'], _x['tb']), []).append(_x)
        _keys = list(_segs); _grp = {}; _gid = 0
        for _q, _k in enumerate(_keys):
            _e = _segs[_k][0]
            _cand = [kk for kk in _keys[:_q] if kk[0] == _k[0] - 1 and _segs[kk][0]['nc'] == _e['nc'] and abs(_segs[kk][0]['tw'] - _e['tw']) < 3] if _e['tb'] < 70 else []
            _prev = _cand[-1] if _cand else None   # 上一页列数、表宽相同的那张（块索引可按块分页，不一定是上一页最后一张、也不一定排到页底）
            if _prev:
                _grp[_k] = _grp[_prev]
            else:
                _gid += 1; _grp[_k] = _gid
        _byg = collections.defaultdict(list)
        for _k in _keys: _byg[_grp[_k]].append(_k)
        for _g, _ks in _byg.items():
            for _c in range(_segs[_ks[0]][0]['nc']):
                _vals = [next((e for e in _segs[_k] if e['col'] == _c), None) for _k in _ks]
                if any(v is None for v in _vals): continue
                _sl = min(v.get('extra_strict', v['extra_pt']) for v in _vals)
                if any(nosp(_vals[0].get('table', '')).startswith(re.sub(r'[^\w]', '', sg)[:10]) for sg in FIXW_SIGS): continue   # 表头标 w-NN 的固定列宽表不按自动列宽判
                if any(nosp(_vals[0].get('table', '')).startswith(sg) and kk == _c for sg, kk in NWL): continue   # 已证明收到极限（再收就多行）
                if _sl >= 10 and not any(v.get('wrapped') or v.get('anywrap') for v in _vals):   # 只查没折行的列（用户原话「空白太多」）
                    err('T13', '%s 第 %d 页：第 %d 列（「%s」）没有折行，右侧仍空约 %.0fpt——按最长一行收窄（以文字成行为标准）' % (_nm, _ks[0][0], _c + 1, _vals[0].get('head', ''), _sl))
        if _nm != '全书': continue
    for _x in _lm_measure(BOOK):
        if _x.get('orphan'):
            err('T8', '全书 第 %d 页：「%s…」末行只剩一两个字（上一行止于「…%s」）——加宽该列或收紧字距（fit_fix 会自动处理，仍在就看 keep_force 的 W/WB/C 记录）' % (_x['page'], _x['cell'][:16], _x['prev'][-10:]))
except Exception as _e:
    err('T8', '末行孤字实测未能完成：%s' % _e)


# B4 一页只有一两行（2026-09-30 用户：「尽量避免在一页中只有一两行的情况」）：正文（去页眉页脚）不超过 2 行的页
for nm, pdf in (('全书', BOOK),) + ((('单册', QREF),) if QON else ()):
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
if '按主题查' not in toc: err('P2', '目录里缺「按主题查（速查入口）」（SD-139 / SD-149）')
# SD-149（2026-10-06 用户）：五种查法放在第五章之后——第五章最后一页之后才出现「按主题查」页
_tp = [i for i in range(6, len(book)) if nosp(book[i].get_text()).startswith('按主题查') or '按主题查按系统' in nosp(book[i].get_text())[:40]]
_c5 = max((i for i in range(len(book)) if '第五章' in book[i].get_text()[:60]), default=-1)
if not _tp: err('P2', '找不到「按主题查」页（SD-149）')
elif _tp[0] < _c5: err('P2', '「按主题查」在第 %d 页，应放在第五章之后（SD-149）' % (_tp[0] + 1))

# R5（SD-154，2026-10-06 用户）：同一条目下同表头的几张表、中间只隔着注——应合成一张（标 one-page）。
# 例外：压到最小仍放不下一页、用户定为照常跨页的几处（范围 A）
R5_EXEMPT = {('4.8', '2.'), ('4.9', '2.'), ('5.6', '3.'), ('2.7', 'A-5'), ('3.6', '块索引')}
import glob as _g
for _f in sorted(_g.glob(os.path.join(REPO, 'notes_src/[1-5]*/*.md'))):
    _sec = os.path.basename(_f).split(' ')[0]; _s = open(_f, encoding='utf-8').read()
    for _m in re.finditer(r'(?ms)^#{3,4} ([^\n]*)\n(.*?)(?=^#{2,4} |\Z)', _s):
        _id = _m.group(1).split('　')[0].split(' ')[0]
        if (_sec, _id) in R5_EXEMPT: continue
        _tb = list(re.finditer(r'(?s)<table[^>]*>.*?</table>', _m.group(2)))
        _hd = lambda t: re.sub(r'<[^>]+>|\s', '', (re.search(r'<tr class="hdr">(.*?)</tr>', t, re.S) or [0, ''])[1])
        for _a, _b in zip(_tb, _tb[1:]):
            _gap = [p.strip() for p in re.split(r'\n\s*\n', _m.group(2)[_a.end():_b.start()]) if p.strip()]
            if _hd(_a.group(0)) and _hd(_a.group(0)) == _hd(_b.group(0)) and all(p.startswith('注') for p in _gap):
                err('R5', '%s「%s」有两张同表头的表、中间只隔注——合成一张、注移到表下、标 one-page（SD-154）' % (_sec, _m.group(1).strip()[:24]))

# V 成品通用校验（verify.py：空白页、标签泄漏、异常项目符号、front matter）
for pdf in (BOOK,) + ((QREF,) if QON else ()):
    r = subprocess.run([sys.executable, os.path.join(T, 'verify.py'), pdf], capture_output=True, text=True)
    if r.returncode: err('V1', '%s verify.py 未通过：%s' % (os.path.basename(pdf), ' / '.join(l for l in r.stdout.splitlines() if '[]' not in l)))

# B 断表 / 孤行 / 标题孤立（SD-71 / SD-77 / SD-79）
for pdf in (BOOK,) + ((QREF,) if QON else ()):
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
# 建议裁定（2026-10-05）：裁定过的建议（改了，或「不改＋理由」）记进 gh-private/排版建议裁定.json
# { "<签名>": {"rule": "Z1", "摘要": "…", "结论": "不改", "理由": "…", "日期": "2026-10-05"} }
# 签名 = 规则号 + 去掉页码后的内容文字（页码随分页漂移，不进签名）。已裁定的建议只计数，不再逐条列出；
# 报告对每条新建议给出签名，裁定后把整块 JSON 粘进裁定文件即可。内容变了签名就变，会重新出现。
import hashlib
def _sig(rule, msg):
    t = re.sub(r'第\s*\[?[\d,\s\[\]…]+\]?\s*页', '', msg)
    t = re.sub(r'[\s，。；：]', '', t)
    return rule + '-' + hashlib.md5(t.encode('utf-8')).hexdigest()[:8]
ADJ_FILE = os.path.join(REPO, '排版建议裁定.json')
try:
    ADJ = json.load(open(ADJ_FILE, encoding='utf-8')) if os.path.exists(ADJ_FILE) else {}
except Exception:
    ADJ = {}
news, settled = [], []
for rule, msg in SUG:
    (settled if _sig(rule, msg) in ADJ else news).append((rule, msg))
out = ['# 排版与规则检查报告', '',
       '全书 %d 页。错误 %d 条（必须改）；建议：新 %d 条（逐条裁定），已裁定 %d 条（结论在 排版建议裁定.json，不再列出）。'
       % (len(book), len(ERR), len(news), len(settled)), '']
out.append('## 错误')
if not ERR: out.append('（无）')
for rule, msg in sorted(ERR, key=lambda x: x[0]): out.append('- [%s] %s' % (rule, msg))
out.append('')
out.append('## 新建议（裁定后把下面的 JSON 行并进 排版建议裁定.json）')
if not news: out.append('（无）')
for rule, msg in sorted(news, key=lambda x: x[0]):
    out.append('- [%s] %s' % (rule, msg))
    out.append('  `"%s": {"rule": "%s", "摘要": "%s", "结论": "", "理由": "", "日期": ""}`'
               % (_sig(rule, msg), rule, re.sub(r'[`"\\]', '', msg)[:40]))
out.append('')
if settled:
    out.append('## 已裁定建议（%d 条，按规则号计数）' % len(settled))
    cnt = {}
    for rule, _ in settled: cnt[rule] = cnt.get(rule, 0) + 1
    out.append('、'.join('%s×%d' % (r, n) for r, n in sorted(cnt.items())))
    out.append('')
txt = '\n'.join(out)
print(txt)
if arg('--out'): open(arg('--out'), 'w', encoding='utf-8').write(txt + '\n')
sys.exit(1 if ERR else 0)
