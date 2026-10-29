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
def _nocmp_hdrs():   # SD-188：源文件首格标 no-cmp 的表（跨列格不是共有内容 / 用户指定不按对比原则），T16 跳过；按表头文字识别
    import glob as _gl
    out = set()
    for f in _gl.glob(os.path.join(REPO, 'notes_src', '*', '*.md')) + [os.path.join(REPO, '速查', '速查源.md')]:
        try: t = open(f, encoding='utf-8').read()
        except Exception: continue
        for m in re.finditer(r'<tr class="hdr">(.*?)</tr>', t, re.S):
            ths = re.findall(r'<th([^>]*)>(.*?)</th>', m.group(1), re.S)
            if len(ths) >= 3 and 'col-center' in ths[-1][0]:   # 图表并排里的三列表，PDF 识别时首列常被吞（1.2「实际升降方向」）：后两列表头也登记
                out.add(''.join(re.sub(r'<[^>]+>|\s', '', x[1]) for x in ths[1:]))
            if len(ths) == 2:   # 第 95 条（2026-10-20）：两列表第二列全是取值 → 生成器自动居中不加点，T15 同口径豁免
                _tb = t[m.end():t.find('</table>', m.end())]
                _vc = [re.sub(r'<[^>]+>|&[a-z]+;|\s', '', c) for c in re.findall(r'<tr(?![^>]*class="(?:hdr|note|premise|warn)")[^>]*>\s*<td[^>]*>.*?</td>\s*<td[^>]*>(.*?)</td>', _tb, re.S)]
                _vc = [re.sub(r'^(约|至少|不少于|不超过|最多|最少|小于|大于)', '', re.sub(r'（[^）]{0,14}）|\([^)]{0,14}\)', '', c)) for c in _vc if c]
                if len(_vc) >= 2 and all(re.search(r'\d', c) and re.fullmatch(r'[\d.,:：;；~～+\-−–≤≥<>%/°×±℃a-zA-Z·→]*', c) for c in _vc):
                    out.add(''.join(re.sub(r'<[^>]+>|\s', '', x[1]) for x in ths)); continue
            if ths and ('no-cmp' in ths[0][0] or 'no-share' in ths[0][0]):   # no-cmp：两型列按表头标注（可居中，SD-165 / 用户点名）；no-share：跨列格不是共有内容——T16 都跳过（共有格居中由生成器保证，SD-192 修订）
                out.add(''.join(re.sub(r'<[^>]+>|\s', '', x[1]) for x in ths))
            elif len(ths) == 2 and re.search(r'(^|\s)col-center(\s|$)', re.sub(r'.*class="([^"]*)".*', r'\1', ths[1][0]) if 'class=' in ths[1][0] else ''):   # 1019R6（记录第 47 条）：两列表取值 / 等级列标 col-center——居中不加点，T15 跳过
                out.add(''.join(re.sub(r'<[^>]+>|\s', '', x[1]) for x in ths))
    return out
NOCMP_HDRS = _nocmp_hdrs()
T = os.path.dirname(os.path.abspath(__file__))
BOOK = arg('--book') or os.path.join(REPO, 'build', 'B737机型理论速查笔记.pdf')
QREF = arg('--qref') or os.path.join(REPO, 'build', 'B737机型理论速查.pdf')
SECTION = '--section' in sys.argv   # 2026-10-07 单节预览（preview.py）：只查表格与版面，跳过成品文件、前言 / 目录、速查等整书检查
MDP = arg('--md') or os.path.join(REPO, 'build', 'book.md')   # 生成这本 PDF 的源（取源文件 <br>、one-page、w-NN 标记）
if SECTION and '--qref' not in sys.argv: QREF = '/nonexistent'
def git_cfg(k):   # 2026-10-14：先读库内 成品设置.gitconfig，没有的键再读本机 git config（同 sync.sh 的 cfg）
    r = subprocess.run(['git', '-C', REPO, 'config', '-f', os.path.join(REPO, '成品设置.gitconfig'), '--get', k], capture_output=True, text=True)
    if r.returncode == 0: return r.stdout.strip()
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
    for rule, tool, extra in (('S1', 'check_src.py', ['--quiet']), ('S2', 'check_blocks.py', [])) + ((('S3', 'check_quickref.py', []),) if os.path.exists(os.path.join(REPO, '速查', '速查源.md')) and os.environ.get('BUILD_QR') else ()):   # SD-146 速查版独立成册：源在 速查/速查源.md；SD-242 起速查源只读溯源，S3 只在 BUILD_QR=1 对照时跑
        r = subprocess.run([sys.executable, os.path.join(T, tool), '--src', 'notes_src'] + extra, cwd=REPO, capture_output=True, text=True)
        if r.returncode:
            tail = [l for l in (r.stdout + r.stderr).strip().splitlines() if l.strip()][-6:]
            err(rule, '%s 未通过：\n      ' % tool + '\n      '.join(tail))
    if not os.path.isdir(os.path.join(REPO, 'notes_src', '3 运行规范')):
        err('S4', '第三章文件夹应为 notes_src/3 运行规范/（2026-10-25 用户改名，取代 SD-73）')

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
    for _f in glob.glob(os.path.join(REPO, 'notes_src', '[1-6]*', '*.md')):
        _m = re.match(r'(\d\.\d+)', os.path.basename(_f))
        if _m: _files[_m.group(1)] = _f
    for _it in re.split(r'(?m)^(?=### \d+\. )', _q):
        _m = re.match(r'### (\d+)\. (.+)', _it)
        if not _m: continue
        for _sec, _addr in re.findall(r'\|(\d\.\d+) ([A-Z]-\d+|第 ?\d+ ?条)\]\]', _it):   # SD-201：全书「x.y 第 N 条」
            if _sec not in _files: continue
            _s = open(_files[_sec], encoding='utf-8').read()
            _n = re.match(r'第 ?(\d+) ?条', _addr)
            _a = re.search(r'(?m)^### ' + (_n.group(1) + r'\. ' if _n else re.escape(_addr) + r'[\s\u3000]'), _s)
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
# 2026-10-14 提速：成品核验与断表检查（各自独立读 PDF、互不依赖）一开始就放到后台并行，下面原位置再取结果——报告顺序与内容不变
_BG = {}
for _pdf in (BOOK,) + ((QREF,) if QON else ()):
    if os.path.exists(_pdf):
        for _tool in ('verify.py', 'check_splits.py'):
            _BG[(_tool, _pdf)] = subprocess.Popen([sys.executable, os.path.join(T, _tool), _pdf], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
class _R:
    def __init__(s, tool, pdf):
        pr = _BG.pop((tool, pdf), None)
        if pr is None:
            r = subprocess.run([sys.executable, os.path.join(T, tool), pdf], capture_output=True, text=True); s.stdout, s.returncode = r.stdout, r.returncode
        else:
            s.stdout, _e = pr.communicate(); s.returncode = pr.returncode
for p in (() if SECTION else (BOOK, BOOK[:-4] + '.docx')) + ((QREF, QREF[:-4] + '.docx') if QON else ()):
    if not os.path.exists(p): err('F1', '缺成品：' + os.path.relpath(p, REPO))
for p in glob.glob(os.path.join(REPO, 'build', '*竖版*')):
    err('F2', '不应再有竖版成品（SD-71）：' + os.path.relpath(p, REPO))
src_m = max((os.path.getmtime(f) for f in glob.glob(os.path.join(REPO, 'notes_src', '*', '*.md'))), default=0)
for p in (() if SECTION else (BOOK,)) + ((QREF,) if QON else ()):
    if os.path.exists(p) and os.path.getmtime(p) < src_m:
        err('F3', '%s 比 notes_src 旧，先重新 sync 再检查' % os.path.basename(p))
# T10 分项落实（B7 / SD-180，2026-10-14 起自动，原为人工校对项）：读 Word 成品逐段判角色
#   错误：段首「– -」双短横、段首原样「- 」（生成器没吃掉源文件的列表短横，10-13 重新校对查出的两类回归）
#   建议：表格里的孤立分项（前一段不是引导语 / 分项 / 编号 / 并列条目）；引导语下只有一个分项（规则要合并成一行）
def _b7(docx, name):
    import zipfile, html as _h
    try: x = zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')
    except Exception: return
    def ptxt(pp):
        runs = re.findall(r'<w:r\b.*?</w:r>', pp, re.S)
        t = ''.join(_h.unescape(''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', r, re.S))) for r in runs).strip()
        navy = any('w:val="1F4E79"' in r and re.search(r'<w:t[^>]*>[^<]*[^\s<]', r) for r in runs)
        return t, navy
    dd, raw = [], []
    for pp in re.findall(r'<w:p\b.*?</w:p>', x, re.S):
        t, _ = ptxt(pp)
        if re.match(r'[–—]\s*[-–—]\s', t): dd.append(t)
        elif re.match(r'-\s', t): raw.append(t)
    if dd: err('T10', '%s 有 %d 段以「– -」双短横开头（如「%s」）——源文件的「- 」没被生成器吃掉' % (name, len(dd), dd[0][:30]))
    if raw: err('T10', '%s 有 %d 段以原样「- 」开头（如「%s」）——应排成「–」分项或「•」并列条目' % (name, len(raw), raw[0][:30]))
    for tc in re.findall(r'<w:tc>.*?</w:tc>', x, re.S):
        if '<w:tbl>' in tc: continue
        P = [q for q in (ptxt(pp) for pp in re.findall(r'<w:p\b.*?</w:p>', tc, re.S)) if q[0]]
        role = ['bul' if t.startswith('●') else 'dash' if re.match(r'–[ \u00a0]', t) else 'num' if re.match(r'[①-⑳]', t)
                else 'lead' if nv and re.search(r'[：:]$', t) else 'plain' for t, nv in P]
        for k, r in enumerate(role):
            if r == 'dash' and (k == 0 or role[k - 1] not in ('lead', 'dash', 'num', 'bul')):
                sug('T10', '%s：孤立分项「%s」——前一段「%s」不是引导语，分项应挂在以「：」结尾的引导语下（SD-180）' % (name, P[k][0][:24], P[k - 1][0][:24] if k else '（格首）'))
            if r == 'lead':
                kids = 0
                for j in range(k + 1, len(role)):
                    if role[j] in ('dash', 'num'): kids += 1
                    else: break
                if kids == 1: sug('T10', '%s：引导语「%s」下只有一个分项「%s」——按 B7 合并成一行' % (name, P[k][0][:24], P[k + 1][0][:24]))
for _d, _nm in ((BOOK[:-4] + '.docx', '全书'),) + (((QREF[:-4] + '.docx', '单册'),) if QON else ()):
    if os.path.exists(_d): _b7(_d, _nm)
if not os.path.exists(BOOK):
    print('成品不全，先跑 sync.sh --full'); ERR and [print(' ', r, m) for r, m in ERR]; sys.exit(1)

# ---------- 逐页分析（两本都查） ----------
DESC_H = re.compile(r'^(说明|具体说明|条件|触发条件|限制|限值|条件 / 限值|定义|处置|处置流程|措施|要求|具体要求|内容|具体内容|描述|工作逻辑|控制逻辑|功能|作用|现象|结果|数值|标准|备注|原因|原理|工作原理|要点|注意事项|适用范围|逻辑|方法|做法|操作|动作|含义|影响|后果)$')
def _explicit(cls_re):
    out = set()
    try:
        bk = open(MDP, encoding='utf-8').read()
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
_bk = MDP
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
SMALL_SIGS = []
try:
    _bm = open(MDP, encoding='utf-8').read()
    _qm = os.path.join(REPO, 'build', 'qr_single.md')   # 2026-10-09：速查单册的 one-page / w-NN 表也要认（此前只读全书）
    if os.path.exists(_qm): _bm += '\n' + open(_qm, encoding='utf-8').read()
    for _m in re.finditer(r'<table class="[^"]*\bone-page\b[^"]*">(?:(?!</table>).)*?<tr class="hdr">(.*?)</tr>', _bm, re.S):
        ONEPAGE_SIGS.append(nosp(re.sub(r'<[^>]+>', '', _m.group(1)))[:12])
    SMALL_SIGS = [nosp(re.sub(r'<[^>]+>', '', _m.group(1)))[:12] for _m in re.finditer(r'<table class="[^"]*\bsmall\b[^"]*">(?:(?!</table>).)*?<tr class="hdr">(.*?)</tr>', _bm, re.S)]   # SD-218：用户点名缩小的表（small）也不按 9pt 查
    # SD-208（2026-10-18）：签名改为「表头 + 第一行」前 30 字（原只取表头 12 字，表头相同的自动宽度表——「类别｜条件」「项目｜说明」——被一并跳过，A/T 脱开条件表的大片空白因此漏报）；
    #   区分锁宽（w-lock，空白 ≥ 30pt 才报）与起始比例（w-NN，fit_fix 多留 14pt，空白 ≥ 24pt 才报）
    FIXW_SIGS = []
    for _m in re.finditer(r'<table[^>]*>((?:(?!</table>).)*?)<tr class="hdr">((?:(?!</tr>).)*\bw-\d+(?:(?!</tr>).)*)</tr>((?:(?!</table>).)*?</tr>)', _bm, re.S):
        _sig = nosp(re.sub(r'<[^>]+>', '', _m.group(2) + _m.group(3)))[:30]
        FIXW_SIGS.append((_sig, 10 ** 6 if ('keep-wide' in _m.group(2) or 'col-eq' in _m.group(2)) else 30 if re.search(r'\bw-lock\b', _m.group(2)) else 24))   # keep-wide：用户点名要满宽的表（R3「紧急与遇险」），不报空白
    for _m in re.finditer(r'<table[^>]*>((?:(?!</table>).)*?)<tr class="hdr">((?:(?!</tr>).)*\b(?:col-eq|keep-wide)\b(?:(?!</tr>).)*)</tr>((?:(?!</table>).)*?</tr>)', _bm, re.S):
        FIXW_SIGS.append((nosp(re.sub(r'<[^>]+>', '', _m.group(2) + _m.group(3)))[:30], 10 ** 6))   # 等宽 / 点名满宽表不一定标 w-NN（速查两型表 col-eq）
    for _m in re.finditer(r'^%%FIG(?:SIDE|RIGHT)%%[^\n]*\n\s*<table[^>]*>((?:(?!</table>).)*?)<tr class="hdr">((?:(?!</tr>).)*)</tr>((?:(?!</table>).)*?</tr>)', _bm, re.S | re.M):
        FIXW_SIGS.append((nosp(re.sub(r'<[^>]+>', '', _m.group(1) + _m.group(2) + _m.group(3)))[:30], 10 ** 6))   # 第 109 条（前提行也算进签名，实测表文从前提行起）
        FIXW_SIGS.append((nosp(re.sub(r'<[^>]+>', '', _m.group(2) + _m.group(3)))[:30], 10 ** 6))   # 图表并排的表占满所在栏，不报列空白（宽度按栏分配）
        _h1 = re.findall(r'<th[^>]*>(.*?)</th>', _m.group(2), re.S)[1:]; _r1 = re.findall(r'<td[^>]*>(.*?)</td>', _m.group(3), re.S)[1:]
        FIXW_SIGS.append((nosp(re.sub(r'<[^>]+>', '', ''.join(_h1) + ''.join(_r1)))[:30], 10 ** 6))   # 实测有时漏掉首列（嵌套表识别），去首列的签名也登记
except Exception:
    FIXW_SIGS = []
BWB = set()
WZ = set()   # 2026-10-23（SD-232 ⑨）：fit_fix 实测调整来回跳、已冻结的列（keep_force 的 WZ: 记录）：T13 / T8 只提示
NWL = set()   # fit_fix 实测证明「再收就多行」的列（keep_force 的 NWL: 记录）：已收到「不增行」的极限，T13 不报
try:
    for _kf in ('keep_force_B737机型理论速查笔记.txt', 'keep_force_B737机型理论速查.txt') + ((os.path.join('preview', 'keep_force_预览.txt'),) if SECTION else ()):   # 全书与速查各有一份记录；单节预览（--section）另读预览自己的一份——全量时不读，免得旧预览记录豁免整书错误
        if not os.path.exists(os.path.join(REPO, 'build', _kf)): continue
        for _l in open(os.path.join(REPO, 'build', _kf), encoding='utf-8'):
            if _l.startswith('NWL:'):
                _k, _sg = _l.strip()[4:].split(':', 1); NWL.add((_sg[:20], int(_k)))
            if _l.startswith('WZ:'):
                _k, _sg = _l.strip()[3:].split(':', 1); WZ.add((_sg[:20], int(_k)))
            if _l.startswith('BWB:'):   # SD-192：借过宽、借出列因此出孤字已撤销的列——借不到，T8 按建议
                _k, _sg = _l.strip()[4:].split(':', 1); BWB.add(_sg[:20])
except Exception:
    pass

def _par(fn, n):
    """把 range(n) 切成连续页段：第一段在本进程跑，其余各段在分叉子进程跑；子进程传回 ERR / SUG 的新增条目，按段序接回（2026-10-14）"""
    J = int(os.environ.get('CHECK_JOBS') or min(os.cpu_count() or 1, 8))
    if J <= 1 or n < 60 or not hasattr(os, 'fork') or os.environ.get('CHECK_SERIAL'):
        fn(range(n), True); return
    import pickle
    cuts = [n * k // J for k in range(J + 1)]
    kids = []
    for k in range(1, J):
        r, w = os.pipe(); n0 = (len(ERR), len(SUG)); pid = os.fork()
        if pid == 0:
            os.close(r)
            try: fn(range(cuts[k], cuts[k + 1]), False); out = ('ok', ERR[n0[0]:], SUG[n0[1]:])
            except BaseException as e: out = ('fail', repr(e), None)
            with os.fdopen(w, 'wb') as f: pickle.dump(out, f)
            os._exit(0)
        os.close(w); kids.append((pid, r))
    fn(range(cuts[0], cuts[1]), True)
    for pid, r in kids:
        with os.fdopen(r, 'rb') as f: out = pickle.load(f)
        os.waitpid(pid, 0)
        if out[0] != 'ok': raise RuntimeError('逐页检查子进程失败：' + out[1])
        ERR.extend(out[1]); SUG.extend(out[2])

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
    mt = [x[1] for x in d.get_toc() if '【737' in x[1]]
    if mt: err('L3', '%s 书签目录里有机型标签 %d 条（如「%s」）——目录、书签不写机型（F11 / SD-195）' % (name, len(mt), mt[0][:30]))
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
    # 2026-10-14 提速：逐页表格检查各页互不相干（只往 ERR / SUG 里添），按连续页段分给分叉子进程，
    # 各段的新增条目按页序接回，结果与串行相同；子进程各自重开 PDF（共用文件描述符的读位置会互相干扰）。CHECK_SERIAL=1 回到串行
    def _tpass(dd, rng):
        for i in rng:
            p = dd[i]
            try: tabs = p.find_tables().tables
            except Exception: continue
            lines = body_lines(p)
            for t in tabs:
                if t.bbox[3] - t.bbox[1] < 8 or t.bbox[1] < 0.08 * H: continue
                tw = t.bbox[2] - t.bbox[0]
                # T5 表内字号底线 8pt（SD-80）
                small = [sp['size'] for b in p.get_text('dict', clip=t.bbox)['blocks'] for l in b.get('lines', []) for sp in l['spans']
                         if sp['text'].strip() and sp['text'].strip() != '●' and sp['size'] < 6.9]   # 分条圆点「●」按半号排，不算小字；2026-10-22 SD-223：底线 7.5pt（压缩后放进一页的表）
                if small and min(small) >= 6.9 and any(sig in nosp(p.get_text()) for sig in ONEPAGE_SIGS): small = []   # R6：源文件标 one-page 的超大表允许 7.5pt
                if small: err('T5', '%s 第 %d 页：表内有 %.1fpt 的字，低于 7pt 底线' % (name, i + 1, min(small)))
                # 2026-10-22（SD-223）：取消「表格一律 9pt」——一页放得下、下截只差一点的表允许收紧到 8pt 留在本页（fit_fix 实测定级），底线仍是 8pt
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
                # T17 列宽比例失调（SD-195，2026-10-16 自审 K-15：1.1 C-3「保障对象」列逐字折行、右列大片空白，T14 没报出）：
                # 某列有格折成 ≥ 3 行且每行平均 ≤ 6 字，同表另有列右侧空白 ≥ 30% 列宽 → 建议（用 autow 实测定宽）
                _nar = []
                for k in range(ncol):
                    _seg = [len(''.join(l[2] for l in x[1]).replace(' ', '')) / len(x[1]) for r in info for kk, x in enumerate(r)
                            if kk == k and x and len(x[1]) >= 3 and not (kk + 1 < len(r) and r[kk + 1] is None)]
                    if _seg and max(_seg) <= 6: _nar.append(k)
                _wide = [k for k in range(ncol) if k not in _nar and colw[k] and spare[k] >= 0.3 * colw[k]]
                if _nar and _wide:
                    sug('T17', '%s 第 %d 页：第 %s 列逐字折行（每行 ≤ 6 字），第 %s 列右侧空白 ≥ 30%%——列宽比例失调，用 autow 按实测定宽' % (
                        name, i + 1, '、'.join(str(k + 1) for k in _nar), '、'.join(str(k + 1) for k in _wide)))
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
                _hd188 = [''.join(l[2] for l in x[1]).strip() if x else '' for x in info[0]] if info else []
                _span188 = any(any(k >= 1 and r[k] and k + 1 < len(r) and r[k + 1] is None for k in range(len(r))) for r in info[1:] if r)
                _two188 = len(_hd188) == 2   # SD-188 ②：两列表第二列按规则靠左加点；① 共有格表独有文字格靠左、数值居中——T3 / T7 不按旧口径提示
                # T3 整列都是一行短内容却没居中（SD-75 / SD-78），第 2 列起、至少 2 格；序号表（首列 ①②③ / 1 2 3）按 SD-90 左对齐，不查
                first = [''.join(l[2] for l in r[0][1]).strip() for r in info[1:] if r and r[0] and r[0][1] and not (len(r) > 1 and all(x is None for x in r[1:]))]   # 跳过整行合并的说明行（SD-129 表后说明进表格末行）
                _sr = lambda x: re.match(r'^([①-⑳]|\d{1,2}[.、]?)$', x)
                serial = len(first) >= 2 and all(_sr(x) for x in (first if _sr(first[-1]) else first[:-1])) and len(first) - (0 if _sr(first[-1]) else 1) >= 2   # 末行可以是表后说明行（SD-129）
                serial = serial or (len(first) >= 1 and all(re.match(r'^[①-⑳]$', x) for x in first))   # SD-218：跨页拆开的序号表，这一截可能只有一两行（③④）
                for k in (range(1, ncol) if not serial and not _two188 and not _span188 else []):
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
                    # SD-180：引导语不加「•」——有「……：」+「–」/ ①② 分项的格是层级格，不按漏点计
                    lead = lambda x: any(re.search(r'[：:]\s*$', a[2]) and re.match(r'\s*[–①-⑳]', b[2]) for a, b in zip(x[1], x[1][1:]))
                    bul = [x for x in cells if x[1][0][2].lstrip().startswith('•')]
                    # 没加点的格只算「句子」：≥ 12 字或以句号结尾；短值（持续 / ≥ 800m）和 ①② 步骤格不算不统一
                    def sentence(x):
                        t0 = ''.join(l[2] for l in x[1]).strip()
                        return not re.match(r'[①-⑳]', t0) and '①' not in t0 and (len(re.sub(r'\s', '', t0)) >= 12 or t0.endswith('。'))
                    # SD-119 修订（用户 2026-10-03「除了父子关系之外，去掉所有单独的句子之前的小圆点」）：单句格不加点是对的，
                    # 只有没加点、格内却有多句（中间有「；」或「。」）的才算不统一
                    multi = lambda x: len([p for p in re.split(r'[；]', ''.join(l[2] for l in x[1]).strip()) if len(p.strip()) >= 4]) >= 2
                    nbs = [x for x in cells if x not in bul and sentence(x) and not lead(x)]   # SD-180：引导语格不算漏点，也不算加点格   # 2026-10-03 用户退回：一列要加点就整列都加（单句也加）
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
                    if cen and lef and not _span188:
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
                            # SD-180（2026-10-13）：引导语不再加「•」——以「：」结尾、不以「•」「–」开头的行后面紧跟「•」，也是子项排成了同级
                            if re.search(r'[：:]$', a0) and not a0.startswith('–') and b0.startswith('•') and (a0.startswith('•') or not re.match(r'^[①-⑳\d]', a0)):
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
                        _hsr = bool(re.fullmatch(r'(序号|#|编号|步骤|条款)', _hd0[_sc].replace(' ', ''))) if _sc < len(_hd0) else False
                        serial = (len(firsts) >= 2 or (len(firsts) >= 1 and (_hsr or all(re.fullmatch(r'[\u2460-\u2473]', f) for f in firsts)))) and all(re.fullmatch(r'([\u2460-\u2473]|\d{1,2}[.、]?|[（(]\s*(\d{1,2}|[a-zA-Z])\s*[)）]|[a-zA-Z][.、)）])', f) for f in firsts)
                        def _sv(t):
                            t = re.sub(r'[\s（()）.、]', '', t)
                            return ord(t) - 0x245F if re.fullmatch(r'[\u2460-\u2473]', t) else int(t) if t.isdigit() else ord(t.lower()) - 96 if re.fullmatch(r'[a-zA-Z]', t) else -1
                        if not serial and len(firsts) >= 4:   # 末尾 1～2 行非编号补充行（如「特殊情况」），与生成器同口径
                            for tn in (1, 2):
                                hd = firsts[:-tn]
                                if len(hd) >= 3 and all(re.fullmatch(r'([\u2460-\u2473]|\d{1,2}[.、]?|[（(]\s*(\d{1,2}|[a-zA-Z])\s*[)）]|[a-zA-Z][.、)）])', f) for f in hd) and [_sv(f) for f in hd] == list(range(1, len(hd) + 1)):
                                    serial = True; firsts = hd; break
                        else:
                            _vs = [_sv(f) for f in firsts]
                            serial = serial and bool(_vs) and _vs == list(range(_vs[0], _vs[0] + len(_vs))) and (_vs[0] == 1 or _hsr or all(re.fullmatch(r'[\u2460-\u2473]', f) for f in firsts))   # 从 1 开始的连续编号才算（襟翼位置 10 / 15 / 25 不算）；SD-218：跨页续截可从 ③ 起
                        if serial and len([h for h in _hd0[_sc + 1:] if h.strip()]) >= 2 and not os.environ.get('DOC_SERIAL_OLD'): serial = False   # SD-207 方案 A：序号列右边两列以上内容的不按序号表（与生成器 serialColOf 同口径）
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
                        b124 = (len(hdrs) >= 3 and not re.search(r'737|仅|两型', h)
                                and all(not re.search(r'[；。]', c.rstrip('；。')) for c in cs))   # 第 124 条（SD-219）：三列以上表整列单项居中不加点，不论字数
                        if (not serial and bul == 0 and not b124 and len(sent) >= 0.6 * len(cs) and h not in EXPLICIT_NOBULLET
                                and (DESC_H.match(h) or re.search(r'(条件|要求|说明|逻辑|措施|处置|要点|内容)$', h))):
                            sug('T11', '%s 第 %d 页：说明类句子列「%s」没有加「•」——整列左对齐加点（SD-102）' % (name, i + 1, h))
                except Exception: serial = False
                # T15 / T16（SD-188，2026-10-15 用户；待改清单第 7、12、14 条）——用户定的排版规则，报错误
                try:
                    _ct = lambda x: ''.join(l[2] for l in x[1]).strip() if x else ''
                    _ph = lambda t: not t or re.fullmatch(r'[—\-–/无\s]+', t)
                    _pad = (collections.Counter(round(l[0] - x[0][0]) for r in info[1:] for x in r if x for l in x[1] if l[2].strip()).most_common(1) or [(6, 0)])[0][0]
                    def _centered(x):
                        c, ls = x; ls = [l for l in ls if l[2].strip()]
                        return bool(ls) and all(l[0] - c[0] > _pad + 2.5 and abs((l[0] - c[0]) - (c[2] - l[1])) < 3 and (c[2] - c[0]) - (l[1] - l[0]) > 2 * _pad + 12 for l in ls)   # 2026-10-19 R4：文字几乎占满格宽（两侧空白合计不超过内边距 + 12pt）时靠左、居中看不出差别，不算居中（4.23「恢复后仍有一台发动机失效」误报）
                    # 2026-10-16：表首的通栏行（前提 / 警告条，只有首格、其余为合并格）不当表头
                    _inf = info
                    while len(_inf) > 2 and _inf[0] and len(_inf[0]) > 1 and _inf[0][0] and all(x is None for x in _inf[0][1:]): _inf = _inf[1:]
                    _hd = [_ct(x) for x in _inf[0]] if _inf else []
                    _body = [r for r in _inf[1:] if r and not (_ct(r[0]).startswith('注') and all(x is None for x in r[1:]))]
                    # T15：只有两列的表（非序号表），第二列每格靠左、首行带「•」
                    # SD-198（2026-10-16 用户）：首列表头标 no-cmp 的同类并列两列表（如 1.13 A-1 系统 A ｜ 系统 B）第二列居中不加点，T15 跳过
                    if len(_hd) == 2 and not serial and re.sub(r'\s', '', ''.join(_hd)) not in NOCMP_HDRS:
                        bad = [r[1] for r in _body if len(r) > 1 and r[1] and not _ph(_ct(r[1])) and (not _ct(r[1]).startswith('•') or _centered(r[1])) and not re.search(r'[：:]\s*$', (r[1][1][0][2] if r[1][1] else '')) and not re.match(r'\s*([\u2460-\u2473]|\d{1,2}[.、)）]\s)', _ct(r[1]))]   # 格内自带 ①② / 1. 编号的条目不另加点（每项一个标记）   # 引导语开头（SD-180「以下任一：」+「–」分项）不加点
                        if bad: err('T15', '%s 第 %d 页：两列表「%s」列有 %d 格没有靠左加点——两列表第二列一律靠左加「•」（SD-188 ②）' % (name, i + 1, _hd[1], len(bad)))
                    # T16：有跨列共有格的表，独有的文字格不居中（纯数值格除外）
                    _VAL = re.compile(r'[\d.,:：;；~～+\-−–≤≥<>%()（）/°×±℃℉a-zA-Z·•\s]*')
                    if len(_hd) >= 3 and not serial and re.sub(r'\s', '', ''.join(_hd)) not in NOCMP_HDRS and any(any(k >= 1 and r[k] and k + 1 < len(r) and r[k + 1] is None for k in range(len(r))) for r in _body):
                        cnt = 0
                        for r in _body:
                            for k in range(1, len(r)):
                                x = r[k]
                                if not x or (k + 1 < len(r) and r[k + 1] is None): continue
                                t = _ct(x)
                                tv = re.sub(r'^\s*(约|至少|最多|不少于|不超过|不大于|不小于)\s*', '', re.sub(r'[（(][^）)]*[）)]', '', t))   # 括注不算文字；「约 1s」「至少 5s」这类带量词的数值与生成器同口径按数值格（2026-10-19）
                                if _ph(t) or re.fullmatch(r'[•\s]*(是|否|有|无|适用|不适用)', t) or (re.search(r'\d', tv) and _VAL.fullmatch(tv)): continue
                                if _centered(x): cnt += 1
                        if cnt: err('T16', '%s 第 %d 页：有共有格的表里 %d 个独有文字格居中——共有格居中、独有格靠左（SD-188 ①）' % (name, i + 1, cnt))
                except Exception: pass
                # T4 并列长句未分条（SD-78）：一格内 ≥2 个「；」、各分句 ≥ 20 字宽、却没有 •
                for r in ([] if serial else info[1:]):
                    for x in r:
                        if not x: continue
                        txt = ''.join(l[2] for l in x[1]).strip()
                        if re.search(r'[•–▪①-⑳]', txt) or re.match(r'(注|解释|出处|公司差异)[：:]', txt): continue
                        # SD-203 起多项格项间统一「；」：行末的「；」是分行项之间的分号，不算一行里的并列分句
                        txt_nl = '\n'.join(l[2] for l in x[1]).strip()
                        parts = [s for s in re.split(r'；(?!\s*\n)', txt_nl) if s.strip()]
                        parts = [s.replace('\n', '') for s in parts]
                        if len(parts) >= 3 and all(vis(s) >= 20 for s in parts[:-1]):
                            sug('T4', '%s 第 %d 页：「%s…」含 %d 个并列长分句，考虑按语义分条加「•」' % (name, i + 1, txt[:20], len(parts)))
    _par(lambda rng, first: _tpass(d if first else pymupdf.open(pdf), rng), len(d))
    return d

# 2026-10-14 提速：单册的 scan 放进分叉子进程与全书同时跑；scan 只往 ERR / SUG 里添，子进程把新增条目原样传回、接在全书之后（顺序与串行相同）
_qpid = _qfd = None
if QON and hasattr(os, 'fork') and not os.environ.get('CHECK_SERIAL'):
    import pickle
    _n0 = (len(ERR), len(SUG)); _r, _w = os.pipe(); _qpid = os.fork()
    if _qpid == 0:
        os.close(_r)
        try: scan(QREF, '单册', 'B737机型理论速查'); _out = ('ok', ERR[_n0[0]:], SUG[_n0[1]:])
        except BaseException as _e: _out = ('fail', repr(_e), None)
        with os.fdopen(_w, 'wb') as _f: pickle.dump(_out, _f)
        os._exit(0)
    os.close(_w); _qfd = _r
book = scan(BOOK, '全书', r'(第[零一二三四五六七八九]章|前言|使用说明|总目录|目录|速查目录|按主题查|主题目录|附录)')   # 1019R6 使用说明页   # SD-209 L05：附录缩写表   # SD-96 页眉左侧为章名
if _qpid:
    with os.fdopen(_qfd, 'rb') as _f: _out = pickle.load(_f)
    os.waitpid(_qpid, 0)
    if _out[0] != 'ok': raise RuntimeError('单册 scan 子进程失败：' + _out[1])
    ERR.extend(_out[1]); SUG.extend(_out[2])
elif QON: scan(QREF, '单册', 'B737机型理论速查')   # SD-146 册名   # SD-139 单册停出：只有新近生成的单册才检查   # 单册页眉左侧为册名（2026-09-30 用户定；SD-96 右侧为块名）
# T8 末行孤字（SD-84 / SD-85，错误级）：用 layout_measure.measure() 实测——与 fit_fix 的自动修复看同一批格子，
# 检查器不再维护第二套测量（2026-10-05，APU 火警「1s」教训：两套眼睛必然漏）
try:
    from layout_measure import measure as _lm_measure
    for _nm, _pdf in (('全书', BOOK),) + ((('单册', QREF),) if QON else ()):
        _res = _lm_measure(_pdf, srcmd=(MDP if _nm == '全书' else os.path.join(REPO, 'build', 'qr_single.md')))   # SD-208：与 fit_fix 同口径传源文件（知道原文 <br>），否则把 <br> 处换行当折行、报假 T14
        # T13 列宽多余（SD-151，错误级）：与 fit_fix 实测收窄同口径（严格余量：本格行数不增）；跨页的表按各段最小余量（页底段 + 下页页顶续段，列数、表宽相同算同一张）
        _segs = collections.OrderedDict()
        for _x in _res:
            if _x.get('nw') and _x.get('col', -1) >= 0 and 'tb' in _x:
                _segs.setdefault((_x['page'], _x['tb']), []).append(_x)
        _keys = list(_segs); _grp = {}; _gid = 0
        for _q, _k in enumerate(_keys):
            _e = _segs[_k][0]
            _cand = [kk for kk in _keys[:_q] if kk[0] == _k[0] - 1 and (_segs[kk][0]['nc'] == _e['nc'] or (abs(_segs[kk][0]['nc'] - _e['nc']) == 1 and _segs[kk][0].get('head') and _segs[kk][0].get('head') == _e.get('head'))) and abs(_segs[kk][0]['tw'] - _e['tw']) < 3] if _e['tb'] < 70 else []   # 1021：续页一截偶尔多认一列（空合并格），首列表头相同也算同一张表
            _prev = _cand[-1] if _cand else None   # 上一页列数、表宽相同的那张（块索引可按块分页，不一定是上一页最后一张、也不一定排到页底）
            if _prev:
                _grp[_k] = _grp[_prev]
            else:
                _gid += 1; _grp[_k] = _gid
        _byg = collections.defaultdict(list)
        for _k in _keys: _byg[_grp[_k]].append(_k)
        _tocp = set()   # 2026-10-25（SD-237 目录两栏连续排）：目录页是两栏排版用的表、各行靠点线填满，不量列宽
        try:
            _dt = pymupdf.open(_pdf); _tocp = {i + 1 for i in range(min(10, len(_dt))) if _dt[i].get_text().lstrip().startswith('目录')}
        except Exception: pass
        for _g, _ks in _byg.items():
            if _ks[0][0] in _tocp: continue
            for _c in range(_segs[_ks[0]][0]['nc']):
                _vals = [next((e for e in _segs[_k] if e['col'] == _c), None) for _k in _ks]
                if any(v is None for v in _vals): continue
                _sl = min(v.get('extra_strict', v['extra_pt']) for v in _vals)
                _thr = max([th for sg, th in FIXW_SIGS if sg and re.sub(r'[^\w]', '', nosp(_vals[0].get('table', ''))).startswith(re.sub(r'[^\w]', '', sg)[:24])] or [10])   # 1020R2：多条签名命中取最宽松（图表并排的 10**6 优先于 w-NN 的 24）   # 锁宽 30pt、起始比例 24pt、自动宽度 10pt
                if _sl < 30 and any(nosp(_vals[0].get('table', '')).startswith(sg) and kk == _c for sg, kk in NWL): continue   # 已证明收到极限（再收就多行）；1020R1：余量 ≥ 30pt 的不认（收 24pt 不可能多行，旧记录误锁）
                _wz = any(re.sub(r'[^\w]', '', nosp(_vals[0].get('table', ''))).startswith(sg) and kk == _c for sg, kk in WZ)   # SD-232 ⑨：列宽实测调整来回跳、fit_fix 已冻结
                if _sl >= _thr and not any(v.get('wrapped') or v.get('anywrap') for v in _vals) and _wz:
                    sug('T13', '%s 第 %d 页：第 %d 列（「%s」）右侧仍空约 %.0fpt——fit_fix 实测调整来回跳、已冻结（列宽冻结记录），个别接受' % (_nm, _ks[0][0], _c + 1, _vals[0].get('head', ''), _sl))
                elif _sl >= _thr and not any(v.get('wrapped') or v.get('anywrap') for v in _vals):   # 只查没折行的列（用户原话「空白太多」）
                    err('T13', '%s 第 %d 页：第 %d 列（「%s」）没有折行，右侧仍空约 %.0fpt——按最长一行收窄（以文字成行为标准）' % (_nm, _ks[0][0], _c + 1, _vals[0].get('head', ''), _sl))
        # T14 有空位却折行（2026-10-09 用户，速查第 135 页「发动机限制、喘振或失速」：「在右侧还有空白区域空间的情况下……这一列可以适当加宽以减小行数」）：
        # 同一张表里短格（≤ 20 字）被迫折成两行，而别的列还有「本格行数不增」的空位、合计够它排成一行——定宽表（w-NN）也查（T13 跳过定宽表，fit_fix 不改定宽表，所以要在这里拦）
        _tcols = collections.defaultdict(dict)
        for _x in _res:
            if _x.get('nw') and _x.get('col', -1) >= 0: _tcols[(_x['page'], _x.get('table', ''))][_x['col']] = _x
        _t14 = set()
        for _x in _res:
            if _x.get('nw') or _x.get('grow') or _x.get('orphan') or _x.get('span') or not _x.get('cell'): continue
            _cols = _tcols.get((_x['page'], _x.get('table', '')))
            if not _cols or _x['col'] not in _cols: continue
            _don = [(k, v) for k, v in _cols.items() if k != _x['col'] and not v.get('anywrap') and v.get('extra_strict', 0) >= 6]
            _have = sum(v['extra_strict'] for k, v in _don)
            if _don and _have >= _x['extra_pt'] + 2 and (_x['page'], _x.get('table', '')) not in _t14:
                _t14.add((_x['page'], _x.get('table', '')))
                _dk, _dv = max(_don, key=lambda kv: kv[1]['extra_strict'])
                err('T14', '%s 第 %d 页：第 %d 列「%s」折成 2 行，第 %d 列（「%s」）右侧仍空约 %.0fpt——把空位让给折行的列（定宽表用 autow --from-report 重定）' % (
                    _nm, _x['page'], _x['col'] + 1, _x['cell'][:20], _dk + 1, _dv.get('head', ''), _have))
        # T8 末行孤字（SD-192，2026-10-16 用户：「可在不增加行数的情况下适当收窄其余列宽度以减少本列的孤字问题」）：
        # 同表另有列实测可让出（收窄后本列行数不增）≥ 所需宽度却留着孤字 → 错误（fit_fix 应借宽）；同表没有列能让 → 建议（SD-172 个别接受）
        # 可让宽度与 fit_fix 同口径：跨页的表按各段（同组）最小值；单列不够时几列合计（每列 ≥ 2pt 才算）
        _lendg = {}
        for _g, _ks in _byg.items():
            for _k in _ks:
                for _e in _segs[_k]:
                    _key = (_g, _e['col']); _lendg[_key] = min(_lendg.get(_key, 10 ** 9), _e.get('lend') or 0)
        _gof = {}
        for _k, _g in _grp.items():
            for _e in _segs[_k]: _gof[(_k[0], _e.get('table', ''))] = _g
        for _x in _res:
            if not _x.get('orphan'): continue
            _g = _gof.get((_x['page'], _x.get('table', '')))
            _don = sorted([(v, c) for (g, c), v in _lendg.items() if g == _g and c != _x['col'] and v >= 2], reverse=True) if _g else []
            _tried = any(nosp(_x.get('table', '')).startswith(b) for b in BWB)   # 借过宽、借出列因此出孤字已撤销 → 借不到
            _tbn = re.sub(r'[^\w]', '', nosp(_x.get('table', '')))
            _don = [(v, c) for v, c in _don if not any(_tbn.startswith(sg) and kk == c for sg, kk in WZ)]   # SD-232 ⑨：已冻结的列不算可让出
            _tried = _tried or any(_tbn.startswith(sg) and kk == _x['col'] for sg, kk in WZ)
            if _don and sum(v for v, c in _don) >= _x['extra_pt'] + 1 and not _x.get('span') and not _tried:   # 跨列格不借宽（fit_fix 同口径），按建议
                err('T8', '%s 第 %d 页：「%s…」末行只剩一两个字，需约 %.0fpt；同表%s 可让出（不增行）——向这些列借宽（SD-192）' % (
                    _nm, _x['page'], _x['cell'][:16], _x['extra_pt'], '、'.join('第 %d 列约 %.0fpt' % (c + 1, v) for v, c in _don[:3])))
            else:
                sug('T8', '%s 第 %d 页：「%s…」末行只剩一两个字（上一行止于「…%s」）——%s，个别接受（SD-172 / 192）' % (_nm, _x['page'], _x['cell'][:16], _x['prev'][-10:], '借宽后借出列出孤字已撤销' if _tried else '同表没有可让出宽度的列'))
        if _nm != '全书': continue
    # T14（续，2026-10-10 用户：速查第 23 条 N1 表）：定宽表右侧有空位、长句列却折行（fixw_grow.py 同口径；--apply 可直接改）
    try:
        from fixw_grow import scan as _fg_scan
        for (_f, _p), _v in _fg_scan(REPO).items():
            err('T14', '%s 第 %d 页：定宽表 %s 右侧有空位，%s——按实测加宽（python3 ../pub/tools/fixw_grow.py --repo . --apply）' % (
                _v['book'], _v['page'], os.path.basename(_f), '、'.join('第 %d 列可加宽约 %.0fpt' % (c + 1, e) for c, e in _v['cols'].items())))
    except Exception as _e:
        sug('T14', '定宽表加宽检查未能运行：%s' % _e)
    try:   # 1020R1（用户 2026-10-20 第 290 页、第 10 张截图）：自动宽度的表右侧有空位、某列却折行——fit_fix 应已加宽，报错误
        for (_f, _p), _v in _fg_scan(REPO, auto=True).items():
            err('T14', '%s 第 %d 页：表格右侧有空位，%s——加宽以减少行数（fit_fix 未加宽：查 keep_force 的 WB / NWL 记录或表头 w- 标注）' % (
                _v['book'], _v['page'], '、'.join('第 %d 列可加宽约 %.0fpt' % (c + 1, e) for c, e in _v['cols'].items())))
    except Exception as _e:
        sug('T14', '自动宽度表加宽检查未能运行：%s' % _e)
except Exception as _e:
    sug('T8', '末行孤字实测未能完成：%s' % _e)


# B4 一页只有一两行（2026-09-30 用户：「尽量避免在一页中只有一两行的情况」）：正文（去页眉页脚）不超过 2 行的页
for nm, pdf in (('全书', BOOK),) + ((('单册', QREF),) if QON else ()):
    dd = pymupdf.open(pdf); Hh = dd[0].rect.height
    for i in range(1, len(dd) - 1):
        ln = [l for l in body_lines(dd[i]) if l[1] > 0.06 * Hh]
        # 2026-10-11：整页插图（图 + 图注）只有图注一行文字，不算孤页
        if 0 < len(ln) <= 2 and dd[i].get_images():
            if sum(r.height for x in dd[i].get_images(full=True) for r in dd[i].get_image_rects(x[0])) > 0.35 * Hh: continue
        if 0 < len(ln) <= 2:
            err('B4', '%s 第 %d 页只有 %d 行（「%s」）——调整上一页间距或内容，避免孤页' % (nm, i + 1, len(ln), ln[0][4].strip()[:20]))

# P08 / P09 插图版面（1020R1，用户 2026-10-20：「图片的下面和左右空间都很大，为什么把图片放这么小」「尽量利用空白区域」）：
#   并排的图与表之间空 > 40pt → 错误（图应加大到填满，或表加宽）；独立小图（< 版心 35%）下方到页脚空 > 25% 版心高 → 建议（fit_fix 应已放大）
_FIXFIG = set()   # 第 96 / 103 条：图行标「定」（用户定过宽度）的图，K-P09 不报
for _mdf in (MDP, os.path.join(REPO, 'build', 'qr_single.md')):
    try:
        for _l in open(_mdf, encoding='utf-8'):
            _fm = re.match(r'%%FIG\w*%%[^|]*\|\s*(图\s*[\d.]+-\d+)', _l)
            if _fm and re.search(r'\|\s*定\s*(\||$)', _l): _FIXFIG.add(re.sub(r'\s', '', _fm.group(1)))
    except Exception: pass
for nm, pdf in (('全书', BOOK),) + ((('单册', QREF),) if QON else ()):
    dd = pymupdf.open(pdf)
    for i, pg in enumerate(dd):
        ims = [b for b in pg.get_text('dict')['blocks'] if b['type'] == 1 and (b['bbox'][2] - b['bbox'][0]) > 40]
        if not ims: continue
        W, H = pg.rect.width, pg.rect.height
        bl = [b for b in pg.get_text('blocks') if b[1] > 40 and b[3] < H - 45]
        bottom = max([b[3] for b in bl] + [im['bbox'][3] for im in ims])
        cap = next((b[4].replace('\n', '')[:20] for b in bl if re.match(r'\s*图 \d', b[4])), '')
        tabs = pg.find_tables().tables
        for im in ims:
            x0, y0, x1, y1 = im['bbox']
            side = [tb for tb in tabs if tb.bbox[1] < y1 and tb.bbox[3] > y0 and (tb.bbox[0] > x1 - 5 or tb.bbox[2] < x0 + 5)]
            if side:   # 空位量到并排一侧最远的文字（表或表上的引导段）；图已与表同高（≥ 85%）时不报——图再大就超出表高
                ty0, ty1 = min(tb.bbox[1] for tb in side), max(tb.bbox[3] for tb in side)
                right = min(tb.bbox[0] for tb in side) > x1
                txt = [b for b in bl if b[3] > y0 - 60 and b[1] < y1 and ((b[0] > x1 - 5) if right else (b[2] < x0 + 5)) and not re.match(r'\s*图 \d', b[4])]
                gap = (min([b[0] for b in txt] + [min(tb.bbox[0] for tb in side)]) - x1) if right else (x0 - max([b[2] for b in txt] + [max(tb.bbox[2] for tb in side)]))
                _cn = re.match(r'图\s*[\d.]+-\d+', cap); _cn = re.sub(r'\s', '', _cn.group(0)) if _cn else ''
                if gap > 40 and (y1 - y0) < 0.85 * (ty1 - ty0) and _cn not in _FIXFIG: err('P09', '%s 第 %d 页：并排的图（%s）与表之间空约 %.0fpt——加大图宽（源文件 FIGSIDE / FIGRIGHT 的 mm）填满空位' % (nm, i + 1, cap, gap))
            elif (x1 - x0) < 0.35 * (W - 72) and (H - 50 - bottom) > 0.25 * (H - 90):
                sug('P08', '%s 第 %d 页：插图（%s）只占版心 %.0f%% 宽，下方到页脚还空约 %.0fpt——加大图宽' % (nm, i + 1, cap, 100 * (x1 - x0) / (W - 72), H - 50 - bottom))

# K-USERPIN（第 116 条，2026-10-20 用户：说过的「定义」列居中 1020R1 没改到）：用户点名的排版要求逐条在成品上实测
if not SECTION and os.path.exists(os.path.join(REPO, '用户点名排版台账.json')):
    try:
        _up = subprocess.run([sys.executable, os.path.join(T, 'check_userpins.py'), REPO], capture_output=True, text=True, timeout=900).stdout
        for _l in _up.splitlines():
            if _l.startswith('USERPIN 错误 '): err('USERPIN', _l[len('USERPIN 错误 '):] + '——用户点名的排版要求（《用户点名排版台账.json》）')
            elif _l.startswith('USERPIN 提示 '): sug('USERPIN', _l[len('USERPIN 提示 '):])
    except Exception as _e:
        sug('USERPIN', '用户点名排版核对未能运行：%s' % _e)

# K-B5 留白页、K-Z3 过度压缩（SD-229 / 230，2026-10-23 用户：「为什么这一页还存在过度压缩的情况？明明下面还有空间」）：
# 原来只在交付前手动跑，现在并进本报告的错误级，每次全量自动拦
if not SECTION:
    _srcs = sorted(glob.glob(os.path.join(REPO, 'notes_src', '*', '*.md'))) + [os.path.join(REPO, '速查', '速查源.md')]
    for _pdf, _nm in ((BOOK, '全书'), (QREF, '单册')):
        if not os.path.exists(_pdf): continue
        for _tool, _code, _args in (('check_gaps.py', 'B5', []), ('check_overshrink.py', 'Z3', ['--src'] + _srcs)):
            try:
                _o = subprocess.run([sys.executable, os.path.join(T, _tool), _pdf] + _args, capture_output=True, text=True, timeout=900).stdout
                for _l in _o.splitlines():
                    if _l.startswith('第 ') and '页' in _l[:12]: err(_code, '%s %s' % (_nm, _l))
                    elif _l.startswith('提示 「'): sug(_code, '%s %s' % (_nm, re.sub(r'（第 \d+ 页）$', '', _l[3:])))   # SD-232：fit_fix 实测退一级会注掉页 / 拆表 / 多页的表，只提示（压缩退级记录_<书名>.json）
            except Exception as _e:
                sug(_code, '%s %s 未能运行：%s' % (_nm, _tool, _e))

# K-B6 排版漂移（SD-238，2026-10-26 1026R1：只删 6.5 第 5 条，前面 240 页正文版面却跟基线不同）：源文件没改的部分与最近的基线逐页比对
if not SECTION and os.path.exists(BOOK):
    try:
        _o = subprocess.run([sys.executable, os.path.join(T, 'check_drift.py'), '--repo', REPO, '--pdf', BOOK], capture_output=True, text=True, timeout=600).stdout
        _dl = [l for l in _o.splitlines() if l.startswith('第 ')]
        for _l in _dl[:10]: err('B6', '全书 ' + _l)
        if len(_dl) > 10: err('B6', '全书 另有 %d 页漂移（python3 check_drift.py 看全部）' % (len(_dl) - 10))
        for _l in _o.splitlines():
            if _l.startswith('提示 '): sug('B6', '全书 ' + _l[3:])
    except Exception as _e:
        sug('B6', '全书 check_drift.py 未能运行：%s' % _e)

# K-L8 页脚页码落点（SD-236，2026-10-24 用户：「点击右下角页码……返回到页码对应知识点的主题目录那一页」）：footer_toc_links.py --check 的不符 / 缺链接并进错误级
if not SECTION and os.path.exists(BOOK):
    try:
        _o = subprocess.run([sys.executable, os.path.join(T, 'footer_toc_links.py'), BOOK, '--check'], capture_output=True, text=True, timeout=600).stdout
        for _l in _o.splitlines():
            if _l.startswith('  ✗ '): err('L8', '全书 ' + _l[4:])
    except Exception as _e:
        sug('L8', '全书 footer_toc_links.py 未能运行：%s' % _e)

# K-L7 自绘图数值核对、K-RULES 规则自检（2026-10-21 规则重组轮 A1）：结果并进本报告的错误级
if not SECTION:
    for _tool, _tag, _code in (('check_figs.py', 'FIGS 错误 ', 'L7'), ('check_rules.py', 'RULES 错误 ', 'RULES')):
        try:
            _o = subprocess.run([sys.executable, os.path.join(T, _tool), REPO] + (['--no-copies'] if _code == 'RULES' else []), capture_output=True, text=True, timeout=300).stdout
            for _l in _o.splitlines():
                if _l.startswith(_tag): err(_code, _l[len(_tag):])
        except Exception as _e:
            sug(_code, '%s 未能运行：%s' % (_tool, _e))

# K-T38（Q06；2026-10-21 规则自检发现登记了没实现）：速查一条 ≤ 1 页、以表为主。按 PDF 实测条目内容高度（标题到下一条标题，
# 跨页时各页内容高度相加，不计页底留白）；含插图、one-page、split-ok 的条目按 Q06 例外不报。存量较多，暂为建议级，等用户定处理方式
if QON and not SECTION:
    try:
        import hashlib
        try: _KEEP = json.load(open(os.path.join(REPO, '速查超页保留.json'), encoding='utf-8'))
        except (OSError, ValueError): _KEEP = {}
        _ents = [e.split('\n%%PART%%')[0].split('\n## ')[0] for e in re.split(r'\n### ', open(os.path.join(REPO, '速查', '速查源.md'), encoding='utf-8').read())[1:]]
        _qd = pymupdf.open(QREF); _qh = _qd[0].rect.height; _pos = {}; _lns = {}
        for _pi in range(len(_qd)):
            _L = []
            for _b in _qd[_pi].get_text('dict')['blocks']:
                if _b['type'] == 1: _L.append((_b['bbox'][1], _b['bbox'][3])); continue
                for _l in _b.get('lines', []):
                    if not _l['spans'] or _l['bbox'][1] > _qh - 48: continue
                    _L.append((_l['bbox'][1], _l['bbox'][3]))
                    _m = re.match(r'(\d+)\.\s', ''.join(s['text'] for s in _l['spans']).strip())
                    if _m and abs(_l['spans'][0]['size'] - 10.5) < 0.3 and any('Bold' in s['font'] for s in _l['spans']):
                        _pos[int(_m.group(1))] = (_pi, _l['bbox'][1])
            _lns[_pi] = _L
        for _n in range(1, len(_ents) + 1):
            _e = _ents[_n - 1]; _t = _e.split('\n')[0].strip()
            if '<table' not in _e: err('T38', '速查第 %d 条「%s」没有表格——速查以表为主（Q06）' % (_n, _t[:24]))
            if _n not in _pos or _n + 1 not in _pos or any(k in _e for k in ('%%FIG', 'one-page', 'split-ok')): continue
            if _KEEP.get(_t, {}).get('指纹') == hashlib.sha1(re.sub(r' class="[^"]*"', '', _e).encode('utf-8')).hexdigest()[:12]: continue   # 2026-10-21：《速查超页保留.json》登记的例外（内容改了指纹对不上照报）
            (_p1, _y1), (_p2, _y2) = _pos[_n], _pos[_n + 1]; _h = 0
            for _pi in range(_p1, _p2 + 1):
                _L = [(a, b) for a, b in _lns[_pi] if (_pi > _p1 or a >= _y1 - 1) and (_pi < _p2 or a < _y2 - 1)]
                if _L: _h += max(b for a, b in _L) - min(a for a, b in _L)
            if _h > 525: sug('T38', '速查第 %d 条「%s」（第 %d 页起）内容约 %.1f 页——一条 ≤ 1 页，拆条或精简（Q06）' % (_n, _t[:24], _p1 + 1, _h / 521))
    except Exception as _e:
        sug('T38', '速查一条一页核对未能运行：%s' % _e)

# K-T41（原 K-T40，与序号表编号重号，2026-10-23 改号；2026-10-22 用户：速查第 95 页「放一行」）：「或」「且」「以及」等连接词单独占一行——并到上一行或改写成一句
for _nm, _dd in (('全书', book), ('单册', pymupdf.open(QREF) if os.path.exists(QREF) else None)):
    if _dd is None: continue
    for _i, _p in enumerate(_dd):
        for _b in _p.get_text('dict')['blocks']:
            for _l in _b.get('lines', []):
                _t = ''.join(_s['text'] for _s in _l['spans']).strip()
                if _t in ('或', '或者', '且', '以及', '和', '及', '并且', '或：'): err('T41', '%s 第 %d 页：连接词「%s」单独成行——并到上一行或改写成一句' % (_nm, _i + 1, _t))

# P 前言（SD-73 / SD-74）与总目录（单节预览不查）
pre = 0 if SECTION else next((i for i in range(1, 5) if nosp(book[i].get_text()).find('前言') >= 0), None)
if SECTION: pass
elif pre is None:   # SD-240（2026-10-26 版权处理）：前言.md 已删、不设前言页；只有前言文件还在时才报缺
    if os.environ.get('DOC_PREFACE') and os.path.exists(os.environ['DOC_PREFACE']): err('P1', '全书第 2～5 页找不到前言')
else:
    t = book[pre].get_text()
    if SIGN and SIGN not in t: err('P1', '前言缺署名 %s（SD-74）' % SIGN)
    if re.search(r'以\s*(SOP|FCOM)[^。]{0,6}为准', t): err('P1', '前言不应再有「以 SOP / FCOM 为准」一句（SD-73）')
toc = '主题目录' if SECTION else ''.join(book[i].get_text() for i in range(1, min(6, len(book))) if book[i].get_text().lstrip().startswith('目录'))   # 只看目录页（前言来源里的公司《运行手册》不算）
if re.search(r'运行手册|飞行技术提示|模拟机笔记|第四章\s*技术提示|06\s*技术提示', toc): err('P2', '目录仍出现旧章名（运行手册 / 飞行技术提示 / 模拟机笔记 / 技术提示）——2026-10-25 起为 运行规范 / 故障处置，SD-241 起第四章为「特殊情况与风险」')
if '速查目录' not in toc and '主题目录' not in toc: err('P2', '目录里缺「速查目录」（SD-233；原 SD-139 / 149「主题目录」）')
# SD-233（2026-10-23 两本整合）：「速查目录」放在总目录之后、第一章之前（取代 SD-149「主题目录放在最后一章之后」）
_tp = [i for i in range(1, len(book)) if re.match(r'(速查目录|主题目录)', nosp(book[i].get_text())[:12]) or nosp(book[i].get_text()).startswith(('速查目录', '主题目录'))]
_c1 = next((i for i in range(len(book)) if nosp(book[i].get_text()).find('第一章') >= 0 and i > 3 and not re.search(r'目录', nosp(book[i].get_text())[:20])), len(book))
if SECTION: pass
elif not _tp: err('P2', '找不到「速查目录」页（SD-233）')
elif _tp[0] > _c1: err('P2', '「速查目录」在第 %d 页，应放在总目录之后、第一章之前（SD-233）' % (_tp[0] + 1))

# R5（SD-154，2026-10-06 用户）：同一条目下同表头的几张表、中间只隔着注——应合成一张（标 one-page）。
# 例外：压到最小仍放不下一页、用户定为照常跨页的几处（范围 A）
R5_EXEMPT = {('*', '电气故障处置'), ('*', '各类液压故障处置'), ('*', 'GPS 干扰对各系统的影响'), ('*', '起飞程序职责交接与推力管理矩阵')}   # 2026-10-23（SD-233 改结构后 4.8 / 4.14 / 1.11 / 2.7 改了号，节号写死的豁免落空）：只按条目标题片段认，不看节号   # 2026-10-19 批 5：改用标题片段，条目改号后不失效   # 2026-10-19 SD-209 批 3：GNSS 干扰 5.12 第 3 条 → 1.11 第 29 条   # 2026-10-15 SD-187：液压相关由 4.9 改为 4.14；SD-189：GNSS 干扰由 5.6 改为 5.12
import glob as _g
for _f in sorted(_g.glob(os.path.join(REPO, 'notes_src/[1-6]*/*.md'))):
    _sec = os.path.basename(_f).split(' ')[0]; _s = open(_f, encoding='utf-8').read()
    for _m in re.finditer(r'(?ms)^#{3,4} ([^\n]*)\n(.*?)(?=^#{2,4} |\Z)', _s):
        _id = _m.group(1).split('　')[0].split(' ')[0]
        if (_sec, _id) in R5_EXEMPT or any(_a in ('*', _sec) and len(_b) > 3 and _b in _m.group(1) for _a, _b in R5_EXEMPT): continue
        _tb = list(re.finditer(r'(?s)<table[^>]*>.*?</table>', _m.group(2)))
        _hd = lambda t: re.sub(r'<[^>]+>|\s', '', (re.search(r'<tr class="hdr">(.*?)</tr>', t, re.S) or [0, ''])[1])
        for _a, _b in zip(_tb, _tb[1:]):
            _gap = [p.strip() for p in re.split(r'\n\s*\n', _m.group(2)[_a.end():_b.start()]) if p.strip()]
            if _hd(_a.group(0)) and _hd(_a.group(0)) == _hd(_b.group(0)) and all(p.startswith('注') for p in _gap):
                err('R5', '%s「%s」有两张同表头的表、中间只隔注——合成一张、注移到表下、标 one-page（SD-154）' % (_sec, _m.group(1).strip()[:24]))

# V 成品通用校验（verify.py：空白页、标签泄漏、异常项目符号、front matter）
for pdf in (BOOK,) + ((QREF,) if QON else ()):
    r = _R('verify.py', pdf)
    if r.returncode: err('V1', '%s verify.py 未通过：%s' % (os.path.basename(pdf), ' / '.join(l for l in r.stdout.splitlines() if '[]' not in l)))

# B 断表 / 孤行 / 标题孤立（SD-71 / SD-77 / SD-79）
for pdf in (BOOK,) + ((QREF,) if QON else ()):
    r = _R('check_splits.py', pdf)
    for l in r.stdout.splitlines()[1:]:
        l = l.strip()
        rule = 'B3' if '页底只有标题' in l else 'B1'   # SD-218：B1 坏断点（截行不足 / 页底只剩表头 / 页底分组行 / 断在合并格），B3 标题孤立
        err(rule, '%s %s' % ('全书' if pdf == BOOK else '单册', l))

# Z 压缩表清单（SD-80）：fit_fix.py 为「略超一页」的表选的压缩级别，供人工抽看是否仍清晰易读
for kf in sorted(glob.glob(os.path.join(REPO, 'build', 'keep_force_*.txt'))):
    book_name = os.path.basename(kf)[11:-4]
    for l in open(kf, encoding='utf-8'):
        l = l.strip()
        if l.startswith('~'):   # SD-218：只有点名整页（one-page）的表还会压缩
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
