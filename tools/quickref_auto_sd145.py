#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""quickref.py —— 由完整版自动生成《速查版》（2026-10-05 用户：「保留一个速查版的手册……所有知识点的一个集合，
随着完整版更新，我需要的时候你再给我」；用户选「每块一条要点」）

做法：读 build/book.md（完整版拼好的源），每一节一张表「条目 ｜ 要点」，全书每个知识点块一行：
  · 要点＝这一块里带红（<em>，限制 / 门槛）或蓝（<b>，要背的数值）标记的原句，按「；」「。」「<br>」切成短句后原样保留（不改字、不截半句）；
    表格行的短句前加行名（「行名：……」）；同一块最多取 MAXC 句；
  · 没有红蓝标记的块，取第一句原文作要点；
  · 速查版独立成册：不写正文页码、「另 N 项」等指向完整版的内容（2026-10-05 用户：「速查版是独立的，所以不用添加正文项目或者相关」）。
不改 notes_src、不进 git；只生成 build/quickref.md 和 build/B737机型理论知识笔记速查版.docx / .pdf。
用法：python3 quickref.py [--repo ~/flight-repos/gh-private]
"""
import sys, os, re, html, subprocess, shutil, tempfile

T = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1]) if '--repo' in sys.argv else os.path.expanduser('~/flight-repos/gh-private')
BUILD = os.path.join(REPO, 'build')
BOOK_MD = os.path.join(BUILD, 'book.md')
BOOK_PDF = os.path.join(BUILD, 'B737机型理论知识笔记.pdf')
OUT_MD = os.path.join(BUILD, 'quickref.md')
OUT_DOCX = os.path.join(BUILD, 'B737机型理论知识笔记速查版.docx')
MAXC = 4          # 每块最多取几句
MAXLEN = 70       # 单句过长（视觉宽度，汉字算 2）的不进要点，避免整段搬运

plain = lambda s: html.unescape(re.sub(r'<[^>]+>', '', s)).strip()
def vis(s): return sum(2 if ord(ch) > 0x2E80 else 1 for ch in plain(s))

def split_top(s, seps='；。'):
    """在标签外、括号外按 ；。<br> 切句；切点落在 <em>/<b>/<strong> 或（）内部的不切。"""
    out, buf, depth, par, i = [], '', 0, 0, 0
    while i < len(s):
        if s.startswith('<br', i):
            j = s.find('>', i) + 1
            if depth == 0: out.append(buf); buf = ''
            else: buf += s[i:j]
            i = j; continue
        if s[i] == '<':
            j = s.find('>', i) + 1; tag = s[i:j]
            if re.match(r'</(em|b|strong)\b', tag): depth = max(0, depth - 1)
            elif re.match(r'<(em|b|strong)\b', tag): depth += 1
            buf += tag; i = j; continue
        ch = s[i]; buf += ch
        if ch in '（(': par += 1
        elif ch in '）)': par = max(0, par - 1)
        elif ch in seps and depth == 0 and par == 0: out.append(buf); buf = ''
        i += 1
    out.append(buf)
    return [x.strip() for x in out if plain(x).strip('；。 ')]

def balanced(s):
    for t in ('em', 'b', 'strong'):
        if len(re.findall(r'<%s\b' % t, s)) != len(re.findall(r'</%s>' % t, s)): return False
    return True

def clean(s):
    s = re.sub(r'[-]', '', s)
    s = re.sub(r'<(?!/?(em|b|strong)\b)[^>]+>', '', s)          # 只保留颜色 / 加粗标记
    s = re.sub(r'^\s*[-–—•]\s*', '', s).strip().rstrip('；。;')
    if not balanced(s): s = html.escape(plain(s), quote=False)
    return s

SERIAL = re.compile(r'^\s*([\u2460-\u2473]|\d{1,2}[.、]?|[（(]\s*(\d{1,2}|[a-zA-Z])\s*[)）]|[a-zA-Z][.、)）])\s*$')

def block_points(lines):
    """一块的原文行 → 要点短句列表：红蓝标记句优先，其次黑粗句，最后取第一整句。"""
    marked, bold, first = [], [], None
    rows_hdr = None
    for l in lines:
        if not l.strip() or l.startswith('%%') or l.startswith('<!--'): continue
        if '<table' in l or '</table>' in l and not l.startswith('<tr'): continue
        if l.startswith('<tr'):
            if 'class="hdr"' in l: continue
            cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', l, re.S)
            if not cells: continue
            label = plain(cells[0]); rest = cells[1:] if len(cells) > 1 else cells
            for c in rest:
                for seg in split_top(c):
                    p = plain(seg)
                    if not p or p in ('—', '-'): continue
                    lab = label if (label and len(label) <= 14 and label not in p and len(cells) > 1 and not SERIAL.match(label)) else ''
                    item = (lab + '：' if lab else '') + clean(seg)
                    if vis(item) > MAXLEN * 2: continue
                    if re.search(r'<(em|b)\b', seg): marked.append(item)
                    elif '<strong>' in seg: bold.append(item)
            if first is None:   # 第一整句：首个数据行，行名＋第一个内容格（只按「。」和 <br> 切）
                for c in rest:
                    segs = split_top(c, seps='。')
                    if segs:
                        lab = label if (label and len(label) <= 14 and len(cells) > 1 and not SERIAL.match(label)) else ''
                        cand = (lab + '：' if lab else '') + clean(segs[0])
                        if vis(cand) <= MAXLEN * 2: first = cand
                        break
            continue
        txt = re.sub(r'^#+\s*', '', l)
        for seg in split_top(txt):
            item = clean(seg)
            if not plain(item) or vis(item) > MAXLEN * 2: continue
            if re.search(r'<(em|b)\b', seg): marked.append(item)
            elif '<strong>' in seg: bold.append(item)
        if first is None:
            segs = split_top(txt, seps='。')
            if segs:
                cand = clean(segs[0])
                if plain(cand) and vis(cand) <= MAXLEN * 2: first = cand
    seen, uniq = set(), []
    if not marked: marked = bold        # 没有红蓝标记：用黑粗的关键句
    for m in marked:
        k = plain(m)
        if k not in seen: seen.add(k); uniq.append(m)
    if uniq:
        pts = uniq[:MAXC]
        while len(pts) > 1 and plain(pts[-1]).endswith(('：', ':')): pts.pop()   # 末条是引导句（下文被截掉）就不要
        return pts   # 速查版独立成册（用户 2026-10-05）：不写「另 N 项」、正文页等指向完整版的内容
    return [first] if first else []

def build_md():
    src = open(BOOK_MD, encoding='utf-8').read().split('\n')
    out = ['# 机型理论知识速查', '', '%%PAGEBREAK%%', '']
    chap = sec = None; blocks = []; cur = None
    def flush_sec():
        if sec is None or not blocks: return
        out.append('<table class="ftn split-ok">')
        out.append('<tr class="hdr"><th class="col-left">条目</th><th class="col-left">要点</th></tr>')
        for bid, title, body in blocks:
            pts = block_points(body)
            cell = '<br>'.join(pts) if pts else '—'
            out.append('<tr><td><strong>%s</strong>　%s</td><td>%s</td></tr>' % (bid, html.escape(title, quote=False), cell))
        out.append('</table>'); out.append('')
    for l in src:
        m2 = re.match(r'^## (第[一二三四五六七八九十]+章.*)$', l)
        m3 = re.match(r'^### (\d+\.\d+)[\s　]+(.*)$', l)
        m4 = re.match(r'^#### ([A-H]-\d+|\d+)[\s　.．]+(.*)$', l)
        if m2:
            if cur: blocks.append(cur); cur = None
            flush_sec(); blocks = []; sec = None
            chap = m2.group(1); out += ['', '## ' + chap, '']; continue
        if m3:
            if cur: blocks.append(cur); cur = None
            flush_sec(); blocks = []
            sec = (m3.group(1), m3.group(2)); out += ['### %s　%s' % sec, '']; continue
        if l.startswith('#### '):
            if cur: blocks.append(cur); cur = None
            if m4 and sec: cur = (m4.group(1), re.sub(r'（单位：[^）]*）', '', m4.group(2)).strip(), [])
            continue
        if cur is not None: cur[2].append(l)
    if cur: blocks.append(cur)
    flush_sec()
    open(OUT_MD, 'w', encoding='utf-8').write('\n'.join(out) + '\n')
    return sum(1 for l in out if l.startswith('<tr><td>'))

def to_docx_pdf():
    env = dict(os.environ)
    for k in ('DOC_TOPICS', 'DOC_PREFACE', 'DOC_QRTOPICS', 'DOC_TOPICINDEX'): env.pop(k, None)
    env['DOC_SUBTITLE'] = '速查版 · 全书知识点要点'
    env['NO_SEC_BREAK'] = '1'   # 速查版每节表格都短：节与节连排，不另起一页
    env['KEEP_FORCE'] = os.path.join(BUILD, 'keep_force_B737机型理论知识笔记速查版.txt')
    r = subprocess.run(['node', os.path.join(T, 'build_docx.js'), OUT_MD, OUT_DOCX], env=env, cwd=REPO)
    if r.returncode: sys.exit('速查版 docx 生成失败')
    sof = shutil.which('soffice') or '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    prof = os.path.join(os.environ.get('TMPDIR', '/tmp'), 'lo-sync-profile')
    subprocess.run([sof, '-env:UserInstallation=file://' + prof, '--headless', '--convert-to', 'pdf', '--outdir', BUILD, OUT_DOCX], capture_output=True)

if __name__ == '__main__':
    n = build_md()
    to_docx_pdf()
    pdf = OUT_DOCX[:-5] + '.pdf'
    try:
        import pymupdf; np = len(pymupdf.open(pdf))
    except Exception: np = '?'
    print('速查版：%d 个知识点块，%s 页 → %s' % (n, np, pdf))
