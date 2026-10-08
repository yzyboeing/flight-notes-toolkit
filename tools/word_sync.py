#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_sync.py —— 用户直接在 Word 成品上改稿 → 找出改动 → 对回源文件（2026-10-07 用户，SD-167）

用户原话：「能不能我直接去改 Word 文件，然后你来检查哪里做了变更？……具体的文字增删由我来做，
你最后只负责排版的恢复和优化。」——Word 只是出版成品；改动对回 notes_src / 速查源 后重新出版，排版自然恢复。

用法：
  python3 word_sync.py <改稿.docx> [--base <基线.docx>] [--repo <笔记库>] [--out <报告.md>] [--apply] [--force]

  - 默认只读：列出全部改动（文字增删、整段 / 整行删除、新增段落、颜色 / 加粗修改、批注、自动生成部分的改动）
    和每条在源文件里的位置。
  - --apply：能在源文件里唯一定位的文字增删、整段删除、整行删除直接写进源文件；未改部分的颜色 / 加粗原样保留，
    新加的字按 Word 里的颜色补标签。其余列为「需人工」，由 Claude 按报告处理。目标文件有未提交改动时拒绝（--force 跳过）。
  - 比对方式：找得到同版本号的原版 Word（默认在 用户改稿/基线、Muse交接包 里找）就逐段 / 逐格对比，开没开「修订」都能找全；
    找不到时只读「修订」记录。批注（Word「新建批注」）一律读出，作为排版 / 整体意见。
  - 文件名含「速查」→ 对应 速查/速查源.md；否则对应 notes_src/ 第一～五章。
  - 封面、前言、目录、块索引、按主题查都是自动生成的：这些地方的改动只报告，按排版意见处理。
"""
import argparse, bisect, difflib, glob, html, os, re, subprocess, sys, zipfile
from collections import Counter
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
RED = {'C00000', 'FF0000', 'C0504D', 'E00000', 'CC0000'}
BLUE = {'0B5CAD', '0070C0', '1F4E79', '2E75B6', '0000FF'}
ZW = '​‌‍﻿⁠'   # U+2060：生成器在「以上 / 以下」等词中间插的不换行连接符
BULLET = re.compile(r'^[\s​]*[●•▪◦–][\s　]*')
HOME = os.path.expanduser('~')
BASE_DIRS = [os.path.join(HOME, 'Desktop/飞行理论笔记整理/用户改稿/基线'), os.path.join(HOME, 'Desktop/飞行理论笔记整理/Muse交接包'),
             os.path.join(HOME, 'Desktop/飞行理论笔记整理')]
TAGS = {'em': ('<em>', '</em>'), 'b': ('<b>', '</b>'), 'strong': ('<strong>', '</strong>')}
MARK_CN = {'em': '红', 'b': '蓝', 'strong': '加粗'}


def nrm(s):
    return ''.join(c for c in s if not c.isspace() and c not in ZW)


def marks_of(rpr):
    if rpr is None:
        return frozenset()
    m = set()
    c = rpr.find(W + 'color')
    if c is not None:
        v = (c.get(W + 'val') or '').upper()
        if v in RED: m.add('em')
        elif v in BLUE: m.add('b')
    b = rpr.find(W + 'b')
    if b is not None and (b.get(W + 'val') or 'true').lower() not in ('false', '0', 'off') and not m:
        m.add('strong')
    return frozenset(m)


def mark_str(ms):
    return '、'.join(MARK_CN[x] for x in sorted(ms)) or '黑'


# ================================================================ 读 Word
class Comments:
    def __init__(self, z):
        self.items, self.open = {}, set()
        if 'word/comments.xml' in z.namelist():
            for cm in ET.fromstring(z.read('word/comments.xml')).iter(W + 'comment'):
                paras = [''.join(t.text or '' for t in p.iter(W + 't')) for p in cm.iter(W + 'p')]
                self.items[cm.get(W + 'id')] = {'author': cm.get(W + 'author') or '', 'text': '\n'.join(x for x in paras if x.strip()),
                                                'anchor': '', 'loc': ''}

    def start(self, cid, loc):
        if cid in self.items:
            self.open.add(cid); self.items[cid]['loc'] = self.items[cid]['loc'] or loc

    def end(self, cid):
        self.open.discard(cid)

    def ref(self, cid, loc):
        if cid in self.items and not self.items[cid]['loc']:
            self.items[cid]['loc'] = loc

    def feed(self, s):
        for cid in self.open:
            self.items[cid]['anchor'] += s


def read_para(p, cm, loc):
    """一个段落 → 旧文（修订里删除的字算、插入的不算）/ 新文，及每个字的颜色 / 加粗"""
    U = {'old': [], 'new': [], 'rev': False, 'style': ''}
    ppr = p.find(W + 'pPr')
    if ppr is not None:
        st = ppr.find(W + 'pStyle')
        U['style'] = st.get(W + 'val') if st is not None else ''
        rp = ppr.find(W + 'rPr')
        if rp is not None and (rp.find(W + 'del') is not None or rp.find(W + 'ins') is not None):
            U['rev'] = True

    def run(r, mode):
        rpr = r.find(W + 'rPr')
        mk = marks_of(rpr)
        chg = rpr.find(W + 'rPrChange') if rpr is not None else None
        old_mk = marks_of(chg.find(W + 'rPr')) if chg is not None else mk
        if chg is not None: U['rev'] = True
        for ch in r:
            t = None
            if ch.tag in (W + 't', W + 'delText'): t = ch.text or ''
            elif ch.tag == W + 'tab': t = '\t'
            elif ch.tag in (W + 'br', W + 'cr'): t = '\n'
            elif ch.tag == W + 'noBreakHyphen': t = '-'
            elif ch.tag == W + 'sym' and ch.get(W + 'char'):
                try: t = chr(int(ch.get(W + 'char'), 16))
                except ValueError: t = ''
            elif ch.tag == W + 'commentReference':
                cm.ref(ch.get(W + 'id'), loc)
            if not t: continue
            if mode != 'ins': U['old'].extend((c, old_mk) for c in t)
            if mode != 'del': U['new'].extend((c, mk) for c in t)
            cm.feed(t)

    def walk(el, mode):
        for ch in el:
            tg = ch.tag
            if tg in (W + 'ins', W + 'moveTo'): U['rev'] = True; walk(ch, 'ins')
            elif tg in (W + 'del', W + 'moveFrom'): U['rev'] = True; walk(ch, 'del')
            elif tg == W + 'r': run(ch, mode)
            elif tg == W + 'commentRangeStart': cm.start(ch.get(W + 'id'), loc)
            elif tg == W + 'commentRangeEnd': cm.end(ch.get(W + 'id'))
            elif tg != W + 'pPr': walk(ch, mode)
    walk(p, None)
    return U


def read_docx(path, quick):
    """→ 段落列表（正文段落、表格每格的每段），每段带位置、旧文 / 新文、颜色"""
    z = zipfile.ZipFile(path)
    body = ET.fromstring(z.read('word/document.xml')).find(W + 'body')
    cm = Comments(z)
    units, ctx = [], {'h2': '', 'h3': '', 'gen': True, 'tbl': 0, 'ver': ''}

    def loc_str(extra=''):
        return (' '.join(x for x in (ctx['h2'], ctx['h3']) if x) + ' ' + extra).strip()

    def strip_prefix(lst, rx):
        s = ''.join(c for c, _ in lst); m = rx.match(s)
        return lst[len(m.group(0)):] if m else lst

    def add(U, kind, extra, pos, row=''):
        U['old'], U['new'] = strip_prefix(U['old'], BULLET), strip_prefix(U['new'], BULLET)
        if quick and U['style'] == 'Heading3':   # 速查知识点标题前的「93. 」是自动编号
            U['old'], U['new'] = strip_prefix(U['old'], re.compile(r'^\s*\d+\.\s*')), strip_prefix(U['new'], re.compile(r'^\s*\d+\.\s*'))
        big_no = re.fullmatch(r'\s*\d{1,2}\s*', ''.join(c for c, _ in U['new']) or ' ') is not None and kind == 'p'   # 章首页的大号「02」
        U.update(kind=kind, loc=loc_str(extra), pos=pos, row=row, h2=ctx['h2'], h3=ctx['h3'], gen=ctx['gen'] or U.get('gen', False) or big_no)
        U['o'] = ''.join(c for c, _ in U['old']); U['om'] = [m for _, m in U['old']]
        U['n'] = ''.join(c for c, _ in U['new']); U['nm'] = [m for _, m in U['new']]
        U['prev'] = ctx.get('tail', '')
        t = nrm(U['n'] or U['o'])
        if t and not U['gen']: ctx['tail'] = t[-8:]
        units.append(U)

    def heading(U):
        st = U['style']; t = nrm(''.join(c for c, _ in U['new']) or ''.join(c for c, _ in U['old']))
        if st == 'Heading1': ctx.update(h2='', h3='', tbl=0, gen=True)   # 章首页（本章内容、章目录）/ 速查目录：自动生成，到第一个节标题为止
        elif st == 'Heading2': ctx.update(h2=t, h3='', tbl=0, gen=False)
        elif st == 'Heading3': ctx.update(h3=re.sub(r'^\d+\.', '', t) if quick else t, tbl=0)
        elif not quick and t == '按主题查': ctx['gen'] = True
        mv = re.search(r'版本号\s*(\d{4}R\d+)', ''.join(c for c, _ in U['new']))
        if mv and not ctx['ver']: ctx['ver'] = mv.group(1)

    def table(tb):
        ctx['tbl'] += 1; tno = ctx['tbl']
        rows = tb.findall(W + 'tr')
        first = nrm(''.join(t.text or '' for tc in rows[0].findall(W + 'tc') for t in tc.iter(W + 't'))) if rows else ''
        gen_tbl = first.startswith('块主题条目') or ctx['h3'] == '块索引'   # 每节开头的块索引（跨页时分成几张表）
        for r, tr in enumerate(rows):
            trpr = tr.find(W + 'trPr'); rf = ''
            if trpr is not None:
                rf = 'del' if trpr.find(W + 'del') is not None else ('ins' if trpr.find(W + 'ins') is not None else '')
            col = 0
            for tc in tr.findall(W + 'tc'):
                tcpr = tc.find(W + 'tcPr'); gs = tcpr.find(W + 'gridSpan') if tcpr is not None else None
                for k, p in enumerate(tc.findall(W + 'p')):
                    extra = '表%d 第%d行第%d列' % (tno, r + 1, col + 1)
                    U = read_para(p, cm, loc_str(extra))
                    U['gen'] = gen_tbl
                    add(U, 'cell', extra, (tno, r, col, k), rf)
                col += int(gs.get(W + 'val') or 1) if gs is not None else 1

    def visit(el):
        for ch in el:
            if ch.tag == W + 'p':
                U = read_para(ch, cm, loc_str()); heading(U); add(U, 'p', '', None)
            elif ch.tag == W + 'tbl':
                table(ch)
            elif ch.tag in (W + 'sdt', W + 'sdtContent', W + 'customXml'):
                visit(ch)
    visit(body)
    return units, cm.items, ctx['ver']


# ================================================================ 源文件：去标签、去空白的投影，每个字记回原位置
def project(raw):
    out, mp, i, n, bol = [], [], 0, len(raw), True
    if raw.startswith('---\n'):
        e = raw.find('\n---', 4)
        if e > 0: i = raw.find('\n', e + 4) + 1 or n
    while i < n:
        c = raw[i]
        if bol and raw.startswith('%%', i):                 # 生成指令行（插图、分段）
            j = raw.find('\n', i); i = n if j < 0 else j; continue
        if bol and c == '#':
            j = i
            while j < n and raw[j] == '#': j += 1
            if j < n and raw[j] == ' ': i = j + 1; bol = False; continue
        if raw.startswith('<!--', i):
            j = raw.find('-->', i); i = n if j < 0 else j + 3; continue
        if c == '<':
            m = re.match(r'</?[a-zA-Z][^<>]*>', raw[i:i + 300])
            if m: i += len(m.group(0)); continue
        if c == '&':
            m = re.match(r'&(#\d+|#x[0-9a-fA-F]+|[a-zA-Z]+);', raw[i:i + 12])
            if m:
                ch = html.unescape(m.group(0))
                if ch and not ch.isspace(): out.append(ch); mp.append((i, i + len(m.group(0))))
                i += len(m.group(0)); bol = False; continue
        if c == '-' and raw[i + 1:i + 2] == ' ' and (bol or re.search(r'<br\s*/?>\s*$', raw[max(0, i - 12):i])):
            i += 2; continue
        bol = (c == '\n')
        if not c.isspace() and c not in ZW:
            out.append(c); mp.append((i, i + 1))
        i += 1
    return ''.join(out), mp


class Corpus:
    def __init__(self, repo, quick):
        fs = [os.path.join(repo, '速查', '速查源.md')] if quick else sorted(glob.glob(os.path.join(repo, 'notes_src', '[1-5]*', '*.md')))
        self.files = []
        for f in fs:
            raw = open(f, encoding='utf-8').read()
            nm, mp = project(raw)
            starts = [s for s, _ in mp]
            heads = [(nrm(m.group(1)), bisect.bisect_left(starts, m.start())) for m in re.finditer(r'^#{2,3} (.+)$', raw, re.M)]
            self.files.append({'file': f, 'raw': raw, 'norm': nm, 'map': mp, 'heads': heads})

    def head_at(self, fi, pos):
        h = ''
        for t, p in self.files[fi]['heads']:
            if p <= pos: h = t
            else: break
        return h

    def find(self, key):
        hits = []
        for fi, f in enumerate(self.files):
            k = f['norm'].find(key)
            while k >= 0:
                hits.append((fi, k)); k = f['norm'].find(key, k + 1)
        return hits


def locate(C, un, a, b, h2, h3, quick, prev=''):
    """在源文件投影里找这条改动（规格化后 un[a:b]），返回 (文件号, 起, 止, '') 或 (None, None, None, 原因)"""
    if not un and not prev:
        return None, None, None, '空段落'
    sid = (re.match(r'(\d\.\d{1,2})', h2) or [None, ''])[1] if h2 else ''
    last = 0
    def narrow(hits):
        if len(hits) > 1 and sid and not quick:
            h = [x for x in hits if os.path.basename(C.files[x[0]]['file']).startswith(sid + ' ')]
            hits = h or hits
        if len(hits) > 1 and h3:
            h = [x for x in hits if h3[-10:] in C.head_at(x[0], x[1])]
            hits = h or hits
        return hits
    for L in (10, 20, 40, 10 ** 6):
        lo, hi = max(0, a - L), min(len(un), b + L)
        cands = [(un[lo:hi], lo)] + ([(prev + un[lo:hi], lo - len(prev))] if prev else [])   # 多处时再带上前一段的结尾
        found_any = False
        for key, off in cands:
            if not key: continue
            hits = narrow(C.find(key))
            found_any = found_any or bool(hits)
            if len(hits) == 1:
                fi, p = hits[0]
                return fi, p + (a - off), p + (b - off), ''
            if hits: last = len(hits)
        if not found_any:
            return None, None, None, '源文件里找不到（原版之后源文件已改过此处，或是自动生成的文字）'
    return None, None, None, '源文件里有 %d 处相同文字，无法唯一定位' % last


# ================================================================ 比对
TOK = re.compile(r'[0-9A-Za-z.%°]')


def char_edits(o, n, nm):
    sm = difflib.SequenceMatcher(None, o, n, autojunk=False)
    ops = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        while i1 > 0 and j1 > 0 and TOK.match(o[i1 - 1]) and o[i1 - 1] == n[j1 - 1] and (TOK.match(o[i1:i1 + 1] or ' ') or TOK.match(n[j1:j1 + 1] or ' ')):
            i1 -= 1; j1 -= 1                              # 改动落在数字 / 英文词中间：扩到整个词
        while i2 < len(o) and j2 < len(n) and TOK.match(o[i2]) and o[i2] == n[j2] and (TOK.match(o[i2 - 1:i2] or ' ') or TOK.match(n[j2 - 1:j2] or ' ')):
            i2 += 1; j2 += 1
        if ops and i1 < ops[-1]['i2']: i1 = ops[-1]['i2']; j1 = max(j1, ops[-1].get('j2', j1))
        if tag == 'equal' or (not nrm(o[i1:i2]) and not nrm(n[j1:j2])): continue   # 只改了空白的不算
        if ops and i1 - ops[-1]['i2'] <= 2:          # 相隔 0～2 个字的小改动合成一条
            p = ops[-1]; gap = o[p['i2']:i1]
            p.update(old=p['old'] + gap + o[i1:i2], new=p['new'] + gap + n[j1:j2], i2=i2, nmk=p['nmk'] + nm[j1:j2])
        else:
            ops.append({'i1': i1, 'i2': i2, 'j2': j2, 'old': o[i1:i2], 'new': n[j1:j2], 'nmk': list(nm[j1:j2])})
        ops[-1]['j2'] = j2
    return ops


def fmt_in_edited(o, n, om, nm):
    """文字也改了的段落里，没改的字有没有改颜色 / 加粗"""
    out, cur = [], None
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, o, n, autojunk=False).get_opcodes():
        if tag != 'equal': cur = None; continue
        for k in range(i2 - i1):
            a_, b_ = om[i1 + k], nm[j1 + k]; c = o[i1 + k]
            if a_ != b_ and not c.isspace():
                if cur and cur['end'] == i1 + k and cur['from'] == a_ and cur['to'] == b_:
                    cur['text'] += c; cur['end'] += 1
                else:
                    cur = {'text': c, 'from': a_, 'to': b_, 'end': i1 + k + 1}; out.append(cur)
            else:
                cur = None
    return out


def fmt_edits(o, om, nm):
    """文字没变、颜色 / 加粗变了的连续片段"""
    out, cur = [], None
    for k, c in enumerate(o):
        if om[k] != nm[k] and not c.isspace():
            if cur and cur['end'] == k and cur['from'] == om[k] and cur['to'] == nm[k]:
                cur['text'] += c; cur['end'] = k + 1
            else:
                cur = {'text': c, 'from': om[k], 'to': nm[k], 'end': k + 1}; out.append(cur)
    return out


def pair_units(A, B):
    """原版段落序列 A 与改稿段落序列 B 对齐 → [(a 或 None, b 或 None)]"""
    ka = [(u['gen'], nrm(u['o'])) for u in A]; kb = [(u['gen'], nrm(u['o'])) for u in B]
    out = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ka, kb, autojunk=False).get_opcodes():
        if tag == 'equal':
            out += [(A[i1 + k], B[j1 + k]) for k in range(i2 - i1)]; continue
        j = j1
        for i in range(i1, i2):
            best, bj = 0.0, None
            for jj in range(j, min(j2, j + 30)):
                r = difflib.SequenceMatcher(None, ka[i][1], kb[jj][1], autojunk=False).ratio()
                if r > best: best, bj = r, jj
                if r > 0.9: break
            if bj is not None and best >= 0.45:
                out += [(None, B[x]) for x in range(j, bj)] + [(A[i], B[bj])]; j = bj + 1
            else:
                out.append((A[i], None))
        out += [(None, B[x]) for x in range(j, j2)]
    return out


# ================================================================ 写入
def tags_open_at(raw, pos):
    st = []
    for m in re.finditer(r'<(/?)(em|b|strong)>', raw[raw.rfind('\n', 0, pos) + 1:pos]):
        if not m.group(1): st.append(m.group(2))
        elif m.group(2) in st: st.remove(m.group(2))
    return set(st)


def place_insert(raw, pos, text, want):
    """在 raw[pos] 插入 text，颜色 / 加粗与插入处不一致时：先试挪到紧跟的收尾标签之后，不行就补 / 拆标签"""
    have = tags_open_at(raw, pos)
    if want == have:
        return raw[:pos] + text + raw[pos:], ''
    m = re.match(r'(?:</(?:em|b|strong)>)+', raw[pos:])
    if m and tags_open_at(raw, pos + len(m.group(0))) == want:
        p2 = pos + len(m.group(0)); return raw[:p2] + text + raw[p2:], ''
    s = text
    for t in sorted(want - have): s = TAGS[t][0] + s + TAGS[t][1]
    drop, note = have - want, ''
    if drop:
        s = ''.join(TAGS[t][1] for t in sorted(drop)) + s + ''.join(TAGS[t][0] for t in sorted(drop))
        note = '新加的字是%s、插入处是%s，已拆开标签，请核对' % (mark_str(frozenset(want)), mark_str(frozenset(have)))
    return raw[:pos] + s + raw[pos:], note


def clean_around(raw, pos):
    """只清理改动所在的那一行：空标签、连续 <br>、格首 / 格尾 <br>、只剩「注：」的行"""
    ls = raw.rfind('\n', 0, pos) + 1; le = raw.find('\n', pos); le = len(raw) if le < 0 else le
    line = raw[ls:le]
    line = re.sub(r'<(em|b|strong)>(\s*)</\1>', r'\2', line)
    line = re.sub(r'(?:<br>\s*){2,}', '<br>', line)
    line = re.sub(r'<br>\s*(</td>)', r'\1', line)
    line = re.sub(r'(<td[^>]*>)\s*<br>', r'\1', line)
    line = line.rstrip(' \t')
    if re.fullmatch(r'注：\s*(?:——)?\s*', line):
        line = ''
    return raw[:ls] + line + raw[le:]


def find_base(path, ver):
    m = re.match(r'(B737机型理论知识(?:笔记|速查))(\d{4}R\d+)?', os.path.basename(path))
    if not m: return None
    v = ver or m.group(2) or ''
    for d in BASE_DIRS:
        f = os.path.join(d, '%s%s.docx' % (m.group(1), v))
        if v and os.path.exists(f) and os.path.abspath(f) != os.path.abspath(path):
            return f
    return None


# ================================================================ 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx'); ap.add_argument('--base'); ap.add_argument('--repo', default=os.path.join(HOME, 'flight-repos/gh-private'))
    ap.add_argument('--out'); ap.add_argument('--apply', action='store_true'); ap.add_argument('--force', action='store_true')
    a = ap.parse_args()
    quick = '速查' in os.path.basename(a.docx)
    B, comments, ver = read_docx(a.docx, quick)
    has_rev = any(u['rev'] for u in B)
    base = a.base or find_base(a.docx, ver)
    if base:                                       # 逐段对比原版：old 取原版，new 取改稿「接受全部修订后」的文字
        A, _, bver = read_docx(base, quick)
        for u in B: u['o'], u['om'] = u['n'], u['nm']
        pairs = pair_units(A, B)
        mode = '逐段 / 逐格对比原版 Word（%s）%s' % (os.path.basename(base), '；改稿里有修订记录，已按「接受全部修订后」的文字比对' if has_rev else '')
        if bver and ver and bver != ver: mode += '；注意：原版 %s 与改稿 %s 版本号不同' % (bver, ver)
    else:                                          # 只读修订记录
        pairs = []
        for u in B:
            if u['o'] == u['n'] and u['om'] == u['nm']: continue
            nu = dict(u, o=u['n'], om=u['nm'])
            pairs.append((u if nrm(u['o']) else None, nu if nrm(u['n']) else None))
        mode = '只读修订记录（没找到原版 Word；没开修订的改动看不出来）'
        if not has_rev:
            print('!! 没找到原版 Word，改稿里也没有修订记录：无法判断改了哪里。请用 --base 指定原版。', file=sys.stderr)
    C = Corpus(a.repo, quick)
    rel = lambda f: os.path.relpath(f, a.repo)

    T, D, I, F, G, R = [], [], [], [], [], []      # 文字改动、整段删除、新增段落、格式、自动生成部分、整行删除
    del_cells, prev = {}, ''
    for ua, ub in pairs:
        u = ub or ua
        if ua is not None and ub is not None:
            o, n = ua['o'], ub['o']
            if o == n:
                fe = fmt_edits(o, ua['om'], ub['om'])
                if fe: (G if u['gen'] else F).append({'loc': ub['loc'], 'chg': fe})
            elif u['gen']:
                G.append({'loc': ub['loc'], 'old': o, 'new': n})
            else:
                un = nrm(o)
                fe = fmt_in_edited(o, n, ua['om'], ub['om'])
                if fe: F.append({'loc': ub['loc'], 'chg': fe})
                for e in char_edits(o, n, ub['om']):
                    x0 = len(nrm(o[:e['i1']])); x1 = x0 + len(nrm(e['old']))
                    fi, s, t, why = locate(C, un, x0, x1, ua['h2'], ua['h3'], quick, ua.get('prev') or prev)
                    want = Counter(m for m in e['nmk'] if m is not None).most_common(1)
                    T.append({'loc': ub['loc'], 'old': e['old'], 'new': e['new'], 'want': set(want[0][0]) if want else set(),
                              'fi': fi, 's': s, 't': t, 'why': why})
            prev = nrm(o)[-8:]
        elif ua is not None:
            if u['gen']: G.append({'loc': ua['loc'], 'old': ua['o'], 'new': '（删除）'}); continue
            un = nrm(ua['o'])
            if not un: continue
            if ua['kind'] == 'cell': del_cells.setdefault((ua['h2'], ua['h3'], ua['pos'][0], ua['pos'][1]), []).append(ua)
            fi, s, t, why = locate(C, un, 0, len(un), ua['h2'], ua['h3'], quick, ua.get('prev') or prev)
            D.append({'loc': ua['loc'], 'old': ua['o'], 'u': ua, 'fi': fi, 's': s, 't': t, 'why': why})
        else:
            if u['gen']: G.append({'loc': ub['loc'], 'old': '（新增）', 'new': ub['o']}); continue
            if nrm(ub['o']): I.append({'loc': ub['loc'], 'new': ub['o'], 'after': prev})

    if not base:                                   # 只读修订：Word 标了「整行删除」的行
        for key, dl in del_cells.items():
            if dl and all(x.get('row') == 'del' for x in dl):
                R.append({'loc': dl[0]['loc'].split(' 表')[0] + ' 表%d 第%d行' % (key[2], key[3] + 1), 'cells': [x['o'] for x in dl]})
                D[:] = [d for d in D if d['u'] not in dl]
    if base:                                       # 整行删除：原版这一行的格子全被删了
        base_rows = {}
        for u in A:
            if u['kind'] == 'cell' and not u['gen'] and nrm(u['o']):
                base_rows.setdefault((u['h2'], u['h3'], u['pos'][0], u['pos'][1]), []).append(u)
        for key, dl in del_cells.items():
            allu = base_rows.get(key, [])
            if allu and len(dl) >= len(allu):
                R.append({'loc': allu[0]['loc'].split(' 表')[0] + ' 表%d 第%d行' % (key[2], key[3] + 1), 'cells': [x['o'] for x in allu]})
                D[:] = [d for d in D if d['u'] not in dl]

    applied, manual = [], []
    if a.apply:
        applied, manual = apply_all(a, C, T, D, R, rel)

    report(a, ver, mode, quick, C, rel, T, D, R, I, F, G, comments, applied, manual)


def apply_all(a, C, T, D, R, rel):
    applied, manual, ops = [], [], {}
    for x in T:
        if x['fi'] is None: manual.append(('文字', x['loc'], x['old'], x['new'], x['why'])); continue
        mp = C.files[x['fi']]['map']
        if x['t'] > x['s']:
            start, end = mp[x['s']][0], mp[x['t'] - 1][1]
        else:                                          # 纯插入：放在前一个字之后
            start = end = mp[x['s'] - 1][1] if x['s'] > 0 else mp[0][0]
        ops.setdefault(x['fi'], []).append({'s': start, 'e': end, 'new': x['new'], 'want': x['want'], 'info': ('文字', x['loc'], x['old'], x['new'])})
    for x in D:
        if x['fi'] is None: manual.append(('整段删除', x['loc'], x['old'], '', x['why'])); continue
        f = C.files[x['fi']]; mp = f['map']
        start, end = mp[x['s']][0], mp[x['t'] - 1][1]
        ls = f['raw'].rfind('\n', 0, start) + 1; le = f['raw'].find('\n', end); le = len(f['raw']) if le < 0 else le
        whole = nrm(project(f['raw'][ls:le])[0]) == nrm(project(f['raw'][start:end])[0])
        op = {'s': ls, 'e': min(le + 1, len(f['raw'])), 'new': '', 'want': set(), 'line': True} if whole else {'s': start, 'e': end, 'new': '', 'want': set()}
        op['info'] = ('整段删除', x['loc'], x['old'], '')
        ops.setdefault(x['fi'], []).append(op)
    for r in R:
        hit = []
        for fi, f in enumerate(C.files):
            for m in re.finditer(r'<tr\b[^>]*>.*?</tr>', f['raw'], re.S):
                rn = nrm(project(m.group(0))[0])
                if all(nrm(c) in rn for c in r['cells']) and len(rn) <= sum(len(nrm(c)) for c in r['cells']) + 4:
                    hit.append((fi, m))
        if len(hit) != 1 or re.search(r'rowspan|colspan', hit[0][1].group(0)):
            manual.append(('整行删除', r['loc'], ' ｜ '.join(r['cells']), '', '找不到唯一的一行，或该行有合并格')); continue
        fi, m = hit[0]; raw = C.files[fi]['raw']
        ops.setdefault(fi, []).append({'s': m.start(), 'e': m.end() + (1 if raw[m.end():m.end() + 1] == '\n' else 0), 'new': '', 'want': set(), 'line': True,
                                       'info': ('整行删除', r['loc'], ' ｜ '.join(r['cells']), '')})
    targets = sorted({C.files[fi]['file'] for fi in ops})
    if targets and not a.force:
        dirty = subprocess.run(['git', '-C', a.repo, 'status', '--porcelain', '--'] + [rel(f) for f in targets], capture_output=True, text=True).stdout.strip()
        if dirty:
            print('!! 目标源文件有未提交的改动，先提交或加 --force：\n' + dirty, file=sys.stderr); sys.exit(2)
    for fi, lst in ops.items():
        f = C.files[fi]; raw = f['raw']
        lst.sort(key=lambda z: (z['s'], z['e']))
        keep = []
        for op in lst:
            if keep and op['s'] < keep[-1]['e']:
                manual.append(op['info'][:4] + ('与上一条改动重叠',)); continue
            keep.append(op)
        for op in reversed(keep):                     # 从后往前写，前面的位置不受影响
            s, e = op['s'], op['e']
            if op.get('line') and raw[e:e + 1] == '\n' and raw[max(0, s - 2):s] == '\n\n':
                e += 1                                    # 删掉的整行前后都是空行：顺带去掉一个空行
            kept = '' if op.get('line') else ''.join(re.findall(r'<[^<>]+>', raw[s:e]))   # 删掉的范围里只去字、留标签
            raw = raw[:s] + kept + raw[e:]
            note = ''
            if op['new']:
                raw, note = place_insert(raw, s, op['new'], op['want'])
            if not op.get('line'):
                raw = clean_around(raw, s)
            applied.append(op['info'] + (rel(f['file']) + ('（%s）' % note if note else ''),))
        open(f['file'], 'w', encoding='utf-8').write(raw)
    return applied, manual


def report(a, ver, mode, quick, C, rel, T, D, R, I, F, G, comments, applied, manual):
    esc = lambda s: (s or '').replace('|', '｜').replace('\n', '↵')
    def pos(x):
        if x.get('fi') is None: return esc(x.get('why', '—'))
        f = C.files[x['fi']]; k = min(x['s'], len(f['map']) - 1)
        return '`%s:%d`' % (rel(f['file']), f['raw'].count('\n', 0, f['map'][k][0]) + 1)
    L = ['# Word 改稿比对报告', '',
         '- 改稿：`%s`%s' % (os.path.basename(a.docx), '（版本号 %s）' % ver if ver else ''),
         '- 比对方式：%s' % mode,
         '- 对应源文件：%s' % ('速查/速查源.md' if quick else 'notes_src/ 第一～五章'),
         '- 统计：文字改动 %d 处、整段删除 %d 处、整行删除 %d 处、新增段落 %d 处、颜色 / 加粗修改 %d 处、批注 %d 条、自动生成部分改动 %d 处'
         % (len(T), len(D), len(R), len(I), sum(len(x['chg']) for x in F), len(comments), len(G)), '']
    if T:
        L += ['## 一、文字改动', '', '| # | 位置 | 原文 → 新文 | 源文件 |', '|---|---|---|---|']
        L += ['| %d | %s | %s → %s | %s |' % (k, esc(x['loc']), esc(x['old']) or '（插入）', esc(x['new']) or '（删除）', pos(x)) for k, x in enumerate(T, 1)] + ['']
    if D or R:
        L += ['## 二、整段 / 整行删除', '', '| # | 位置 | 删除的内容 | 源文件 |', '|---|---|---|---|']
        L += ['| %d | %s | %s | %s |' % (k, esc(x['loc']), esc(x['old'][:100]), pos(x)) for k, x in enumerate(D, 1)]
        L += ['| 行 | %s | %s | 整行 |' % (esc(r['loc']), esc(' ｜ '.join(r['cells']))[:150]) for r in R] + ['']
    if I:
        L += ['## 三、新增段落（Claude 按位置写进源文件，并按规则上色）', '', '| # | 位置 | 新增内容 |', '|---|---|---|']
        L += ['| %d | %s | %s |' % (k, esc(x['loc']), esc(x['new'][:200])) for k, x in enumerate(I, 1)] + ['']
    if F:
        L += ['## 四、颜色 / 加粗修改', '', '| # | 位置 | 文字 | 原 → 新 |', '|---|---|---|---|']
        k = 0
        for x in F:
            for c in x['chg']:
                k += 1; L.append('| %d | %s | %s | %s → %s |' % (k, esc(x['loc']), esc(c['text']), mark_str(c['from']), mark_str(c['to'])))
        L.append('')
    if comments:
        L += ['## 五、批注（排版 / 整体意见）', '', '| # | 位置 | 批注内容 | 批注对象 |', '|---|---|---|---|']
        L += ['| %d | %s | %s | %s |' % (k, esc(c['loc']), esc(c['text']), esc(c['anchor'][:60]))
              for k, (cid, c) in enumerate(sorted(comments.items(), key=lambda z: int(z[0]) if z[0].isdigit() else 0), 1)] + ['']
    if G:
        L += ['## 六、自动生成部分的改动（目录 / 块索引 / 按主题查等，按排版意见处理）', '', '| # | 位置 | 改动 |', '|---|---|---|']
        for k, x in enumerate(G, 1):
            if 'chg' in x: L.append('| %d | %s | 格式：%s |' % (k, esc(x['loc']), '；'.join('%s %s→%s' % (esc(c['text']), mark_str(c['from']), mark_str(c['to'])) for c in x['chg'])))
            else: L.append('| %d | %s | %s → %s |' % (k, esc(x['loc']), esc(x['old'][:60]), esc(x['new'][:60])))
        L.append('')
    if a.apply:
        L += ['## 七、写入结果', '', '已写入 %d 处；需人工 %d 处。' % (len(applied), len(manual)), '']
        L += ['- 已写入（%s）%s：%s → %s　`%s`' % (x[0], esc(x[1]), esc(x[2][:40]) or '（插入）', esc(x[3][:40]) or '（删除）', x[4]) for x in applied]
        L += ['- **需人工**（%s）%s：%s → %s（%s）' % (x[0], esc(x[1]), esc(x[2][:40]), esc(x[3][:40]), esc(x[4])) for x in manual] + ['']
    rep = '\n'.join(L)
    if a.out:
        os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
        open(a.out, 'w', encoding='utf-8').write(rep)
    print(rep)


if __name__ == '__main__':
    main()
