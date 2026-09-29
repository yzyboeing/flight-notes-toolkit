#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_quickref.py —— 速查区与正文数值一致性校验（SD-23）

用法：
    python3 tools/check_quickref.py [--src notes_src]

速查区每条标题下的「详见 x.y Z-n」指明它抄自正文哪一条。本脚本：
  1  核对每个「详见」地址在正文里确实存在（改了编号没同步会在这里报）
  2  把速查区条目里的每个数值（连同单位）拿去正文对应条目里找；
     找不到再到该节全文里找；全节都没有的数值列为「疑似不一致」
  3  否定词核对：速查区与正文里「去掉 不 / 未 / 无 / 非 / 禁止 后相同」的分句，
     否定词个数却不同 → 报错（例：「不根据 X 调整」↔「根据 X 调整」）
  4  行对应核对：表头相同的表，按（第一列, 第二列）对齐行，比较其余短格
     （指示状态、数值、是否点亮…）→ 对不上报「疑似行对应错位」；
     速查区有而正文没有的（第一列, 第二列）组合也报出（合并 / 拆分单元格时最易出错）
只报不改。以哪一边为准，交用户按原始资料 / 现行手册裁定。
局限：只能查「速查区有、正文没有」的数；正文改了而旧值仍出现在该节别处时查不出来。
"""
import io, os, re, sys, glob

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
SRC = os.path.abspath(_arg('--src', 'notes_src'))

UNIT = r'(?:%|kt|ft|fpm|psi|nm|NM|km|m|kg|lb|LB|℃|°C|°|g|min|s|h|Hz|V|hPa|mb|秒|分钟|小时)?'
NUM  = re.compile(r'((?:[<>≤≥±]\s?)?(?:(?<![\d.])[-−])?(?<![\d.])\d+(?:\.\d+)?)\s?(' + UNIT[3:-2] + r')?')   # 数字前紧挨的 – 是区间号，不当负号

def plain(s):
    s = re.sub(r'<small>.*?</small>', ' ', s, flags=re.S)   # SD-33 灰色出处 / 解释不参与数值比对
    s = re.sub(r'<[^>]+>', ' ', s)
    s = s.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
    return re.sub(r'−\s+(?=\d)', '−', s)            # 「− 5%」与「−5%」视为同一个数

def tokens(s):
    out = set()
    s = re.sub(r'B-\d{4}[A-Z]?(\s*[–\-]\s*B-\d{4}[A-Z]?)?', ' ', s)   # 飞机注册号段（SD-32）不当作数值
    s = re.sub(r'第\s*\d+\s*条', ' ', s)                                  # 条目互引不当作数值
    for m in NUM.finditer(plain(s)):
        v = re.sub(r'\s', '', m.group(1)).replace('−', '-').replace('–', '-')
        u = (m.group(2) or '').replace('°C', '℃')
        if re.fullmatch(r'[<>≤≥±]?-?\d', v) and not u: continue     # 单个数字（序号、1 台）不查
        out.add((v.lstrip('<>≤≥±'), u))
    return out

NEG = re.compile(r'禁止|不|未|无|非|没|勿')
PUNCT = re.compile(r'[\s，。；;、：:（）()「」“”"\'·/—–\-→>＞<＜=≥≤±+×*【】\[\]|！!？?]')

def clauses(html):
    s = re.sub(r'</t[dh]>|<br\s*/?>|</tr>', '。', html)
    s = plain(s)
    out = []
    for c in re.split(r'[。；;\n]|，(?=[^，]{6,})', s):
        c = PUNCT.sub('', c)
        if len(c) >= 6: out.append(c)
    return out

def grids(html):
    """表格 → [(表头规范化列表, 行网格)]，rowspan / colspan 展开"""
    res = []
    for tb in re.findall(r'(?s)<table\b.*?</table>', html):
        hdr, rows, occ = None, [], {}
        for cls, body in re.findall(r'(?s)<tr([^>]*)>(.*?)</tr>', tb):
            cells = re.findall(r'(?s)<t([hd])([^>]*)>(.*?)</t[hd]>', body)
            if 'hdr' in cls:
                if hdr is None: hdr = [PUNCT.sub('', plain(x[2])) for x in cells]
                continue
            if 'note' in cls or 'premise' in cls or 'warn' in cls: continue
            ri = len(rows); row = {}; ci = 0
            for _, attr, txt in cells:
                while (ri, ci) in occ: row[ci] = occ[(ri, ci)]; ci += 1
                rs = int((re.search(r'rowspan="(\d+)"', attr) or [0, 1])[1])
                cs = int((re.search(r'colspan="(\d+)"', attr) or [0, 1])[1])
                v = PUNCT.sub('', plain(txt))
                for k in range(cs):
                    row[ci + k] = v
                    for r2 in range(1, rs): occ[(ri + r2, ci + k)] = v
                ci += cs
            while (ri, ci) in occ: row[ci] = occ[(ri, ci)]; ci += 1
            rows.append(row)
        if hdr and rows: res.append((hdr, rows))
    return res

def polarity_issues(item_html, src_html):
    src = clauses(src_html); src_set = set(src)
    idx = {}
    for s in src: idx.setdefault(NEG.sub('', s), []).append(s)
    out = []
    for c in clauses(item_html):
        if c in src_set: continue
        for s in idx.get(NEG.sub('', c), []):
            if len(NEG.findall(c)) != len(NEG.findall(s)):
                out.append('「%s」↔ 正文「%s」' % (c[:24], s[:24])); break
    return out

def row_issues(item_html, src_html):
    """按第一列找正文同名行；第二列是短标签时要求能对上，其余短格要求相等或互相包含"""
    from difflib import SequenceMatcher
    norm = lambda s: s.replace('ftmin', 'fpm')
    out = []
    sg = grids(src_html)
    for hdr, rows in grids(item_html):
        if len(hdr) < 3: continue
        best_iss = None
        for shdr, srows in sg:
            if shdr[:2] != hdr[:2]: continue
            iss = []
            by0 = {}
            for r in srows: by0.setdefault(r.get(0, ''), []).append(r)
            for r in rows:
                cand = by0.get(r.get(0, ''))
                if not cand: continue
                c1 = r.get(1, '')
                best = max(cand, key=lambda s: SequenceMatcher(None, c1, s.get(1, '')).ratio())
                if c1 and len(c1) <= 30 and SequenceMatcher(None, c1, best.get(1, '')).ratio() < 0.5:
                    iss.append('行「%s｜%s」在正文同名表里对不上（最接近：%s）' % (r.get(0, '')[:10], c1[:16], best.get(1, '')[:16])); continue
                same1 = [s for s in cand if s.get(1, '') == best.get(1, '')]
                for ci, h in enumerate(hdr):
                    if ci < 2 or h not in shdr: continue
                    a = norm(r.get(ci, ''))
                    bs = [norm(s.get(shdr.index(h), '')) for s in same1]
                    if not a or len(a) > 24 or not any(bs): continue
                    if any(a == b or a in b or b in a for b in bs if b): continue
                    iss.append('行「%s｜%s」的「%s」：%s ↔ 正文 %s' % (r.get(0, '')[:10], c1[:12], h, a, bs[0]))
            if best_iss is None or len(iss) < len(best_iss): best_iss = iss
        out += best_iss or []
    return out

def main():
    qf = glob.glob(os.path.join(SRC, '0 *', '0 *.md'))
    if not qf: sys.exit('找不到速查区源文件')
    q = io.open(qf[0], encoding='utf-8').read()
    items = re.findall(r'(?ms)^### (\d+)\. ([^\n]*)\n(.*?)(?=^### |^## |\Z)', q)
    secs, blocks = {}, {}
    for f in glob.glob(os.path.join(SRC, '[1-5]*', '*.md')):
        t = io.open(f, encoding='utf-8').read(); sid = os.path.basename(f).split(' ')[0]
        secs[sid] = t
        for m in re.finditer(r'(?ms)^#{3,4} ([A-Z]-\d+)　[^\n]*\n(.*?)(?=^#{3,4} |\Z)', t):
            blocks[(sid, m.group(1))] = m.group(2)
        if sid.split('.')[0] in ('4', '5'):          # 第 4、5 章条目「### N. 标题」，地址写「第 N 条」（SD-30）
            for m in re.finditer(r'(?ms)^### (\d+)\. [^\n]*\n(.*?)(?=^### |\Z)', t):
                blocks[(sid, '第 %s 条' % m.group(1))] = m.group(2)
    errs, warns, n_ptr, n_own = [], [], 0, 0
    for n, title, body in items:
        mp = re.search(r'详见 \[\[[^\]|]+\|(\d+\.\d+) ([A-Z]-\d+|第 \d+ 条)\]\]', body)
        if not mp: n_own += 1; continue
        n_ptr += 1
        key = (mp.group(1), mp.group(2))
        if key not in blocks:
            errs.append('第 %s 条：详见 %s %s 在正文里不存在' % (n, *key)); continue
        mine = tokens(re.sub(r'<small>.*?</small>', ' ', re.sub(r'(?m)^\s*<!--[\s\S]*?-->\s*\n', '', re.sub(r'(?m)^(来源|详见|出处：|解释：|公司差异：)[^\n]*\n', '', body)), flags=re.S))
        blk, whole = tokens(blocks[key]), tokens(secs[key[0]])
        bare_whole = {v for v, _ in whole}
        miss = sorted({'%s%s' % x for x in mine if x not in blk and x not in whole and x[0] not in bare_whole})
        if miss: warns.append('第 %s 条（%s）↔ %s %s：正文里找不到 %s' % (n, title[:16], *key, '、'.join(miss[:8])))
        for x in polarity_issues(body, blocks[key]):
            errs.append('第 %s 条 ↔ %s %s 否定词不一致：%s' % (n, *key, x))
        # 整节找同表头的表：一条速查可能汇总本节相邻两条，「详见」只指其中一条
        for x in row_issues(body, secs[key[0]]):
            warns.append('第 %s 条 ↔ %s %s 疑似行对应错位：%s' % (n, *key, x))
    print('速查区核对（数值 / 否定词 / 行对应）：%d 条有「详见」，%d 条为速查区独有' % (n_ptr, n_own))
    for w in warns: print('  疑似不一致  ' + w)
    for e in errs: print('  错误  ' + e)
    if not errs and not warns: print('  全部通过')
    sys.exit(1 if errs else 0)

if __name__ == '__main__':
    main()
