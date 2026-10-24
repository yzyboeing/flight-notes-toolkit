#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fixosc.py —— 把 fit_fix 里「列宽来回跳」的表按成品实测宽度定宽（w-NN）（2026-10-12 用户：「固定来回跳的列宽」）
用法：python3 fixosc.py [--apply]        （在私有库根目录运行；先跑过一次 sync.sh --full，build/ 里要有 fitfix_trace_*.jsonl）
做法：
  1. 读 build/fitfix_trace_<册名>.jsonl（fit_fix 逐遍记录的各表实际列宽 W 与成品取哪一遍）；
  2. 来回跳＝最后几遍里某张表的列宽出现 A → B → A，或到最后一遍仍在变；
  3. 取成品那一遍的列宽，按版心宽度换算成百分比（最大余数法取整，满宽表合计 100）；
  4. 按 build_docx.js 同一算法算表格签名（全表文字去标点前 80 字），在源文件里唯一定位这张表，把 w-NN 写进表头；
     对不上、不唯一、表头有跨 3 列以上的格、已定宽的表只列出不改。
定宽后这张表不再参与 fit_fix 实测加宽 / 收窄（满宽定宽表），排版结果每次一致。"""
import os, re, sys, json, glob, html, unicodedata

APPLY = '--apply' in sys.argv
B = 'build'
BOOKS = {'B737机型理论速查笔记': sorted(glob.glob('notes_src/*/*.md')), 'B737机型理论速查': ['速查/速查源.md']}

def sig_of(tbl):
    """与 build_docx.js 同口径：表头前的前提行、表尾的通栏注解 / 警示行移出表格，不进签名"""
    rows = []
    for attr, body in re.findall(r'<tr\b([^>]*)>(.*?)</tr>', tbl, re.S):
        cls = (re.search(r'class="([^"]*)"', attr) or [0, ''])[1]
        rows.append((cls, re.findall(r'<t[dh]\b[^>]*>(.*?)</t[dh]>', body, re.S)))
    fd = next((k for k, (c, _) in enumerate(rows) if not re.search(r'hdr|premise', c)), len(rows))
    lastD = max((k for k, (c, _) in enumerate(rows) if not re.search(r'note|warn|premise', c)), default=-1)
    full = lambda c, cs: 'premise' in c or (re.search(r'note|warn', c) and len(cs) == 1)
    keep = [cs for k, (c, cs) in enumerate(rows) if not (full(c, cs) and (('premise' in c and k < fd) or (re.search(r'note|warn', c) and k > lastD)))]
    t = ''.join(html.unescape(re.sub(r'<[^>]+>', '', x)) for cs in keep for x in cs)
    return ''.join(ch for ch in t if unicodedata.category(ch)[0] in 'LN')[:80]

def osc_tables(trace):
    passes, best = {}, None
    for l in open(trace, encoding='utf-8'):
        if not l.strip(): continue
        o = json.loads(l)
        if 'pass' in o: passes[o['pass']] = o['W']
        if 'best' in o: best = o['best']
    if not passes: return {}, None, 0
    ks = sorted(passes)
    best = best if best in passes else ks[-1]
    out = {}
    for sg in passes[best]:
        seq = [passes[k].get(sg) for k in ks if passes[k].get(sg)]
        tail = seq[-5:]
        bounce = any(tail[i] == tail[i - 2] and tail[i] != tail[i - 1] for i in range(2, len(tail)))
        moving = len(seq) >= 2 and seq[-1] != seq[-2]
        if bounce or moving: out[sg] = passes[best][sg]
    total = max(sum(w) for w in passes[best].values() if w)
    return out, best, total

def pct_of(W, total):
    raw = [w * 100.0 / total for w in W]
    want = 100 if sum(W) >= total * 0.985 else int(round(sum(raw)))
    base = [int(x) for x in raw]
    rest = sorted(range(len(raw)), key=lambda k: raw[k] - base[k], reverse=True)
    for k in rest[:max(0, want - sum(base))]: base[k] += 1
    return [max(1, x) for x in base]

def apply_w(tbl, pct):
    m = re.search(r'<tr class="hdr">(.*?)</tr>', tbl, re.S)
    if not m: return None, '没有表头行'
    ths = list(re.finditer(r'<th([^>]*)>', m.group(1)))
    spans = [int((re.search(r'colspan="(\d+)"', t.group(1)) or [0, 1])[1]) for t in ths]
    if sum(spans) != len(pct): return None, '表头列数（%d）与实测列数（%d）不同' % (sum(spans), len(pct))
    if any(s > 2 for s in spans): return None, '表头有跨 3 列以上的格'
    if any(re.search(r'(?:^|\s)w-\d', (re.search(r'class="([^"]*)"', t.group(1)) or [0, ''])[1]) for t in ths): return None, '已定宽'
    row, out, pos, col = m.group(1), [], 0, 0
    for t, s in zip(ths, spans):
        out.append(row[pos:t.start()])
        attr = t.group(1)
        w = 'w-%d' % pct[col] if s == 1 else 'w-%d-%d' % (pct[col], pct[col + 1])
        col += s
        if re.search(r'class="', attr): attr = re.sub(r'class="([^"]*)"', lambda mm: 'class="%s"' % (mm.group(1) + ' ' + w).strip(), attr, count=1)
        else: attr = attr + ' class="%s"' % w
        out.append('<th%s>' % attr); pos = t.end()
    out.append(row[pos:])
    return tbl[:m.start(1)] + ''.join(out) + tbl[m.end(1):], None

def main():
    done = skip = 0
    for book, files in BOOKS.items():
        tr = os.path.join(B, 'fitfix_trace_%s.jsonl' % book)
        if not os.path.exists(tr): print('（%s 没有逐遍记录，先跑 sync.sh --full）' % book); continue
        osc, best, total = osc_tables(tr)
        print('== %s：成品取第 %s 遍，来回跳 / 未收敛的表 %d 张（版心 %d DXA）' % (book, best, len(osc), total))
        srcs = {f: open(f, encoding='utf-8').read() for f in files}
        where = {}
        for f, s in srcs.items():
            for mt in re.finditer(r'<table\b.*?</table>', s, re.S):
                where.setdefault(sig_of(mt.group(0)), []).append((f, mt.start(), mt.end()))
        edits = {}
        for sg, W in sorted(osc.items()):
            loc = where.get(sg, [])
            if len(loc) != 1:
                print('  不改  %s…：源文件里对上 %d 张' % (sg[:20], len(loc))); skip += 1; continue
            f, a, b = loc[0]
            pct = pct_of(W, total)
            new, why = apply_w(srcs[f][a:b], pct)
            if not new: print('  不改  %s：%s…（%s）' % (os.path.basename(f)[:14], sg[:20], why)); skip += 1; continue
            edits.setdefault(f, []).append((a, b, new))
            print('  定宽  %s：%s… → w-%s' % (os.path.basename(f)[:14], sg[:20], '/'.join(map(str, pct)))); done += 1
        if APPLY:
            for f, es in edits.items():
                s = srcs[f]
                for a, b, new in sorted(es, reverse=True): s = s[:a] + new + s[b:]
                open(f, 'w', encoding='utf-8').write(s)
            # 定宽表的成品宽度已含 fit_fix 的实测加宽：撤掉这些表在 keep_force 里的列宽记录（W / NW / F / NWL / WB），免得窄表再叠加一次
            kf = os.path.join(B, 'keep_force_%s.txt' % book)
            pinned = {sg for sg in osc if len(where.get(sg, [])) == 1}
            if os.path.exists(kf):
                keep, gone = [], 0
                for l in open(kf, encoding='utf-8'):
                    m2 = re.match(r'(W|NW|F):\d+:-?\d+:(.*)$|(NWL|WB):\d+:(.*)$', l.rstrip('\n'))
                    sg = m2 and (m2.group(2) if m2.group(1) else m2.group(4))
                    if sg and sg in pinned: gone += 1; continue
                    keep.append(l)
                open(kf, 'w', encoding='utf-8').writelines(keep)
                print('  keep_force：撤掉 %d 条列宽记录' % gone)
    print('合计：定宽 %d 张，不改 %d 张%s' % (done, skip, '' if APPLY else '（预览；加 --apply 写入源文件）'))

if __name__ == '__main__':
    main()
