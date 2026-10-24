#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_overshrink.py —— 过度压缩检查（2026-10-22，SD-229 / 用户：「若表格放不下，可进行压缩，但不要过度压缩」）
用法：python3 check_overshrink.py <成品.pdf> [--max 0] [--src 源文件 ...]（放最后）
找出字号小于正文 9pt 的表（压缩过的表）：
  ① 退一级也放得下本页＝压过头（K-Z3）；
  ② 按字号折回 9pt 估整表高：9pt 一页（版心）放得下的却被压小了。
有问题时退出码 1。

2026-10-23（SD-232，用户：「过度压缩为什么会反复出现」）：
  · 「压过头」的定义只有一处——本文件的 scan_over()。fit_fix.py 定稿退级 import 同一个函数选候选，两边不再各算各的；
  · 候选只是估算，结论以实测为准：fit_fix 把每张候选真退一级重排，结果写进 build/压缩退级记录_<书名>.json；
    记录里实测退不下来的（注掉页 / 拆表 / 页数增加）只出「提示」，没试过、或匹配不到压缩记录的仍报错误。"""
import sys, os, re, json, collections
import pymupdf

SHRINK_F = [1.0, 0.90, 0.85, 0.77, 0.68, 0.57, 0.50]   # 同 fit_fix.py：0 级 9pt → 6 级 7pt，各级表高约为 9pt 时的比例
SIZE_LV = {8.5: 2, 7.5: 5, 7.0: 6}   # 读不到压缩记录时按字号推级（8pt 有 3、4 两级，按行距分）
PITCH_8PT = 8.0 * 1.8   # 8pt 两级的行距分界：实测 3 级（行距 240）15.6～15.8pt、4 级（行距 205）13.0～13.1pt，取中点约 14.4pt


def _norm(s):
    return re.sub(r'[\W_]+', '', s)


def read_levels(kf):
    """keep_force 里的压缩级：签名 → 级（~N 与 Z~N 取较重的一级）"""
    lv = {}
    if kf and os.path.exists(kf):
        for l in open(kf, encoding='utf-8'):
            m = re.match(r'Z?~(\d):(.+)$', l.strip())
            if m: lv[m.group(2)] = max(lv.get(m.group(2), 0), int(m.group(1)))
    return lv


def small_heads(src):
    """点名缩一档（small）、源文件定级（fit-N）的表不算过度压缩：返回这些表表头文字的前 8 字"""
    out = []
    for f in src:
        for m in re.finditer(r'<table class="[^"]*\b(?:small|fit-[1-6])\b[^"]*">\s*(?:<tr class="premise">.*?</tr>\s*)?<tr class="hdr">(.*?)</tr>', open(f, encoding='utf-8').read(), re.S):
            out.append(re.sub(r'<[^>]+>|\s+', '', m.group(1))[:8])
    return out


def _stem(pdf):
    return re.sub(r'\.fitfix$', '', os.path.splitext(os.path.basename(pdf))[0])


def log_path(pdf):
    return os.path.join(os.path.dirname(os.path.abspath(pdf)), '压缩退级记录_%s.json' % _stem(pdf))


def scan_over(pdf, src=(), levels=None):
    """逐页找压缩过的表，返回两类结果：
         over：退一级估计仍放得下本页的表（候选，等实测）——[{page, y0, y1, head, sig（匹配到的压缩记录签名，可能为空）, level, size, table_h, extra, free, note}]
         nine：折回 9pt 一页也放得下的表——[(page, size, h, 占页高 %, 表首)]
       levels：签名 → 压缩级；不传就读 PDF 旁的 keep_force_<书名>.txt。
       表高只算压缩字号的部分：PDF 认出的表框会把紧跟表后的 9pt 注一并框进去（1023R8：估多 40pt、实测多约 70pt 的原因之一）。"""
    if levels is None:
        levels = read_levels(os.path.join(os.path.dirname(os.path.abspath(pdf)), 'keep_force_%s.txt' % _stem(pdf)))
    sigs = [(s[:24], v, s) for s, v in levels.items() if v >= 2]
    small = small_heads(src)
    d = pymupdf.open(pdf); H = d[0].rect.height; AREA = (H - 50) - 36
    over, nine = [], []
    for pn, p in enumerate(d, 1):
        try: ts = [t for t in p.find_tables().tables if t.bbox[3] - t.bbox[1] > 30]
        except Exception: continue
        if not ts: continue
        lines = [(l['spans'][0]['size'], l['bbox'], ''.join(x['text'] for x in l['spans'])) for b in p.get_text('dict')['blocks'] for l in b.get('lines', []) if l['spans'] and ''.join(x['text'] for x in l['spans']).strip()]
        foot = min([bb[1] for z, bb, tt in lines if bb[1] > H * 0.9] or [H - 40])
        bottom = max([bb[3] for z, bb, tt in lines if bb[3] < foot - 2] + [0])
        for t in ts:
            inn = [(round(z * 2) / 2, bb, tt) for z, bb, tt in lines if t.bbox[1] - 1 <= bb[1] <= t.bbox[3] and t.bbox[0] - 1 <= bb[0] <= t.bbox[2]]
            sz = collections.Counter(z for z, bb, tt in inn if z >= 6)
            if not sz: continue
            body = sz.most_common(1)[0][0]
            if body >= 9: continue
            tx = re.sub(r'\s+', '', p.get_text(clip=t.bbox))
            if any(tx.startswith(k) for k in small): continue
            nt = _norm(tx)
            hit = [(v, full) for s, v, full in sigs if nt.startswith(s) or (len(nt) >= 16 and s.startswith(nt[:24]))]
            lvl = max([v for v, full in hit] or [0])
            sig_hit = [full for v, full in hit if v == lvl]
            if not lvl:
                if body == 8.0:
                    ys = sorted(set(round(bb[1], 1) for z, bb, tt in inn if z == body))
                    gs = sorted(g for g in (b - a for a, b in zip(ys, ys[1:])) if 6 < g < 20)
                    lvl = 4 if gs and gs[len(gs) // 2] < PITCH_8PT else 3
                else:
                    lvl = SIZE_LV.get(body, 0)
            if not lvl: continue
            tb = max(bb[3] for z, bb, tt in inn if z < 9)
            table_h = tb - t.bbox[1] + 3   # 末行文字底到表框底的边距约 3pt
            nts = [tt for z, bb, tt in sorted(inn + [(round(z * 2) / 2, bb, tt) for z, bb, tt in lines if t.bbox[3] - 2 <= bb[1] <= t.bbox[3] + 30], key=lambda x: x[1][1]) if bb[1] > tb and z >= 9 and tt.strip().startswith('注')]
            item = {'page': pn, 'y0': round(t.bbox[1]), 'y1': round(t.bbox[3]), 'head': tx[:12], 'sig': sig_hit, 'level': lvl, 'size': body,
                    'table_h': round(table_h), 'free': round(foot - bottom), 'note': re.sub(r'\s+', '', nts[0])[:12] if nts else ''}
            if t.bbox[3] < H * 0.92:
                extra = table_h * (SHRINK_F[lvl - 1] / SHRINK_F[lvl] - 1)
                if bottom + extra <= foot - 14:
                    item['extra'] = round(extra); over.append(item); continue
            h9 = table_h / SHRINK_F[lvl]
            if h9 < 0.8 * AREA:
                nine.append((pn, body, round(t.bbox[3] - t.bbox[1]), round(100 * h9 / AREA), tx[:22]))
    return over, nine


def read_log(pdf):
    try: return json.load(open(log_path(pdf), encoding='utf-8'))
    except Exception: return []


if __name__ == '__main__':
    PDF = sys.argv[1]
    MAXOK = int(sys.argv[sys.argv.index('--max') + 1]) if '--max' in sys.argv else 0
    SRC = sys.argv[sys.argv.index('--src') + 1:] if '--src' in sys.argv else []
    over, nine = scan_over(PDF, SRC)
    log = read_log(PDF)
    errs, hints = [], []
    for c in over:
        e = next((e for e in log if e.get('head', '')[:12] == c['head'] and e.get('from') == c['level']), None)
        msg = '第 %d 页：表压到 %.1fpt（%d 级，表高 %dpt），估计退一级只多约 %dpt、页底还空 %dpt' % (c['page'], c['size'], c['level'], c['table_h'], c['extra'], c['free'])
        if e and e.get('result') == 'fail':
            hints.append('提示 「%s」%d→%d 级已实测退不下来：%s，保持原级（第 %d 页）' % (c['head'], c['level'], c['level'] - 1, e.get('reason', ''), c['page']))   # 文字不带估算 pt，页码放末尾括号：check_layout 去掉它，建议编号每次重建不变
        else:
            if e and e.get('result') == 'nomatch':   # 多半是跨页表的下半截（续页重复表头，开头对不上压缩记录）——先看是不是压缩后又被拆开（check_splits 同报）
                errs.append('%s——本页这张表对不上压缩记录，多半是压缩过的表被拆到两页（应整表同页，或 7pt 仍放不下时恢复 9pt 续排）：%s' % (msg, c['head']))
            else:
                errs.append('%s——压过头了，应退到轻一级（%s）：%s' % (msg, '没有试过退级' if not e else '退过一级仍压过头', c['head']))
    for pn, body, h, pct, head in nine:
        errs.append('第 %d 页：表压到 %.1fpt（高 %dpt），折回 9pt 约 %d%% 页高——9pt 本可放下一页：%s' % (pn, body, h, pct, head))
    for l in errs + hints: print(l)
    print('过度压缩检查：错误 %d 处（允许 %d）、实测退不下来的提示 %d 处' % (len(errs), MAXOK, len(hints)))
    sys.exit(1 if len(errs) > MAXOK else 0)
