#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_gaps.py —— 留白页检查（2026-10-22，用户截图速查第 23 页：「下面这部分上面的空白区域放不下吗？为什么要再下一页重新开始？」）
用法：python3 check_gaps.py <成品.pdf> [--max 0]
找出：本页页底留白 ≥ 15%，下一页开头那一块（条目标题 / 要点 + 第一张表 + 表后紧跟的注 / 公司差异）按实测高度放得进本页留白——
本不该移到下一页。下一页以主题 / 章标题（≥ 11.5pt）开头的另起页属正常，不报。有问题时退出码 1（--max 允许的处数）。
2026-10-24（SD-235 排版来回跳治理 A）：量法抽成 room()，fit_fix 移页复查（拉回）直接调用——两边同一个「放得下」，
不再各算各的（1024R6 第 459 页：fit_fix 从页顶 36pt 量、要 3% 余量，差 0.03pt 没拉回；本检查从首块顶量、5% 余量，判放得下）。"""
import sys, re, pymupdf

TOP = 36
FIT = 0.95   # 放得下：需要高度 ≤ 留白 × 0.95（留 5% 实测余量）
FOOT = re.compile(r'第\s*\d+\s*页\s*$')


def _blocks(p, bot):
    return [b for b in p.get_text('blocks') if b[4].strip() and not FOOT.search(b[4].strip()) and TOP <= b[1] < bot]


def _images(p):
    return [x['bbox'] for x in p.get_image_info() if (x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 4000]


def room(d, i, min_gap=0.15):
    """第 i 页（0 起）页底留白 gap，与第 i+1 页开头那块的实测高度 need；返回 (gap, need, 表的 bbox)。
    不适用（留白不到 min_gap、下一页开头不是表、主题 / 章另起页）返回 None。放得下的判据：need <= FIT * gap"""
    bot = d[0].rect.height - 50
    bs = _blocks(d[i], bot)
    last = max([b[3] for b in bs] + [r[3] for r in _images(d[i])] + [TOP])
    gap = bot - 8 - last
    if gap < min_gap * (bot - TOP): return None
    p2 = d[i + 1]
    try: ts = [t for t in p2.find_tables().tables if t.bbox[3] - t.bbox[1] > 10]
    except Exception: ts = []
    if not ts: return None
    t = ts[0]; b2 = _blocks(p2, bot)
    first = min([b[1] for b in b2] + [t.bbox[1]])
    if t.bbox[1] - first > 90: return None                    # 下一页开头不是这张表（前面还有正文）
    big = [s for b in p2.get_text('dict')['blocks'] for l in b.get('lines', []) for s in l['spans']
           if TOP <= l['bbox'][1] < t.bbox[1] and s['size'] >= 11.5 and s['text'].strip()]
    if big: return None                                       # 主题 / 章另起页
    end = t.bbox[3]
    for r in _images(p2):
        if r[1] < t.bbox[3]: end = max(end, r[3] + 18); first = min(first, r[1])   # 图表并排：图（连图注）也算进这一块
    for b in sorted(b2, key=lambda b: b[1]):
        if b[1] >= end - 1 and b[1] - end < 14 and re.match(r'^\s*(注|公司差异)', b[4]): end = b[3]
    return gap, end - first + 6, t.bbox


def scan(d):
    """全书留白页：[(本页号, 留白, 下一页号, 需要高度, 表头文字)]，页号从 1 起"""
    out = []
    for i in range(len(d) - 1):
        r = room(d, i)
        if r and r[1] <= FIT * r[0]:
            out.append((i + 1, round(r[0]), i + 2, round(r[1]), re.sub(r'\s+', '', d[i + 1].get_text(clip=r[2]))[:24]))
    return out


def main():
    pdf = sys.argv[1]
    maxok = int(sys.argv[sys.argv.index('--max') + 1]) if '--max' in sys.argv else 0
    out = scan(pymupdf.open(pdf))
    for o in out: print('第 %d 页留白 %dpt，第 %d 页开头那块 %dpt 放得下：%s' % o)
    print('留白页检查：%d 处（允许 %d）' % (len(out), maxok))
    sys.exit(1 if len(out) > maxok else 0)


if __name__ == '__main__':
    main()
