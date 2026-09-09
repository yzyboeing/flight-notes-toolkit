#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""blockify.py —— 按「分块计划」把一节改成块编号形态（SD-18）

用法：
    python3 tools/blockify.py <plan.py>

**条目正文一个字都不动**，只做三件事：按计划重排条目、改写标题为 `#### A-1　原标题`、
生成块索引表。计划里的条目标题集合与原文不完全相等时**直接报错退出**（三道校验闸之一）。

plan.py 需定义：

    LEVEL = 4            # 原文用 ### 块 + #### 条目；若原文是并列的 ### 条目则写 3
    FILE  = "~/…/1.6 发动机、APU.md"
    PLAN  = [("A", "块名", ["条目标题1", "条目标题2"]), ...]     # 顺序即新顺序
    GROUPS= {"A": [("组名", 1, 2), ("组名", 3, 5)], ...}          # 块索引里的分组，闭区间
    SPLIT = [["A","B"], ["C"]]                                   # 块索引拆成几张表，防跨页
    NOTE  = "本节按…重排，…（2026-09-09）"                        # 节首溯源说明（不含「本节共 N 块 M 条」，自动加）
    NUMS  = '<table class="ftn">…</table>'                        # 可选，关键数字总表
    BLOCK_NOTE = {"D": "……"}                                      # 可选，挂在某块标题下的说明段

产出后**必须再跑** `diff_body.py`（正文逐行比对）与 `check_blocks.py`（形态校验）。
"""
import io, os, re, sys, importlib.util

def load_plan(p):
    spec = importlib.util.spec_from_file_location('plan', os.path.expanduser(p))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m

def main():
    if len(sys.argv) < 2: sys.exit(__doc__)
    P = load_plan(sys.argv[1])
    f = os.path.expanduser(P.FILE)
    s = io.open(f, encoding='utf-8').read()

    m = re.match(r'\A(---\n.*?\n---\n\n?#[^\n]*\n)', s, re.S)
    assert m, 'front matter / H1 未识别'
    head, body = m.group(1), s[m.end():]

    LV = getattr(P, 'LEVEL', 4)
    parts = re.split(r'(?m)^(#{3,4} .+)$', body)
    items, order = {}, []
    for i in range(1, len(parts), 2):
        h, c = parts[i].strip(), parts[i + 1]
        if (h.startswith('#### ') if LV == 4 else h.startswith('### ')):
            t = h[(5 if LV == 4 else 4):].strip()
            assert t not in items, '条目标题重复：' + t
            items[t] = c; order.append(t)
        elif c.strip():
            bns = [v.strip() for v in getattr(P, 'BLOCK_NOTE', {}).values()]
            assert c.strip() in bns, \
                '!! 块级正文会丢失，请用 BLOCK_NOTE 承接：\n%s\n%s' % (h, c.strip()[:300])

    planned = [t for _, _, ts in P.PLAN for t in ts]
    miss  = [t for t in order if t not in planned]
    extra = [t for t in planned if t not in items]
    if miss or extra:
        print('!! 计划遗漏：', miss); print('!! 计划多余：', extra); sys.exit(1)
    print('条目核对：计划 %d = 原文 %d，无遗漏无多余 OK' % (len(planned), len(order)))

    out, codes = [], {}
    for L, btitle, ts in P.PLAN:
        out.append('### %s　%s\n' % (L, btitle))
        bn = getattr(P, 'BLOCK_NOTE', {}).get(L)
        if bn: out.append(bn.strip() + '\n')
        for n, t in enumerate(ts, 1):
            code = '%s-%d' % (L, n); codes[t] = code
            out.append('#### %s　%s\n%s' % (code, t, items[t].rstrip() + '\n\n'))

    def table(letters):
        o = ['<table class="ftn">',
             '<tr class="hdr"><th>块</th><th>主题</th><th>组</th><th>条目</th></tr>']
        for L in letters:
            btitle = dict((a, b) for a, b, _ in P.PLAN)[L]
            ts = dict((a, c) for a, _, c in P.PLAN)[L]
            gs = P.GROUPS[L]
            for k, (g, lo, hi) in enumerate(gs):
                ent = '<br>'.join('<strong>%s-%d</strong> %s' % (L, i, ts[i - 1])
                                  for i in range(lo, hi + 1))
                hd = ('<td rowspan="%d"><strong>%s</strong></td>'
                      '<td rowspan="%d"><strong>%s</strong></td>'
                      % (len(gs), L, len(gs), btitle)) if k == 0 else ''
                o.append('<tr>%s<td>%s</td><td>%s</td></tr>' % (hd, g, ent))
        o.append('</table>')
        return '\n'.join(o)

    nav  = '### 块索引\n\n' + '\n\n'.join(table(g) for g in P.SPLIT) + '\n\n'
    nums = getattr(P, 'NUMS', '')
    if nums: nums = '\n### 关键数字总表\n\n' + nums.strip() + '\n\n'
    pre  = '本节共 <strong>%d 块 %d 个知识点</strong>。%s\n\n' % (len(P.PLAN), len(planned), P.NOTE)

    io.open(f, 'w', encoding='utf-8').write(head + '\n' + pre + nav + nums + '\n' + '\n'.join(out))
    print('写入完成：%s' % os.path.basename(f))
    for t in order: print('  %s  %s' % (codes[t], t))

main()
