#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_blocks.py —— 块编号形态一致性校验（SD-18 / SD-19 / SD-20）

用法：
    python3 tools/check_blocks.py --src notes_src [--only 1.6]

在笔记主库根目录执行。这是 check_src.py（查源头）与 verify.py（查排版）之外的第三套校验，
专查「块编号形态」是否自洽——任何模型整理完都必须跑过这一关，不靠人肉复核。

检查项：
  1  块字母从 A 起连续，不跳号
  2  每块内条目从 1 起连续，不跳号
  3  条目标题在本节内唯一
  4  块索引存在（有并列条目的节），且与正文条目**逐条一一对应**（编号 + 标题、顺序一致）
  5  关键数字总表里每个「出处」引用的块编号，在本节确实存在
  6  正文无 <strong>/<em> 嵌套（渲染器会把标签当文字打出来）
  7  面向人的段落里没有漏出 Markdown / HTML 语法名（###、<code> 等）
  8  节首有溯源说明（H1 之后、第一张导航表之前的说明段）
退出码非 0 表示有错。
"""
import io, os, re, sys, glob

def _arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

SRC  = os.path.abspath(_arg('--src', 'notes_src'))
ONLY = _arg('--only')

FM   = re.compile(r'\A---\n(.*?)\n---\n', re.S)
BLK  = re.compile(r'(?m)^### ([A-Z])　(.+)$')
ITEM = re.compile(r'(?m)^#### ([A-Z])-(\d+)　(.+)$')
IDXE = re.compile(r'<strong>([A-Z]-\d+)</strong>\s*([^<]*)')
SRCE = re.compile(r'<td>((?:[A-Z]-\d+)(?:\s*[·、]\s*[A-Z]-\d+)*)</td>\s*</tr>')
NEST = re.compile(r'<(strong|em)\b[^>]*>(?:(?!</?\1\b).)*<\1\b', re.S)
LEAK = re.compile(r'(?<![`\w])(#{2,5}\s|<code>|</code>)')

def strip_tags(t):
    return re.sub(r'<[^>]+>', '', t)

def check(path):
    name = os.path.basename(path)
    t = io.open(path, encoding='utf-8').read()
    errs, warns = [], []
    m = FM.match(t)
    if not m:
        return ['%s：缺 front matter' % name], []
    body = t[m.end():]

    blocks = BLK.findall(body)
    items  = ITEM.findall(body)
    if not blocks:
        if items: errs.append('%s：有 #### 条目却没有 ### 块标题（层级缺失）' % name)
        return errs, warns                       # 单表节（SD-19），不再往下查

    # 1 块字母连续
    letters = [L for L, _ in blocks]
    want = [chr(ord('A') + k) for k in range(len(letters))]
    if letters != want:
        errs.append('%s：块字母不连续 %s，应为 %s' % (name, letters, want))

    # 2 块内条目连续
    per = {}
    for L, n, ti in items: per.setdefault(L, []).append((int(n), ti.strip()))
    for L, _ in blocks:
        ns = [n for n, _ in per.get(L, [])]
        if ns != list(range(1, len(ns) + 1)):
            errs.append('%s：%s 块条目编号不连续 %s' % (name, L, ns))

    # 3 标题唯一
    titles = [ti.strip() for _, _, ti in items]
    dup = {x for x in titles if titles.count(x) > 1}
    if dup: errs.append('%s：条目标题重复 %s' % (name, sorted(dup)))

    # 4 块索引与正文逐条对齐
    mi = re.search(r'(?s)### 块索引\n(.*?)(?=\n### |\Z)', body)
    if not mi:
        errs.append('%s：缺「块索引」' % name)
    else:
        idx = [(c, ti.strip()) for c, ti in IDXE.findall(mi.group(1))]
        real = [('%s-%s' % (L, n), ti.strip()) for L, n, ti in items]
        if idx != real:
            only_i = [x for x in idx if x not in real]
            only_r = [x for x in real if x not in idx]
            if only_i or only_r:
                errs.append('%s：块索引与正文不符　索引多出 %s　正文多出 %s'
                            % (name, only_i[:4], only_r[:4]))
            else:
                errs.append('%s：块索引与正文条目顺序不一致' % name)

    # 5 数字总表出处存在
    mn = re.search(r'(?s)### 关键数字总表\n(.*?)(?=\n### |\Z)', body)
    if mn:
        have = {'%s-%s' % (L, n) for L, n, _ in items}
        bad = sorted({c for grp in SRCE.findall(mn.group(1))
                      for c in re.findall(r'[A-Z]-\d+', grp)} - have)
        if bad: errs.append('%s：关键数字总表引用了不存在的出处 %s' % (name, bad))

    # 6 标签嵌套
    if NEST.search(body):
        errs.append('%s：有 <strong>/<em> 嵌套，渲染时标签会被当成文字打出来' % name)

    # 7 语法名泄漏（只查表格外的说明段落）
    prose = re.sub(r'(?s)<table.*?</table>', '', body)
    prose = re.sub(r'(?m)^#{1,5} .*$', '', prose)
    if LEAK.search(prose):
        errs.append('%s：正文说明里漏出 Markdown / HTML 语法名' % name)

    # 8 溯源说明
    head = body.split('### ', 1)[0]
    if not strip_tags(head).strip():
        warns.append('%s：节首没有溯源说明' % name)
    elif '（20' not in head:
        warns.append('%s：溯源说明里没有日期' % name)
    return errs, warns

def main():
    files = sorted(glob.glob(os.path.join(SRC, '*', '[0-9]*.md')))
    if ONLY:
        files = [f for f in files if os.path.basename(f).split()[0] == ONLY]
        if not files: sys.exit('没有 id 为 %s 的笔记' % ONLY)
    E, W = [], []
    for f in files:
        e, w = check(f); E += e; W += w
    print('块形态校验 %d 节' % len(files))
    for w in W: print('  提醒  ' + w)
    for e in E: print('  错误  ' + e)
    if not E: print('  全部通过' if not W else '  无错误（%d 条提醒）' % len(W))
    sys.exit(1 if E else 0)

if __name__ == '__main__':
    main()
