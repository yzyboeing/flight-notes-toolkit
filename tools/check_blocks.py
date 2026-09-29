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
  4  块索引存在（有并列条目的节；第零章不设，SD-23），且与正文条目**逐条一一对应**（编号 + 标题、顺序一致）
  5  关键数字总表里每个「出处」引用的块编号，在本节确实存在
  6  正文无 <strong>/<em> 嵌套（渲染器会把标签当文字打出来）
  7  面向人的段落里没有漏出 Markdown / HTML 语法名（###、<code> 等）
  8  节首不放整理说明（SD-30：「本节共…」、带日期的改动记录移到主库「整理记录」）；第 4、5 章不分块、条目「N. 标题」连续编号
退出码非 0 表示有错。
"""
import io, os, re, sys, glob

def _arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

SRC  = os.path.abspath(_arg('--src', 'notes_src'))
ONLY = _arg('--only')

FM   = re.compile(r'\A---\n(.*?)\n---\n', re.S)
BLK  = re.compile(r'(?m)^### ([A-Z])　(.+)$')
ITEM = re.compile(r'(?m)^#{3,4} ([A-Z])-(\d+)　(.+)$')   # SD-43：条目直接用 ### A-1（不再有块标题）
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

    if '.' not in name.split()[0]:               # 扁平节（第零章速查区，SD-23）
        return check_flat(name, body)

    if name.split('.')[0] in ('4', '5'):         # 第 4、5 章不分块（SD-30）
        return check_numbered(name, body)

    blocks = BLK.findall(body)
    items  = ITEM.findall(body)
    if blocks: errs.append('%s：正文里不再写「### A　块主题」块标题，块主题只在块索引里（SD-43）' % name)
    if re.search(r'(?m)^#### [A-Z]-\d+　', body): errs.append('%s：条目标题应为「### A-1　…」（SD-43）' % name)
    if not items:
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

    # 4 块索引与正文逐条对齐（第零章速查区不设块索引，SD-23）
    mi = re.search(r'(?s)### 块索引\n(.*?)(?=\n### |\Z)', body)
    if name.startswith('0.'):
        if mi: errs.append('%s：第零章不设块索引（SD-23）' % name)
    elif not mi:
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

    # 8 节首不放整理说明（SD-30）
    head = body.split('### ', 1)[0]
    if re.search(r'本节共\s*<strong>', head) or re.search(r'(?m)^<strong>20\d\d-\d\d-\d\d', head):
        errs.append('%s：节首还有整理说明（SD-30：移到主库「整理记录」）' % name)

    # 10 id / 文件名前缀 / H1 三处同号（SD-21）
    fid = re.search(r'^id:\s*"?([0-9.]+)"?\s*$', m.group(0), re.M)
    pre = name.split()[0]
    if fid and fid.group(1) != pre:
        errs.append('%s：front matter id 为 %s，与文件名前缀 %s 不符' % (name, fid.group(1), pre))
    h1 = re.search(r'(?m)^#\s+([0-9.]+)\s*　', body)
    if h1 and h1.group(1) != pre:
        errs.append('%s：H1 编号为 %s，与文件名前缀 %s 不符' % (name, h1.group(1), pre))

    return errs, warns

def check_numbered(name, body):
    """SD-30：第 4、5 章——不设块索引与块标题，条目「### N. 标题」每节从 1 连续编号"""
    errs, warns = [], []
    if re.search(r'(?m)^### 块索引', body): errs.append('%s：第 4、5 章不设块索引（SD-30）' % name)
    if re.search(r'(?m)^#### |^### [A-Z]　', body): errs.append('%s：第 4、5 章不应有块标题或 #### 条目（SD-30）' % name)
    head = body.split('### ', 1)[0]
    if re.search(r'本节共\s*<strong>', head) or re.search(r'(?m)^<strong>20\d\d-\d\d-\d\d', head):
        errs.append('%s：节首还有整理说明（SD-30）' % name)
    nums, titles = [], []
    for tx in re.findall(r'(?m)^### (.+)$', body):
        m = re.match(r'(\d+)\. (.+)$', tx)
        if not m: errs.append('%s：条目标题不是「N. 标题」格式：%s' % (name, tx[:30])); continue
        nums.append(int(m.group(1))); titles.append(m.group(2).strip())
    if nums != list(range(1, len(nums) + 1)):
        errs.append('%s：条目编号不连续 %s' % (name, nums))
    dup = {x for x in titles if titles.count(x) > 1}
    if dup: errs.append('%s：条目标题重复 %s' % (name, sorted(dup)))
    if NEST.search(body): errs.append('%s：<strong>/<em> 嵌套' % name)
    stale = re.findall(r'(?<![0-9A-Za-z.\-])[A-H]-\d{1,2}(?!\d)', re.sub(r'\d+\.\d+\s*[A-H]-\d{1,2}', '', re.sub(r'(?m)^#.*$', '', strip_tags(body))))   # 「x.y A-n」是跨章地址（第 1–3 章），合法   # 4 位数的 B-xxxx 是飞机注册号（SD-32），不算块编号
    if stale: warns.append('%s：正文里还有块编号式引用 %s（第 4、5 章已改为「第 N 条」）' % (name, sorted(set(stale))[:5]))
    return errs, warns

def check_flat(name, body):
    """SD-23：第零章一页连续速查表——## 块标题 + ### 全章连续编号「N. 标题」"""
    errs, warns = [], []
    if re.search(r'(?m)^### 块索引', body): errs.append('%s：第零章不设块索引（SD-23）' % name)
    if re.search(r'(?m)^#### ', body): errs.append('%s：扁平节不应出现 #### 标题' % name)
    heads = re.findall(r'(?m)^(##|###) (.+)$', body)
    if not heads or heads[0][0] != '##':
        errs.append('%s：第一个标题应是 ## 块标题' % name)
    nums, titles = [], []
    for lv, tx in heads:
        if lv != '###': continue
        m = re.match(r'(\d+)\. (.+)$', tx)
        if not m: errs.append('%s：条目标题不是「N. 标题」格式：%s' % (name, tx[:30])); continue
        nums.append(int(m.group(1))); titles.append(m.group(2).strip())
    if nums != list(range(1, len(nums) + 1)):
        bad = [n for k, n in enumerate(nums, 1) if n != k][:5]
        errs.append('%s：条目编号不连续（应为 1–%d），首个异常 %s' % (name, len(nums), bad))
    dup = {x for x in titles if titles.count(x) > 1}
    if dup: warns.append('%s：条目标题重复 %s' % (name, sorted(dup)))
    if NEST.search(body): errs.append('%s：<strong>/<em> 嵌套' % name)
    return errs, warns

def check_numbering(files):
    """SD-21：节编号在章内连续不跳号；MANIFEST 与 000 总目录同号"""
    errs, warns = [], []
    bych = {}
    for f in files:
        pre = os.path.basename(f).split()[0]
        if '.' not in pre: continue
        ch, no = pre.split('.', 1)
        if not no.isdigit(): continue
        bych.setdefault(ch, []).append((int(no), pre))
    for ch in sorted(bych, key=lambda x: int(x) if x.isdigit() else 99):
        nums = sorted(n for n, _ in bych[ch])
        start = 0 if 0 in nums else 1            # x.0 是本章导读，允许从 0 起
        want = list(range(start, start + len(nums)))
        if nums != want:
            miss = [n for n in want if n not in nums]
            errs.append('第 %s 章节编号不连续：实际 %s，缺 %s（SD-21 要求删除后顺延，不留空号）'
                        % (ch, nums, miss))
    # MANIFEST
    mf = os.path.join(SRC, 'MANIFEST.txt')
    if os.path.exists(mf):
        ids = set()
        for ln in io.open(mf, encoding='utf-8'):
            mm = re.match(r'\s*([0-9]+(?:\.[0-9]+)?)\s', ln)
            if mm: ids.add(mm.group(1))
        real = set(os.path.basename(f).split()[0] for f in files)
        if ids != real:
            only_mf = sorted(ids - real); only_rl = sorted(real - ids)
            errs.append('MANIFEST 与实际笔记不符：表内多出 %s，表内缺 %s' % (only_mf or '无', only_rl or '无'))
    # 000 总目录
    toc = os.path.join(SRC, '000 总目录.md')
    if not os.path.exists(toc):   # 2026-09-29 起总目录移到仓库根目录「整理规范/全库总目录.md」
        toc = os.path.join(os.path.dirname(SRC), '整理规范', '全库总目录.md')
    if os.path.exists(toc):
        t = io.open(toc, encoding='utf-8').read()
        listed = set(re.findall(r'<tr><td><strong>([0-9]+(?:\.[0-9]+)?)</strong>', t))
        real = set(os.path.basename(f).split()[0] for f in files)
        if listed and listed != real:
            errs.append('000 总目录与实际笔记不符：目录多出 %s，目录缺 %s'
                        % (sorted(listed - real) or '无', sorted(real - listed) or '无'))
    return errs, warns

def main():
    files = sorted(glob.glob(os.path.join(SRC, '*', '[0-9]*.md')))
    if ONLY:
        files = [f for f in files if os.path.basename(f).split()[0] == ONLY]
        if not files: sys.exit('没有 id 为 %s 的笔记' % ONLY)
    E, W = [], []
    for f in files:
        e, w = check(f); E += e; W += w
    if not ONLY:                                  # 编号连续性与清单对账是全库级的
        e, w = check_numbering(files); E += e; W += w
    print('块形态校验 %d 节' % len(files))
    for w in W: print('  提醒  ' + w)
    for e in E: print('  错误  ' + e)
    if not E: print('  全部通过' if not W else '  无错误（%d 条提醒）' % len(W))
    sys.exit(1 if E else 0)

if __name__ == '__main__':
    main()
