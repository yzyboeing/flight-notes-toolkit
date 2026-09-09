#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""diff_body.py —— 改结构前后的正文逐行比对（三道校验闸之二）

用法：
    python3 tools/diff_body.py <改前.md> <改后.md>

重排 / 分块 / 合并之后，**除新增的导航块外，原文每一行正文必须一字不差地保留**。
本工具做行级集合比对（Counter 差集），不是肉眼通读——肉眼在几百行表格上必然漏。

比对时会：
  · 去掉 front matter 与所有标题行（标题本来就要改）
  · 去掉改后文件里 H1 到第一个「### 字母　块名」之间的导航区（溯源说明 / 块索引 / 数字总表）
  · 其余每一行去首尾空白后按出现次数比对

丢失 = 原文有、现文没有；新增 = 现文有、原文没有。
两者都为空才算通过。被有意取代的行（如旧的节首编排说明）会出现在「丢失」里，
**必须逐条确认其内容确已并入新的溯源说明**，不能默认忽略。
"""
import io, os, re, sys
from collections import Counter

def lines(path, strip_nav=False):
    t = io.open(path, encoding='utf-8').read()
    if strip_nav:
        t = re.sub(r'(?s)^.*?\n(### [A-Z]　)', r'\1', t, count=1)
    t = re.sub(r'(?s)\A---\n.*?\n---\n', '', t)
    t = re.sub(r'(?m)^#{1,5} .*$', '', t)
    return [l.strip() for l in t.split('\n') if l.strip()]

def main():
    if len(sys.argv) < 3: sys.exit(__doc__)
    old, new = sys.argv[1], sys.argv[2]
    a, b = lines(old), lines(new, True)
    ca, cb = Counter(a), Counter(b)
    miss  = list((ca - cb).elements())
    extra = list((cb - ca).elements())
    print('%s　原文正文 %d 行 / 现文 %d 行' % (os.path.basename(new), len(a), len(b)))
    if miss:
        print('  !! 丢失 %d 行：' % len(miss))
        for x in miss[:20]: print('    -', x[:140])
    if extra:
        print('  !! 新增 %d 行：' % len(extra))
        for x in extra[:20]: print('    +', x[:140])
    if not miss and not extra: print('  正文逐行一致 OK')
    sys.exit(1 if (miss or extra) else 0)

main()
