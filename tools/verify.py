#!/usr/bin/env python3
"""verify.py —— 交付前必跑的校验（见 prompt/SKILL.md 第八节）

用法：python3 tools/verify.py build/xxx.pdf
"""
import sys, re
import pymupdf

d = pymupdf.open(sys.argv[1])
blank, leak, bullet, fm = [], [], [], []
for i in range(d.page_count):
    t = d[i].get_text()
    body = re.sub(r'第 \d+ 页', '', t).strip()
    if not body and i > 0: blank.append(i + 1)
    if re.search(r'</?(em|strong|td|th|tr)>', t): leak.append(i + 1)
    if re.search(r'[▪•]', t): bullet.append(i + 1)
    if re.search(r'title_cn:|aircraft:', t): fm.append(i + 1)
print('页数', d.page_count)
print('空白页      ', blank)
print('标签泄漏    ', leak)
print('项目符号残留', bullet)
print('front matter', fm)
sys.exit(1 if (blank or leak or bullet or fm) else 0)
