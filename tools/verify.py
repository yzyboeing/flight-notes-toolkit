#!/usr/bin/env python3
"""verify.py —— 交付前必跑的校验（见 prompt/SKILL.md 第八节）

用法：python3 tools/verify.py build/xxx.pdf
"""
import sys, re
try:
    import pymupdf
except ImportError:  # PyMuPDF 旧版模块名
    import fitz as pymupdf

d = pymupdf.open(sys.argv[1])
blank, leak, bullet, fm = [], [], [], []
for i in range(d.page_count):
    t = d[i].get_text()
    body = re.sub(r'第 \d+ 页', '', t).strip()
    if not body and i > 0: blank.append(i + 1)
    if re.search(r'</?(em|strong|td|th|tr)>|\[\[|\]\]', t): leak.append(i + 1)      # 含未还原的 [[双链]]
    # SD-75 允许渲染器生成的「▪ 主题 / – 子项」纯文本层级；其余项目符号仍视为残留。
    # LibreOffice / PyMuPDF 有时会吞掉符号后的空格；行首的 ▪ 仍是渲染器层级标记。
    t_no_hierarchy = re.sub(r'(?m)^\s*[▪•]\s*', '', t)   # SD-78：渲染器为并列句加的「•」同样放行
    if re.search(r'[▪•]', t_no_hierarchy): bullet.append(i + 1)
    if re.search(r'title_cn:|aircraft:', t): fm.append(i + 1)
print('页数', d.page_count)
print('空白页      ', blank)
print('标签泄漏    ', leak)
print('异常项目符号  ', bullet)
print('front matter', fm)
sys.exit(1 if (blank or leak or bullet or fm) else 0)
