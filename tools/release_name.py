#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""release_name.py —— 版本名与基线标签互换（SD-142，2026-10-05 用户）

用户原话：「以后新的文件名改为：B737 机型理论知识笔记1005R1。后面的 1005 是 10 月 5 号……R1 是 10 月 5 号的第一版，
如果是 10 月 5 号做的第二版就是 R2。以此类推。」

  版本名  MMDDRn      例：1005R1、1005R7、1006R1
  git 标签 baseline/YYYYMMDD（当天第 1 版）、baseline/YYYYMMDD-n（当天第 n 版，n ≥ 2）——标签保留年份，避免跨年重名
  成品文件 B737机型理论知识笔记<版本名>.docx / .pdf

用法：
  python3 release_name.py                 # 最新 baseline/* 标签对应的版本名
  python3 release_name.py --next          # 今天再出一版时的版本名和标签（出基线前用）
  python3 release_name.py baseline/20261005-7   # → 1005R7
  python3 release_name.py 1005R7          # → baseline/20261005-7（年份取今年）
"""
import sys, re, subprocess, datetime, os

REPO = os.path.expanduser('~/flight-repos/gh-private')
STEM = 'B737机型理论知识笔记'

def tags():
    out = subprocess.run(['git', '-C', REPO, 'tag', '--list', 'baseline/*'], capture_output=True, text=True).stdout.split()
    return [t for t in out if re.fullmatch(r'baseline/\d{8}(-\d+)?', t)]

def to_name(tag):
    m = re.fullmatch(r'(?:baseline/)?(\d{4})(\d{4})(?:-(\d+))?', tag)
    if not m: raise SystemExit('不是基线标签：' + tag)
    return '%sR%d' % (m.group(2), int(m.group(3) or 1))

def to_tag(name, year=None):
    m = re.fullmatch(r'(\d{4})R(\d+)', name)
    if not m: raise SystemExit('不是版本名：' + name)
    y = year or datetime.date.today().year
    n = int(m.group(2))
    return 'baseline/%d%s%s' % (y, m.group(1), '' if n == 1 else '-%d' % n)

def key(t):
    m = re.fullmatch(r'baseline/(\d{8})(?:-(\d+))?', t)
    return (m.group(1), int(m.group(2) or 1))

if __name__ == '__main__':
    a = sys.argv[1:]
    if not a:
        t = max(tags(), key=key); print(to_name(t), t, STEM + to_name(t))
    elif a[0] == '--next':
        today = datetime.date.today().strftime('%Y%m%d')
        n = max([key(t)[1] for t in tags() if key(t)[0] == today], default=0) + 1
        t = 'baseline/%s%s' % (today, '' if n == 1 else '-%d' % n)
        print(to_name(t), t, STEM + to_name(t))
    elif a[0].startswith('baseline/') or re.fullmatch(r'\d{8}(-\d+)?', a[0]):
        print(to_name(a[0]))
    else:
        print(to_tag(a[0]))
