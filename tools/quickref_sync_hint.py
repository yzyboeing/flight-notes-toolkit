#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""quickref_sync_hint.py —— 正文改了，列出要跟着改的速查条目（SD-146，2026-10-06；取代 SD-92 的「速查 → 正文」方向）
用法（在笔记库根目录）：python3 quickref_sync_hint.py [--base HEAD]
速查版独立成册（速查/速查源.md），以完整版正文为准；每条速查第一行 <!-- 详见 [[节|x.y A-n]] --> 绑定正文块。
做法：对比 notes_src 与 --base（默认 HEAD，即上次提交），找出内容改过的正文块（A-n / 第 4、5 章「N.」），
列出绑定了这些块的速查条目，提示「速查是否同步」。只读、只提示；sync.sh 在提交前调用。
忽略只删「详见」或注释的改动。"""
import sys, os, re, glob, subprocess

def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
BASE = arg('--base', 'HEAD')
QF = os.path.join('速查', '速查源.md')
if not os.path.exists(QF): sys.exit(0)

def blocks(text, sid):
    out = {}
    for m in re.finditer(r'(?ms)^#{3,4} ([A-Z]-\d+)　[^\n]*\n(.*?)(?=^#{3,4} |\Z)', text): out[m.group(1)] = m.group(2)
    if sid[0] in '45':
        for m in re.finditer(r'(?ms)^### (\d+)\. [^\n]*\n(.*?)(?=^### |\Z)', text): out[m.group(1)] = m.group(2)
    return out
norm = lambda s: re.sub(r'\s+', '', re.sub(r'[^\n]*详见[^\n]*', '', re.sub(r'<!--.*?-->', '', s, flags=re.S)))

changed = set()
for f in glob.glob('notes_src/[1-5]*/*.md'):
    sid = os.path.basename(f).split(' ')[0]
    old = subprocess.run(['git', 'show', '%s:%s' % (BASE, f)], capture_output=True, text=True).stdout
    bo, bn = blocks(old, sid), blocks(open(f, encoding='utf-8').read(), sid)
    for k in set(bo) | set(bn):
        if norm(bo.get(k, '')) != norm(bn.get(k, '')): changed.add(sid + ' ' + k)
if not changed: sys.exit(0)

q = open(QF, encoding='utf-8').read()
hits = []
for n, (t, b) in enumerate(re.findall(r'(?ms)^### ([^\n]*)\n(.*?)(?=^### |^## |\Z)', q), 1):
    labs = {re.sub(r'第\s*(\d+)\s*条', r'\1', x).strip() for x in re.findall(r'\[\[[^\]|]*\|([^\]]*)\]\]', b)}
    got = sorted(labs & changed)
    if got: hits.append((n, t, got))
if hits:
    print('\n正文本次改了 %d 个块，其中 %d 条速查绑定了它们——核对速查是否同步（SD-146）：' % (len(changed), len(hits)))
    for n, t, got in hits: print('  · 第 %d 条「%s」← %s' % (n, t[:24], '、'.join(got)))
    print()
