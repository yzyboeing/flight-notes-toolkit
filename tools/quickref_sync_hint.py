#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""quickref_sync_hint.py —— 速查区改动后，列出后续章节中要一并修改的位置（SD-92，2026-09-30 用户）
用法（在笔记库根目录）：python3 quickref_sync_hint.py [--base HEAD]
用户原话：「以后只要在速查区做了更改和优化，就一定去检索后续章节中有没有相对应的内容，也要一并做更改。」
做法：
  1. 对比速查区源文件与 --base（默认 HEAD，即上次提交），找出改动过的条目（### N.）；
  2. 列出每条「详见」指向的正文条目；
  3. 取改动行里的数值和关键词，在第 1～5 章全文检索，列出所有出现位置（不限于「详见」）；
  4. 同时看这次正文有没有一起改：「详见」指向的正文文件本次没有改动的，标「！正文未改」。
只读、只提示，不改文件；sync.sh 在提交前调用。改排版标注的一致性另由 check_layout.py S5 检查。"""
import sys, os, re, glob, subprocess

def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
BASE = arg('--base', 'HEAD')
QF = 'notes_src/0 基础知识速查区/0 基础知识速查区.md'
if not os.path.exists(QF): sys.exit(0)
old = subprocess.run(['git', 'show', '%s:%s' % (BASE, QF)], capture_output=True, text=True).stdout
new = open(QF, encoding='utf-8').read()
if old == new: sys.exit(0)

def items(txt):
    d = {}
    for m in re.finditer(r'(?ms)^### (\d+)\. ([^\n]*)\n(.*?)(?=^### |^## |\Z)', txt):
        d[m.group(1)] = (m.group(2).strip(), m.group(3))
    return d
oi, ni = items(old), items(new)
norm = lambda v: re.sub(r'\s+', '', v[0] + v[1]) if v else ''   # 只看内容，空行 / 缩进变化不算
changed = [k for k in ni if norm(oi.get(k)) != norm(ni[k])] + [k for k in oi if k not in ni]
if not changed: sys.exit(0)

changed_files = set(subprocess.run(['git', '-c', 'core.quotepath=off', 'diff', '--name-only', BASE, '--', 'notes_src'],
                                   capture_output=True, text=True).stdout.split('\n'))
changed_files |= set(subprocess.run(['git', '-c', 'core.quotepath=off', 'ls-files', '--others', '--exclude-standard', 'notes_src'],
                                    capture_output=True, text=True).stdout.split('\n'))
body = {}
for f in glob.glob('notes_src/[1-5]*/*.md'):
    body[f] = open(f, encoding='utf-8').read()
plain = lambda t: re.sub(r'<[^>]+>', '', t)

print('\n速查区本次改动了 %d 条——按 SD-92 检查后续章节是否一并修改：' % len(changed))
for k in sorted(changed, key=int):
    title, txt = ni.get(k) or oi.get(k)
    otxt = (oi.get(k) or ('', ''))[1]
    # 改动的行
    ol, nl = set(otxt.split('\n')), set(txt.split('\n'))
    diff_lines = [plain(x) for x in (nl ^ ol) if plain(x).strip()]
    print('\n· 第 %s 条「%s」' % (k, title[:24]))
    links = re.findall(r'\[\[([^\]|]+)\|(\d\.\d+) ([^\]]+)\]\]', txt)
    for fname, sec, addr in links:
        f = next((x for x in body if os.path.basename(x).startswith(sec + ' ')), None)
        flag = '' if f and f in changed_files else '　！正文未改'
        print('    详见 %s %s%s' % (sec, addr, flag))
    # 关键词：改动行里的数值（带单位）和 4 字以上的中文词组
    keys = set()
    for ln in diff_lines:
        keys |= set(re.findall(r'\d+(?:\.\d+)?\s?(?:kt|ft|fpm|psi|nm|NM|kg|m|min|s|%|℃|°|g|mb|hPa)', ln))
        keys |= set(w for w in re.findall(r'[一-鿿]{4,8}', ln) if w not in ('详见正文', '以下情况'))
    hits = {}
    for key in sorted(keys, key=len, reverse=True)[:12]:
        where = [f for f, s in body.items() if key in plain(s)]
        if len(where) > 4: continue          # 太泛的词（如「公司差异」）不列，免得淹没真正相关的位置
        for f in where:
            hits.setdefault(os.path.basename(f)[:-3], set()).add(key)
    if hits:
        print('    后续章节中出现改动关键词的位置（逐一核对要不要一起改）：')
        for f, ks in sorted(hits.items()):
            print('      %s：%s' % (f, '、'.join(sorted(ks))[:80]))
print()
