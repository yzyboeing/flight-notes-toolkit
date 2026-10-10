#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""preview.py —— 单节预览：只重建改动的那几节，1～3 分钟出 PDF 和排版检查（2026-10-07 用户：「为什么每次都这么久」）
用法（在笔记库根目录）：
    python3 preview.py 1.4 2.7          只排这几节（节号）
    python3 preview.py 1.4 --open       排完打开 PDF
    python3 preview.py 1.4 --no-merge   不把列宽记录写回全书
做法：
  1. assemble.py + make_book.py 生成 build/book.md（几秒）；从中截出指定的节，拼成 build/preview/预览.md（带章标题，不出封面目录、不出按主题查）。
  2. 复制全书的 keep_force 记录作起点，用同一套 fit_fix.py（最多 4 遍）+ build_docx.js 排版，出 build/preview/预览.pdf。
  3. check_layout.py --section 只查这几节的表格与版面，报告写 build/preview/预览检查报告.md，错误逐条打印。
  4. 列宽 / 字距记录（W / WB / NW / NWL / F / C，与分页位置无关）按表签名写回全书 keep_force，最后一次全量重建直接用上，不必从头收敛。
     整表同页、压缩级（与分页位置有关）不写回，全量时重新判断。
单节里改到 0 后，再跑一次 sync.sh --full 出版。"""
import sys, os, re, subprocess, shutil, json

T = os.path.dirname(os.path.abspath(__file__))
ROOT = os.getcwd()
secs = [a for a in sys.argv[1:] if re.fullmatch(r'\d\.\d{1,2}', a)]
if not secs: print(__doc__); sys.exit(1)
B = os.path.join(ROOT, 'build'); P = os.path.join(B, 'preview'); os.makedirs(P, exist_ok=True)
git = lambda *a: subprocess.run(['git'] + list(a), capture_output=True, text=True, cwd=ROOT).stdout.strip()

# 1. 组装并截出指定节
subprocess.run([sys.executable, os.path.join(T, 'assemble.py'), '--src', 'notes_src', '--out', 'build'], cwd=ROOT, check=True, capture_output=True)
subprocess.run([sys.executable, os.path.join(T, 'make_book.py')], cwd=ROOT, check=True, capture_output=True)
book = open(os.path.join(B, 'book.md'), encoding='utf-8').read()
heads = [(m.start(), m.group(1), m.group(2)) for m in re.finditer(r'(?m)^(##|###) ([^\n]*)$', book)]
out, last_ch = ['# 预览\n'], None
for k, (pos, lvl, title) in enumerate(heads):
    if lvl == '##': ch = (pos, title); continue
    sid = title.split('　')[0]
    if sid not in secs: continue
    end = next((p for p, l, _ in heads[k + 1:]), len(book))
    if last_ch != ch[1]: out.append('\n## %s\n' % ch[1]); last_ch = ch[1]
    out.append(book[pos:end])
missing = [s for s in secs if not any(s == t.split('　')[0] for _, l, t in heads if l == '###')]
if missing: sys.exit('找不到节：' + '、'.join(missing))
MD = os.path.join(P, '预览.md'); open(MD, 'w', encoding='utf-8').write(''.join(out))

# 2. 排版：起点用全书的记录
BOOKKF = os.path.join(B, 'keep_force_B737机型理论知识笔记.txt')
DOCX = os.path.join(P, '预览.docx'); KF = os.path.join(P, 'keep_force_预览.txt')
if os.path.exists(BOOKKF): shutil.copy2(BOOKKF, KF)
env = dict(os.environ, DOC_FIGS=os.environ.get('DOC_FIGS', '3.2-1_,2.1-1_,2.1-2_'), NODE_PATH=os.path.join(ROOT, 'node_modules'), FITFIX_MAXPASS='4',
           DOC_EDITION=git('config', '--get', 'notes.docEdition') or 'preview', DOC_HEADER='B737-NG / B737-8 机型理论知识笔记',
           DOC_AUTHOR=git('config', '--get', 'notes.docAuthor'), DOC_TOPICS='', DOC_PREFACE='/nonexistent')
r = subprocess.run([sys.executable, os.path.join(T, 'fit_fix.py'), MD, DOCX], env=env, cwd=ROOT, capture_output=True, text=True)
print('\n'.join(l for l in r.stdout.splitlines() if l.startswith('fit_fix')))
if r.returncode: sys.exit('排版失败：' + (r.stderr or r.stdout)[-800:])
PDF = os.path.join(P, '预览.pdf')
if not os.path.exists(PDF): sys.exit('没有生成 预览.pdf')

# 3. 检查（只查这几节）
REP = os.path.join(P, '预览检查报告.md')
subprocess.run([sys.executable, os.path.join(T, 'check_layout.py'), '--repo', ROOT, '--book', PDF, '--md', MD, '--section', '--no-src', '--out', REP],
               cwd=ROOT, capture_output=True, text=True)
rep = open(REP, encoding='utf-8').read() if os.path.exists(REP) else ''
errs = re.findall(r'(?m)^- \[([A-Z]\d+)\] (.*)$', rep.split('## 新建议')[0]) if '## 错误' in rep else []
errs = [(c, m) for c, m in errs if c not in ('F1', 'F3', 'P1', 'P2', 'R5', 'C1', 'V1')]   # 整书类检查不在预览范围
import pymupdf
print('预览：%s，%d 页 → %s' % ('、'.join(secs), len(pymupdf.open(PDF)), os.path.relpath(PDF, ROOT)))
print('排版检查：错误 %d 条' % len(errs) + ('' if not errs else '（必须改）'))
for c, m in errs: print('  [%s] %s' % (c, m[:160]))

# 4. 列宽记录写回全书
if '--no-merge' not in sys.argv and os.path.exists(BOOKKF):
    WIDTH = ('W:', 'WB:', 'NW:', 'NWL:', 'F:', 'C:')
    sig_of = lambda l: (l.split(':', 3)[-1] if l.startswith(('W:', 'NW:', 'F:')) else l.split(':', 2)[-1].rsplit('|', 1)[0] if l.startswith('C:') else l.split(':', 2)[-1])
    new = [l.rstrip('\n') for l in open(KF, encoding='utf-8') if l.startswith(WIDTH)]
    sigs = {sig_of(l) for l in new}
    old = [l.rstrip('\n') for l in open(BOOKKF, encoding='utf-8')]
    keep = [l for l in old if not (l.startswith(WIDTH) and sig_of(l) in sigs)]
    open(BOOKKF, 'w', encoding='utf-8').write('\n'.join(keep + new) + '\n')
    print('列宽记录已写回全书：%d 张表' % len(sigs))
if '--open' in sys.argv: subprocess.run(['open', PDF])
sys.exit(1 if errs else 0)
