#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_book.py —— 拼出「全书合订本」的源文件 build/book.md

用法（在笔记主库根目录，assemble.py 跑完之后）：
    python3 tools/make_book.py [--src notes_src] [--out build]

顺序：封面 → 自动目录 → **全库总目录** → **笔记编排规范** → 六章正文。

前两块不是新写的内容，直接取自库里已有的两份文件，避免出现第二份会过期的副本：
    notes_src/000 总目录.md      （id: MOC-ROOT，assemble.py 不当它是节）
    notes_src/_编排规范.md        （id: SPEC，同上）
两份文件整体降一级（## → ###）后挂在各自的分册标题下。

`build/full.md` 只有六章正文；**全书 docx 必须用 book.md 渲染，不要用 full.md**，
否则总目录与编排规范会被静默丢掉。
"""
import io, os, re, sys

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

SRC = os.path.abspath(_arg('--src', 'notes_src'))
OUT = os.path.abspath(_arg('--out', 'build'))

# (文件名, 在书里用的分册名；None = 用文件自己的 H1)。顺序即成书顺序。
FRONT = [('000 总目录.md', '全库总目录'),
         ('_编排规范.md',  None)]

def prep(path):
    """去 front matter，取 H1 当分册名，其余标题整体降一级"""
    t = io.open(path, encoding='utf-8').read()
    t = re.sub(r'(?s)\A---\n.*?\n---\n', '', t)
    m = re.match(r'\s*#\s+([^\n]*)\n', t)
    if not m:
        raise SystemExit('%s 没有 H1 标题' % os.path.basename(path))
    title, t = m.group(1).strip(), t[m.end():]
    t = re.sub(r'(?m)^####\s', '##### ', t)
    t = re.sub(r'(?m)^###\s',  '#### ',  t)
    t = re.sub(r'(?m)^##\s',   '### ',   t)
    return title, t.strip()

def main():
    full_p = os.path.join(OUT, 'full.md')
    if not os.path.exists(full_p):
        raise SystemExit('缺 %s —— 先跑 assemble.py' % full_p)
    full = io.open(full_p, encoding='utf-8').read()
    i = full.index('%%PAGEBREAK%%')

    parts, names = [full[:i].rstrip() + '\n'], []
    for fn, override in FRONT:
        p = os.path.join(SRC, fn)
        if not os.path.exists(p):
            print('  !  缺 %s，跳过' % fn); continue
        title, body = prep(p)
        title = override or title
        names.append(title)
        parts.append('\n' + '%%PAGEBREAK%%' + '\n\n## ' + title + '\n\n' + body + '\n')
    parts.append('\n' + full[i:])

    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, 'book.md')
    io.open(out, 'w', encoding='utf-8').write('\n'.join(parts))
    mods = re.findall(r'(?m)^## (.+)$', io.open(out, encoding='utf-8').read())
    print('book.md 已生成：%d 个分册 -> %s' % (len(mods), out))
    for k, m in enumerate(mods, 1): print('  %d. %s' % (k, m))

main()
