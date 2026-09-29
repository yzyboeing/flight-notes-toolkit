#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_book.py —— 拼出「全书合订本」的源文件 build/book.md

用法（在笔记主库根目录，assemble.py 跑完之后）：
    python3 tools/make_book.py [--src notes_src] [--out build]

顺序：封面 → 自动目录 → 六章正文（2026-09-29 起总目录与编排规范移出成品，见仓库根目录「整理规范/」）。

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
# 2026-09-29 用户决定：全库总目录、笔记编排规范、标记说明等整理说明不进成品，
# 移到仓库根目录「整理规范/」作为整理规范与提示词。成品只含封面、自动目录与六章正文。
FRONT = []

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
    # 总目录 / 编排规范里的 [[双链]] 在 Word 里还原为纯文本（与 assemble.py 一致）
    t = re.sub(r'\[\[([^\]|]+?)(?:\|([^\]]+))?\]\]', lambda m: (m.group(2) or m.group(1)).strip(), t)
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
    text = '\n'.join(parts)
    # 章首页简介（2026-09-29）：取各章「_N … 目录.md」front matter 的 blurb 字段，
    # 以 %%CHAPDESC%% 行跟在章标题后，供 build_docx.js 排章首页；没有 blurb 的章不加
    CN = '零一二三四五六七八九'
    def blurb(n):
        for d in os.listdir(SRC):
            if re.match(r'%d\s' % n, d) and os.path.isdir(os.path.join(SRC, d)):
                for f in os.listdir(os.path.join(SRC, d)):
                    if f.startswith('_') and f.endswith('.md'):
                        m = re.search(r'(?m)^blurb:\s*"?(.*?)"?\s*$', io.open(os.path.join(SRC, d, f), encoding='utf-8').read())
                        if m: return m.group(1)
        return None
    def addb(m):
        k = CN.find(m.group(1)); b = blurb(k) if k >= 0 else None
        return m.group(0) + ('\n\n%%CHAPDESC%% ' + b if b else '')
    text = re.sub(r'(?m)^## 第(.)章[^\n]*$', addb, text)
    io.open(out, 'w', encoding='utf-8').write(text)
    mods = re.findall(r'(?m)^## (.+)$', io.open(out, encoding='utf-8').read())
    print('book.md 已生成：%d 个分册 -> %s' % (len(mods), out))
    for k, m in enumerate(mods, 1): print('  %d. %s' % (k, m))

main()
