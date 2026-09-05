#!/usr/bin/env python3
"""check_src.py —— 源码级校验，不依赖 LibreOffice

用法：
    python3 check_src.py [--src notes_src] [--quiet]

在 vault 根目录执行。检查项（任一失败退出码为 1）：
  1. front matter 缺失 / 无 id
  2. id 重复
  3. id 与文件名前缀不一致（警告，不致命）
  4. 双链指向不存在的笔记
  5. HTML 标签开闭不配对（em / strong / td / th / tr / table）
  6. 行首项目符号（▪ • - *）—— 会被 OneNote 转成列表，且渲染器不认
  7. 表格首行 colspan 合计与表内最大列数不符
     —— build_docx.js 的 nCols 只按首行推算，首行 colspan 写错会静默算错整表列宽
  8. MANIFEST.txt 条目数与实际笔记数不符（警告）

verify.py 查的是排版结果（空白页、跨页、标签泄漏到 PDF），
本脚本查的是源头。两者互补，不能互相替代。
"""
import io, os, re, sys, glob, collections

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

SRC   = os.path.abspath(_arg('--src', 'notes_src'))
QUIET = '--quiet' in sys.argv

FM      = re.compile(r'\A---\n(.*?)\n---\n', re.S)
WIKI    = re.compile(r'\[\[([^\]|]+?)(?:\|([^\]]+))?\]\]')
BULLET  = re.compile(r'(?m)^\s*(?:[▪•]|[-*]\s)')
TAGS    = ('em', 'strong', 'td', 'th', 'tr', 'table')
TR      = re.compile(r'<tr([^>]*)>([\s\S]*?)</tr>')
CELL    = re.compile(r'<(td|th)([^>]*)>([\s\S]*?)</\1>')
COLSPAN = re.compile(r'colspan="(\d+)"')


def table_widths(block):
    """返回该表块每一行的列宽合计（colspan 累加）。"""
    out = []
    for _cls, inner in TR.findall(block):
        out.append(sum(int(COLSPAN.search(at).group(1)) if COLSPAN.search(at) else 1
                       for _tag, at, _txt in CELL.findall(inner)))
    return out

errors, warns = [], []


def rel(p):
    return os.path.relpath(p, SRC)


def main():
    files = [f for f in glob.glob(os.path.join(SRC, '**', '*.md'), recursive=True)
             if os.sep + '.' not in f]
    if not files:
        sys.exit('未在 %s 下找到任何 .md' % SRC)

    names = {os.path.splitext(os.path.basename(f))[0] for f in files}
    ids   = collections.defaultdict(list)
    notes = 0

    for f in sorted(files):
        t = io.open(f, encoding='utf-8').read()
        m = FM.match(t)
        if not m:
            errors.append('缺 front matter: %s' % rel(f))
            continue

        sid = None
        for ln in m.group(1).split('\n'):
            if ln.startswith('id:'):
                sid = ln.split(':', 1)[1].strip().strip('"\'')
                break
        if not sid:
            errors.append('front matter 无 id: %s' % rel(f))
            continue

        is_note = bool(re.fullmatch(r'\d+(\.\d+)?', sid))
        if is_note:
            notes += 1
            ids[sid].append(rel(f))
            # id 与文件名前缀是否一致
            base = os.path.basename(f)
            pre  = re.match(r'(\d+(?:\.\d+)?)\s', base)
            if pre and pre.group(1) != sid:
                warns.append('id(%s) 与文件名前缀(%s) 不一致: %s' % (sid, pre.group(1), rel(f)))

        body = t[m.end():]

        # 行首项目符号：只查正文笔记，MOC / 总目录用 markdown 列表是正常的
        if is_note:
            for m2 in BULLET.finditer(body):
                ln_no = body[:m2.start()].count('\n') + 1
                errors.append('行首项目符号 %s:%d 「%s」'
                              % (rel(f), ln_no, body[m2.start():m2.start() + 20].strip()))

        # 标签配对
        for tag in TAGS:
            o = len(re.findall(r'<%s[ >]' % tag, body))
            c = len(re.findall(r'</%s>' % tag, body))
            if o != c:
                errors.append('<%s> 开闭不配对 (%d 开 / %d 闭): %s' % (tag, o, c, rel(f)))

        # 表格列数一致性：build_docx.js 只按首行算 nCols，首行写错会静默毁掉整表列宽
        for block in re.split(r'\n\s*\n', body):
            if block.count('<tr') < 2:
                continue
            w = table_widths(block)
            if not w:
                continue
            if w[0] != max(w):
                ln_no = body[:body.find(block)].count('\n') + 1
                errors.append('表格首行列宽 %d，表内最大 %d（nCols 会被算成 %d）: %s:%d'
                              % (w[0], max(w), w[0], rel(f), ln_no))

        # 双链
        for m2 in WIKI.finditer(body):
            tgt = m2.group(1).strip()
            if tgt not in names:
                errors.append('断链 %s -> [[%s]]' % (rel(f), tgt))

    for sid, fs in sorted(ids.items()):
        if len(fs) > 1:
            errors.append('id 重复 %s: %s' % (sid, ' / '.join(fs)))

    mani = os.path.join(SRC, 'MANIFEST.txt')
    if os.path.exists(mani):
        n = len([l for l in io.open(mani, encoding='utf-8') if re.match(r'^\d', l)])
        if n != notes:
            warns.append('MANIFEST 条目 %d 与实际笔记 %d 不符' % (n, notes))
    else:
        warns.append('未找到 MANIFEST.txt')

    if not QUIET:
        print('笔记 %d 节，文件 %d 个' % (notes, len(files)))
    for w in warns:
        print('  警告  %s' % w)
    for e in errors:
        print('  错误  %s' % e)
    if not QUIET and not errors and not warns:
        print('  全部通过')

    sys.exit(1 if errors else 0)


if __name__ == '__main__':
    main()
