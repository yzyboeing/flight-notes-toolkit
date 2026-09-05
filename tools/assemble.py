#!/usr/bin/env python3
"""assemble.py —— 从 Obsidian vault 组装章节 md 与全书 md

用法：
    python3 assemble.py [--src .] [--out build]

在 vault 根目录执行。识别规则：
  · 只收 front matter 里 id 为「数字」或「数字.数字」的笔记；MOC / 索引类（id: MOC-*）自动跳过
  · 章内顺序 = id 的数值顺序
  · 输出 build/mod0..5.md（分册）与 build/full.md（全书）
  · [[双链]] 在输出时还原为纯文本，Word 里不出现方括号

组装规则见 prompt/SKILL.md。
"""
import io, os, re, sys, glob

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

SRC = os.path.abspath(_arg('--src', '.'))
OUT = os.path.abspath(_arg('--out', 'build'))

MOD = {
    '0': '基础知识速查区',
    '1': '第一章　系统理论',
    '2': '第二章　机组训练手册',
    '3': '第三章　运行规范',
    '4': '第四章　模拟机训练',
    '5': '第五章　技术提示',
}

FM = re.compile(r'\A---\n(.*?)\n---\n', re.S)
WIKI = re.compile(r'\[\[([^\]|]+?)(?:\|([^\]]+))?\]\]')

def load():
    """返回 {id: (title, body)}"""
    out = {}
    for f in glob.glob(os.path.join(SRC, '**', '*.md'), recursive=True):
        if any(part.startswith('.') for part in os.path.relpath(f, SRC).split(os.sep)):  # 只跳过输入内部的隐藏目录
            continue
        t = io.open(f, encoding='utf-8').read()
        m = FM.match(t)
        if not m:
            continue
        sid = None
        for ln in m.group(1).split('\n'):
            if ln.startswith('id:'):
                sid = ln.split(':', 1)[1].strip().strip('"\'')
                break
        if not sid or not re.fullmatch(r'\d+(\.\d+)?', sid):
            continue                              # MOC / 索引类跳过
        body = t[m.end():]
        h = re.match(r'\s*#\s+([^\n]*)\n', body)
        title = h.group(1).strip() if h else sid
        body = body[h.end():] if h else body
        out[sid] = (title, body.strip())
    return out

def unwiki(t):
    """[[4.4 离场阶段航线规划]] → 4.4（Word 里只保留编号）；[[目标|显示]] → 显示"""
    def rep(m):
        if m.group(2):
            return m.group(2).strip()
        tgt = m.group(1).strip()
        h = re.match(r'(\d+(?:\.\d+)?)\s', tgt)
        return h.group(1) if h else tgt
    return WIKI.sub(rep, t)

def key(sid):
    a, _, b = sid.partition('.')
    return (int(a), int(b) if b else 0)

def demote(t):
    t = re.sub(r'(?m)^####\s', '##### ', t)
    t = re.sub(r'(?m)^###\s',  '#### ',  t)
    return re.sub(r'(?m)^##\s', '### ', t)

def main():
    secs = load()
    if not secs:
        sys.exit('未找到任何带数字 id 的笔记，检查 --src 路径')
    os.makedirs(OUT, exist_ok=True)
    full = ['# 737 理论知识笔记\n']
    total = 0
    for n in '012345':
        ids = sorted([s for s in secs if s.split('.')[0] == n], key=key)
        if not ids:
            continue
        total += len(ids)
        bodies = [re.sub(r'\n{3,}', '\n\n', '## %s\n\n%s' % (secs[i][0], unwiki(secs[i][1]))).strip()
                  for i in ids]
        content = '\n\n---\n\n'.join(bodies)
        if n == '5':                              # 每个技术提示单独起页
            content = re.sub(r'(?m)^(## 5\.\d+　.*)$', r'%%PAGEBREAK%%\n\n\1', content)
            content = re.sub(r'\A%%PAGEBREAK%%\n\n', '', content)
        io.open(os.path.join(OUT, 'mod%s.md' % n), 'w', encoding='utf-8').write(
            '# %s\n\n%s\n' % (MOD[n], content))
        full += ['\n%%PAGEBREAK%%\n', '\n## %s\n' % MOD[n], demote(content) + '\n']
    io.open(os.path.join(OUT, 'full.md'), 'w', encoding='utf-8').write('\n'.join(full))
    print('assembled %d 节 -> %s' % (total, OUT))

if __name__ == '__main__':
    main()
