#!/usr/bin/env python3
"""单册追加正文节（SD-116）：把指定的正文节（默认 4.21 记忆项目、4.22 机动飞行）按速查区的层级追加到 single.md 末尾，
成为单册目录里可点击的主题组。

用法：python3 single_extra.py build/single.md build/mod4.md [4.21,4.22]
- 节标题「## 4.21　记忆项目」→ 主题组「### 记忆项目（4.21）」；条目「### N. 标题」→「#### N. 标题」（编号不变，与全书一致）
- 单册里没有的章节的「详见」行去掉（避免死链接）；节内的「见第 N 条」保留（指本组）
"""
import re, sys

single, src = sys.argv[1], sys.argv[2]
want = (sys.argv[3] if len(sys.argv) > 3 else '4.21,4.22').split(',')
lines = open(src, encoding='utf-8').read().split('\n')
out, cur = [], None
for ln in lines:
    m = re.match(r'^## (\d+\.\d+)\s*　?\s*(.+)$', ln)
    if m:
        cur = m.group(1) if m.group(1) in want else None
        if cur:
            out += ['', '', f'### {m.group(2).strip()}（{cur}）', '']
        continue
    if cur is None:
        continue
    if re.match(r'^#\s', ln):          # 章标题，结束
        cur = None; continue
    if re.match(r'^详见\s', ln.strip()):
        continue
    ln = re.sub(r'^### ', '#### ', ln)
    out.append(ln)
if not out:
    sys.exit('single_extra：没有找到要追加的节 ' + ','.join(want))
# 追加内容要落在速查区的 %%COMPACT%% 范围内，生成器才会给主题组建目录条目（可点击）
body = open(single, encoding='utf-8').read().rstrip('\n')
extra = '\n'.join(out).rstrip('\n')
if body.endswith('%%ENDCOMPACT%%'):
    body = body[:-len('%%ENDCOMPACT%%')].rstrip('\n') + '\n' + extra + '\n\n%%ENDCOMPACT%%\n'
else:
    body = body + '\n\n%%COMPACT%%\n' + extra + '\n\n%%ENDCOMPACT%%\n'
open(single, 'w', encoding='utf-8').write(body)
print('单册追加：' + '、'.join(want))
