"""两本标题与主题对照表：速查每条 → 绑定块标题、主题归属；完整版每块 → 主题归属"""
import io, os, re, json, sys
R = os.path.expanduser('~/flight-repos/gh-private')
OUT = sys.argv[1]
# 完整版块标题
blk = {}; order = []
for root, _, fs in os.walk(R + '/notes_src'):
    for f in sorted(fs):
        m = re.match(r'(\d\.\d+) ', f)
        if not (f.endswith('.md') and m): continue
        sec = m.group(1)
        for line in io.open(os.path.join(root, f), encoding='utf-8'):
            line = re.sub(r'\s*\{★+\}\s*$', '', line)
            h = re.match(r'### ([A-Z]-\d+)\s*　?(.*)', line) or re.match(r'### (\d+)\.\s*(.*)', line)
            if h:
                k = '%s %s' % (sec, h.group(1)); blk[k] = h.group(2).strip(); order.append(k)
def skey(k):
    a, b = k.split(' '); x, y = a.split('.'); return (int(x), int(y), b)
order.sort(key=skey)
# 完整版主题
# 2026-10-23（SD-233）：完整版主题取「速查目录」（qr_directory.py 现算；只记直接项），原「按主题查索引.json」已撤
import subprocess, tempfile
_tmp = os.path.join(tempfile.mkdtemp(), 'd.json')
subprocess.run([sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'qr_directory.py'), '--repo', R, '--out', _tmp], check=True, capture_output=True)
full = json.load(io.open(_tmp, encoding='utf-8'))
fth = {}
for d in full:
    for t in d['themes']:
        for it in t['items']:
            p = it.split('|')
            if len(p) > 2 and p[2] == 'x': continue
            fth.setdefault(p[0], []).append('%s›%s' % (d['dim'], t['theme']))
# 速查条目
q = io.open(R + '/速查/速查源.md', encoding='utf-8').read()
qj = json.load(io.open(R + '/速查/速查按主题查.json', encoding='utf-8'))
short = qj.get('short', {})
part = grp = ''; n = 0; rows = []
for line in q.split('\n'):
    if line.startswith('%%PART%%'): part = line[9:].strip()
    elif line.startswith('## '): grp = line[3:].strip()
    elif line.startswith('### '):
        line = re.sub(r'\s*\{★+\}\s*$', '', line)
        n += 1; rows.append([n, part, grp, line[4:].strip(), []])
    elif line.startswith('<!-- 详见') and rows and not rows[-1][4]:
        rows[-1][4] = re.findall(r'\|([^\]]+)\]\]', line)
qth = []
for d in qj['dims']:
    for t in d.get('themes', []):
        qth.append((d['dim'], t))
def qthemes(title, blocks):
    out = []
    for dim, t in qth:
        hit = (any(norm(b) in t.get('items', []) for b in blocks) and title not in t.get('exclude', [])) or title in t.get('see', [])
        if hit:
            o = t.get('order', [])
            out.append('%s›%s%s' % (dim, t['theme'], '（序 %d）' % (o.index(title) + 1) if title in o else ''))
    return out
def norm(b):
    b = b.replace('第 ', '').replace(' 条', ''); return b
L = ['# 两本标题与主题对照表（1016R9）', '',
     '> 自动生成，供审核「标题规范化、一致化与主题目录逻辑顺序」用。速查标题与绑定块标题不同不一定是问题（速查汇编、有意精简），逐条判断。「速查主题归属」只列本段本组以外的主题目录归属（↗ 交叉）与阶段主题里的顺序（序 n）；本组默认归属不列。', '',
     '## 一、速查条目 ↔ 完整版绑定块', '',
     '| 条 | 段 › 组 | 速查标题 | 主题目录短标题 | 速查主题归属 | 绑定块 → 完整版标题 | 完整版主题归属 |', '|---|---|---|---|---|---|---|']
for n, p, g, t, b in rows:
    bt = '<br>'.join('%s：%s' % (x, blk.get(norm(x), '（找不到块）')) for x in b) or '（无绑定）'
    ft = '<br>'.join(sorted({y for x in b for y in fth.get(norm(x), [])})) or '—'
    st = short.get(t, '')
    qt = '<br>'.join(qthemes(t, b)) or '—'
    L.append('| %d | %s › %s | %s | %s | %s | %s | %s |' % (n, p, g, t, st, qt, bt, ft))
L += ['', '## 二、完整版各块主题归属（未挂主题的块标「—」）', '', '| 块 | 标题 | 主题 |', '|---|---|---|']
for k in order:
    L.append('| %s | %s | %s |' % (k, blk[k], '<br>'.join(fth.get(k, [])) or '—'))
io.open(OUT, 'w', encoding='utf-8').write('\n'.join(L) + '\n')
print('速查', len(rows), '条；完整版', len(order), '块；未挂主题', sum(1 for k in order if k not in fth), '块；无绑定速查', sum(1 for r in rows if not r[4]))
