#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""remark_body.py —— 正文章节（第 1–5 章）按「克制版」统一标红（SD-27）

用法：python3 tools/remark_body.py <章号 1–5 或节号，如 1 / 1.6> [--src notes_src] [--write] [--report]

规则（只加减 <em>/<strong>，不改文字；写入前自检去标签全文不变）：
  红  ① 带限定的数值：同一分句里有 < > ≤ ≥、最大 / 最小 / 至少 / 不少于 / 不超过 / 以上 / 以下 / 以内 / 限制 / 极限 等；
        或位于表头含「限制 / 极限 / 限值 / 最大 / 最小 / 阈值 / 门槛」的列
      ② 硬性要求与限制性措辞（禁止 / 不得 / 不可 / 无法 / 不提供 / 必须 / 须 / 立即 / 仅……到分句结束）
      ③ 危险后果分句；④ 警示行整句
  粗  原有的黑粗全部保留；原有红色短语若不属于以上三类，改为黑粗（强调不丢，只是降为黑色）
  不标  解释、举例、计算过程里的数字（没有限定词的）
报告每个格子 / 段落里红色超过 2 处的位置，供人工复核。
"""
import io, os, re, sys, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import remark_quickref as B

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
SRC = os.path.abspath(_arg('--src', 'notes_src'))
TARGET = sys.argv[1]
LIMIT_HDR = re.compile(r'限制|极限|限值|最大|最小|阈值|门槛|标准')
# 正文的门槛信号：比较符、限定词，以及「达到 / 等待 / 持续 / 超过 / 低于…」这类表示门槛或时限的词
BODY_SIG = re.compile(B.SIG2.pattern + r'|改设|设为|调至|暖车|设 ')      # 与速查区同一套门槛词（SD-26 / SD-27）
EXAMPLE = re.compile(r'举例|示例|例如|例：|算例|比如')

def old_em_to_strong(s):
    """旧红色短语：含数字、要求、危险的交给新规则；其余转黑粗"""
    def f(m):
        inner = m.group(1); pl = re.sub(r'<[^>]+>', '', inner)
        if re.search(r'\d', pl) or B.REQ.search(pl) or B.DANGER.search(pl) or len(pl.strip()) < 2:
            return inner
        if '<strong>' in inner: return re.sub(r'</?strong>', '', inner).join(['<strong>', '</strong>'])
        return '<strong>' + inner + '</strong>'
    return re.sub(r'(?s)<em>((?:(?!</?em>).)*)</em>', f, s)

def strip_num_strong(s):
    return re.sub(r'<strong>([^<]*)</strong>', lambda m: m.group(1) if re.fullmatch(r'[\s\d.,%℃°<>≤≥±~–\-−/ftkgpsim]*', m.group(1).replace('&lt;','').replace('&gt;','')) else m.group(0), s)

def seg_mark(seg, limit_col):
    parts = re.split(r'(<br\s*/?>)', seg); res = []
    in_example = False
    for p in parts:
        if re.match(r'<br', p or ''): res.append(p); continue
        if EXAMPLE.search(re.sub(r'<[^>]+>', '', p)): in_example = True      # 举例之后的计算数字不标
        buf = []
        # 已在 <strong> 里的文字也要处理：逐文本节点
        for t in B.split_tags(p):
            if not t or t.startswith('<'): buf.append(t); continue
            pieces = re.split('([%s])' % B.PUNCT, t); o2 = []
            for pc in pieces:
                if pc in B.PUNCT or not pc: o2.append(pc); continue
                sent = next((x for x in re.split(r'(?<=[。；;])', t) if pc in x), t)
                sig = (bool(BODY_SIG.search(pc)) or limit_col
                       or (len(re.sub(r'\s', '', sent)) <= 70 and bool(BODY_SIG.search(sent)))) and not in_example
                o2.append(B.mark_text(pc, limit_col, sig))
            buf.append(''.join(o2))
        res.append(''.join(buf))
    return ''.join(res)

def flatten(s):
    """同名标签嵌套（<strong>…<strong>…</strong>…</strong>）压平成一层"""
    out = []; depth = {'em': 0, 'strong': 0}
    for tok in re.split(r'(</?(?:em|strong)>)', s):
        m = re.fullmatch(r'<(/?)(em|strong)>', tok)
        if not m: out.append(tok); continue
        tag = m.group(2)
        if m.group(1):
            depth[tag] -= 1
            if depth[tag] == 0: out.append(tok)
        else:
            depth[tag] += 1
            if depth[tag] == 1: out.append(tok)
    return ''.join(out)

def fix_nesting(s):
    # <strong> 里嵌 <em> 渲染器能处理；<em> 里嵌 <strong> 去掉内层 strong
    return re.sub(r'(?s)<em>((?:(?!</?em>).)*)</em>', lambda m: '<em>' + re.sub(r'</?strong>', '', m.group(1)) + '</em>', s)

def do_table(tb, report, where):
    hdr = re.search(r'(?s)<tr class="hdr">(.*?)</tr>', tb)
    heads = [re.sub(r'<[^>]+>', '', h) for h in re.findall(r'(?s)<th[^>]*>(.*?)</th>', hdr.group(1))] if hdr else []
    if heads[:3] == ['块', '主题', '条目'] or heads[:2] == ['节', '主题']: return tb      # 块索引 / 章索引不动
    occ = {}; out = []; ri = 0
    for m in re.finditer(r'(?s)(<tr([^>]*)>)(.*?)(</tr>)', tb):
        open_, cls, inner, close = m.group(1), m.group(2), m.group(3), m.group(4)
        if 'hdr' in cls: out.append((m.span(), m.group(0))); continue
        full = re.search(r'premise|note|warn', cls)
        cells = list(re.finditer(r'(?s)<td([^>]*)>(.*?)</td>', inner)); buf = []; last = 0; ci = 0
        for c in cells:
            while (ri, ci) in occ: ci += 1
            attr, txt = c.group(1), c.group(2)
            if 'warn' in cls:
                new = '<em>' + re.sub(r'</?(em|strong)>', '', txt.strip()) + '</em>'
            else:
                h = heads[ci] if (ci < len(heads) and not full) else ''
                new = fix_nesting(seg_mark(strip_num_strong(old_em_to_strong(txt)), bool(LIMIT_HDR.search(h)) and ci > 0))
            if new.count('<em>') > 2: report.append('%s：%s（红 %d 处）' % (where, re.sub(r'<[^>]+>', '', new)[:28], new.count('<em>')))
            buf.append(inner[last:c.start()] + '<td%s>%s</td>' % (attr, new)); last = c.end()
            rs = int((re.search(r'rowspan="(\d+)"', attr) or [0, 1])[1]); cs = int((re.search(r'colspan="(\d+)"', attr) or [0, 1])[1])
            for k in range(cs):
                for r2 in range(1, rs): occ[(ri + r2, ci + k)] = 1
            ci += cs
        buf.append(inner[last:])
        out.append((m.span(), open_ + ''.join(buf) + close)); ri += 1
    s = tb; 
    for (a, b), rep in reversed(out): s = s[:a] + rep + s[b:]
    return s

def do_file(f, report):
    t = io.open(f, encoding='utf-8').read()
    t = t.replace('&lt;', '').replace('&gt;', '')
    fm = re.match(r'(?s)^---\n.*?\n---\n', t); head = fm.group(0) if fm else ''
    body = t[len(head):]
    parts = re.split(r'(?s)(<table\b.*?</table>)', body); res = []
    item = '?'
    for p in parts:
        if p.startswith('<table'):
            res.append(do_table(p, report, os.path.basename(f).split(' ')[0] + ' ' + item)); continue
        ls = []
        for ln in p.split('\n'):
            mm = re.match(r'^#{3,4} ([A-Z]-\d+)', ln)
            if mm: item = mm.group(1)
            if (not ln.strip() or ln.startswith('#') or re.match(r'^(来源|详见|<!--|>|---)', ln.strip())
                    or '[[' in ln and len(re.sub(r'\[\[[^\]]*\]\]', '', ln).strip()) < 12):
                ls.append(ln); continue
            new = fix_nesting(seg_mark(strip_num_strong(old_em_to_strong(ln)), False))
            if new.count('<em>') > 2: report.append('%s %s 段落：%s（红 %d 处）' % (os.path.basename(f).split(' ')[0], item, re.sub(r'<[^>]+>', '', new)[:28], new.count('<em>')))
            ls.append(new)
        res.append('\n'.join(ls))
    new = flatten(head + ''.join(res))
    new = new.replace('', '&lt;').replace('', '&gt;')
    old = io.open(f, encoding='utf-8').read()
    pl = lambda s: re.sub(r'</?(em|strong)>', '', s)
    assert pl(old) == pl(new), '文字发生变化：' + f
    return old, new

def main():
    pat = os.path.join(SRC, TARGET + ' *', '*.md') if '.' not in TARGET else os.path.join(SRC, TARGET.split('.')[0] + ' *', TARGET + ' *.md')
    files = [f for f in sorted(glob.glob(pat)) if not os.path.basename(f).startswith('_')]
    report = []; tot = [0, 0, 0, 0]
    for f in files:
        old, new = do_file(f, report)
        tot[0] += old.count('<em>'); tot[1] += new.count('<em>'); tot[2] += old.count('<strong>'); tot[3] += new.count('<strong>')
        if '--write' in sys.argv and new != old: io.open(f, 'w', encoding='utf-8').write(new)
    print('%d 个文件：红 %d → %d，粗 %d → %d' % (len(files), *tot))
    print('红色超过 2 处的格子 / 段落：%d' % len(report))
    if '--report' in sys.argv:
        for r in report: print('  ' + r)
    if '--write' in sys.argv: print('已写入')

if __name__ == '__main__':
    main()
