#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""remark_quickref.py —— 速查区统一标红 / 加粗（SD-26）

用法：python3 tools/remark_quickref.py [--src notes_src] [--write] [--report]

规则（只加减 <em>/<strong> 标记，不改任何文字）：
  红（<em>）  R1 限制值 / 门槛值：带 < > ≤ ≥、最大 / 最小 / 不少于 / 至少 / 以上 / 以下 等限定的数值；
                「限制类条目」表格里非首列的数值一律算限制值
              R2 硬性要求：禁止 / 严禁 / 不得 / 不可 / 不能 / 不要 / 必须 / 立即 / 只能 / 仅当 / 仅在 起到分句结束
              R3 危险后果：含 失速 / 失控 / 撞地 / 触地危险 / 停不住 / 超压 等的分句（≤ 36 字）
  粗（<strong>）保留原有的「关键条件 / 结论」短语；数字、名称、电门位这类短标记上的粗体去掉
  不标        「参考类条目」的数值（几何尺寸、参考下降率、计算示例…），表格首列的标签
自动结果交人工复核；OVERRIDE 里写逐条例外。
"""
import io, os, re, sys, glob

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
SRC = os.path.abspath(_arg('--src', 'notes_src'))

LIMIT_ITEMS = {46, 24, 3, 4, 5, 6, 8, 12, 15, 16, 17, 25, 26, 32, 33, 34, 47, 60, 63, 64, 72, 73, 87, 109}
REF_ITEMS = {1, 2, 14, 27, 35, 37, 85, 86, 88, 105, 106, 107}

UNIT = r'(?:N1|N2|人|单位|个点|段|kt 地速|ft/min|°/s|ft|fpm|psi|nm|km|kg|LB|lb|mph|mbar|inHg|hPa|min|kt|Hz|m/s|℃|°|%|g|m|s|h|V|夸脱|次|nm)'
NUMTOK = re.compile(
    r'VREF\d+(?:\s?\+\s?\d+)?'
    r'|(?:(?:最大|最小|最低|最高|最少|不少于|不低于|不高于|不小于|不大于|不超过|至少|最多)\s?)?'
    r'(?:[<>≤≥＜＞=±\ue001\ue002]\s?|\+(?=\d))?[-−]?(?:M\s?)?(?:\d+(?:\.\d+)?|\.\d+)(?:M(?![A-Za-z]))?(?:\s?(?:%|℃|°))?'
    r'(?:\s?(?:[–～~\-]|±|到|至)\s?[-−]?(?:M\s?)?(?:\d+(?:\.\d+)?|\.\d+)(?:M(?![A-Za-z]))?)?'
    r'(?:\s?' + UNIT + r')?'
    r'(?:（\d[^）]{0,10}）)?'
    r'(?:\s?(?:及以上|及以下|以上|以下|以内))?')
SIGNAL = re.compile(r'[<>≤≥＜＞\ue001\ue002]|最大|最小|最低|最高|不少于|不低于|不小于|不大于|不超过|至少|以上|以下|以内|超过|低于|高于|限制|极限|红线|上限|下限|间隔|持续')
REQ = re.compile(r'(?<!非)(?<!不是)(?<!并非)(?<!无需)(?<!不需)(?<!不)(?:禁止|严禁|不得|不可(?!预|靠|见|用)|不能(?!保证)|不要|不应|无法|不提供|不适用|不包括|不代表|必须|立即|只能|只可|(?<!必)(?<!无)须(?!知)|(?<!不)仅(?!供|为|是|作|考虑|表示|限于)|切勿|不允许)[^，。；;,（）()「—\n]{0,30}')
DANGER = re.compile(r'失速|失控|撞地|触地危险|不足以停住|超压状况|超轮速|无法放出|压力丧失')
PUNCT = '，。；;：:'

def esc_plain(s):
    return s.replace('&lt;', '<').replace('&gt;', '>')

def split_tags(s):
    return re.split(r'(<[^>]+>)', s)

def mark_text(txt, limit_mode, sig, nreq=False):
    """在一段纯文本里加 <em>：R2 → R3 → R1（后者不进入已标区域）"""
    spans = []
    for m in ([] if nreq else REQ.finditer(txt)):
        spans.append((m.start(), m.end()))
    # R3：以标点切分句
    pos = 0
    for part in ([] if nreq else re.split('([%s])' % PUNCT, txt)):
        if part and part not in PUNCT and DANGER.search(part) and len(part.strip()) <= 36:
            a = pos + (len(part) - len(part.lstrip())); b = pos + len(part.rstrip())
            if not any(x < b and a < y for x, y in spans): spans.append((a, b))
        pos += len(part)
    if limit_mode or sig:
        for m in NUMTOK.finditer(txt):
            tok = m.group(0).strip()
            if not tok or not re.search(r'\d', tok): continue
            has_unit = re.search(UNIT + r'|（|以上|以下|以内|M|VREF', tok) or re.search(r'[<>≤≥＜＞=±\ue001\ue002]|最|不|至少', tok)
            if not has_unit and not limit_mode: continue
            if re.fullmatch(r'\d', tok) and not limit_mode: continue
            a = m.start() + (len(m.group(0)) - len(m.group(0).lstrip())); b = a + len(tok)
            if any(x < b and a < y for x, y in spans): continue
            spans.append((a, b))
    spans.sort()
    out, last = [], 0
    for a, b in spans:
        if a < last: continue
        out.append(txt[last:a]); out.append('<em>' + txt[a:b] + '</em>'); last = b
    out.append(txt[last:])
    return ''.join(out)

def strip_marks(s):
    s = s.replace('<em>', '').replace('</em>', '')
    # 数字、ASCII、≤3 个汉字的短粗体去掉；其余保留（关键条件 / 结论）
    def f(m):
        inner = m.group(1)
        pl = re.sub(r'<[^>]+>', '', inner)
        cjk = len(re.findall(r'[一-鿿]', pl))
        return m.group(0) if (cjk >= 4 or pl.strip().endswith(('：', ':'))) else inner   # 「注意：」这类引导标签保留
    return re.sub(r'<strong>((?:(?!</?strong>).)*)</strong>', f, s, flags=re.S)

def process_segment(seg, limit_mode, first_col):
    """seg 是一个格子或段落的内部 HTML（可能含 <br>、<strong>）"""
    if first_col: limit_mode, allow_sig = False, False
    else: allow_sig = True
    parts = re.split(r'(<br\s*/?>)', seg)
    res = []
    for p in parts:
        if re.match(r'<br', p or ''): res.append(p); continue
        sig = allow_sig and bool(SIGNAL.search(re.sub(r'<[^>]+>', '', p)))
        # 按分句决定是否有限定信号，逐段处理文本节点
        toks = split_tags(p); depth = 0; buf = []
        for t in toks:
            if t.startswith('<'):
                buf.append(t); continue
            if not t: continue
            if first_col:      # 首列：只标带单位的数值（limit_mode 关），不标要求类短语
                pieces = re.split('([%s])' % PUNCT, t)
                buf.append(''.join(pc if (pc in PUNCT or not pc) else mark_text(pc, False, True, nreq=True) for pc in pieces)); continue
            # 分句级信号
            pieces = re.split('([%s])' % PUNCT, t); o2 = []
            for pc in pieces:
                if pc in PUNCT or not pc: o2.append(pc); continue
                s2 = True
                o2.append(mark_text(pc, limit_mode, s2))
            buf.append(''.join(o2))
        res.append(''.join(buf))
    return ''.join(res)

def process_item(n, body, ref, lim):
    lines = body.split('\n'); out = []
    for ln in lines:
        if re.match(r'^(来源|详见)', ln) or ln.startswith('#'): out.append(ln); continue
        out.append(ln)
    s = '\n'.join(out)
    head = re.match(r'(?s)((?:(?:来源|详见)[^\n]*\n)*)', s).group(1)
    rest = strip_marks(s[len(head):])
    def cell(m):
        tag, attr, inner = m.group(1), m.group(2), m.group(3)
        if tag == 'th': return m.group(0)
        return '<td%s>%s</td>' % (attr, process_segment(inner, lim, m.group(4) == 'first'))
    # 标注首列：在每个 <tr> 的第一个 <td> 上做记号
    def mark_first(tr):
        body = tr.group(2)
        if 'class="hdr"' in tr.group(1) or 'colspan' in body.split('</td>')[0] and ('premise' in tr.group(1) or 'note' in tr.group(1) or 'warn' in tr.group(1)):
            return tr.group(0)
        return tr.group(0)
    # 逐表处理，首列判定：行内第一个 td 且该行不是通栏行，且该 td 在表格第 0 列（rowspan 占位后第一个 td 不是首列）
    def do_table(tm):
        tb = tm.group(0); rows = re.findall(r'(?s)(<tr[^>]*>)(.*?)(</tr>)', tb)
        occ0 = 0; newrows = []
        for open_, inner, close in rows:
            full = re.search(r'class="(premise|note|warn)"', open_)
            cells = list(re.finditer(r'(?s)<t([hd])([^>]*)>(.*?)</t[hd]>', inner))
            ncols = max(len(re.findall(r'<th', tb.split('</tr>')[0])), 2) if 'class="hdr"' in tb else 2
            buf, last = [], 0
            col0_taken = occ0 > 0
            for k, c in enumerate(cells):
                first = (k == 0 and not col0_taken and not full and ncols > 1)
                buf.append(inner[last:c.start()])
                if c.group(1) == 'h': buf.append(c.group(0))
                else:
                    if full and 'warn' in full.group(1):
                        seg = '<em>' + re.sub(r'</?em>', '', strip_marks(c.group(3)).strip()) + '</em>'
                    else:
                        seg = process_segment(c.group(3), lim and not full, first)
                    buf.append('<td%s>%s</td>' % (c.group(2), seg))
                last = c.end()
                if first:
                    rs = re.search(r'rowspan="(\d+)"', c.group(2)); occ0 = int(rs.group(1)) if rs else 1
            buf.append(inner[last:])
            if occ0 > 0: occ0 -= 1
            newrows.append(open_ + ''.join(buf) + close)
        it = iter(newrows)
        return re.sub(r'(?s)<tr[^>]*>.*?</tr>', lambda m: next(it), tb)
    parts = re.split(r'(?s)(<table\b.*?</table>)', rest); res = []
    for p in parts:
        if p.startswith('<table'): res.append(do_table(re.match(r'(?s).*', p)))
        else:
            ls = []
            for ln in p.split('\n'):
                ls.append(ln if (not ln.strip() or ln.startswith('#')) else process_segment(ln, False, False))
            res.append('\n'.join(ls))
    rest = ''.join(res)
    if ref:   # 参考类条目：数值不标，只保留 R2 / R3
        rest = re.sub(r'<em>([^<]*\d[^<]*)</em>', lambda m: m.group(0) if REQ.match(m.group(1)) or DANGER.search(m.group(1)) else m.group(1), rest)
    return head + rest

def main():
    qf = glob.glob(os.path.join(SRC, '0 *', '0 *.md'))[0]
    q = io.open(qf, encoding='utf-8').read()
    cut = q.index('## 按物理量索引')
    body, tail = q[:cut].replace('&lt;', '\ue001').replace('&gt;', '\ue002'), q[cut:]
    parts = re.split(r'(?m)(^### \d+\. [^\n]*\n)', body)
    out = [parts[0]]; stats = []
    for i in range(1, len(parts), 2):
        h, b = parts[i], parts[i + 1]
        n = int(re.match(r'### (\d+)', h).group(1))
        # 块标题（## …）在 b 的末尾，分开处理
        m = re.search(r'(?m)^## ', b)
        main_, blk = (b[:m.start()], b[m.start():]) if m else (b, '')
        nb = process_item(n, main_, n in REF_ITEMS, n in LIMIT_ITEMS)
        stats.append((n, main_.count('<em>'), nb.count('<em>'), main_.count('<strong>'), nb.count('<strong>')))
        out += [h, nb, blk]
    new = (''.join(out) + tail).replace('\ue001', '&lt;').replace('\ue002', '&gt;')
    plain = lambda s: re.sub(r'</?(em|strong)>', '', s)
    assert plain(new) == plain(q), '文字发生了变化，中止'
    if '--report' in sys.argv:
        for s in stats: print('第 %d 条：红 %d → %d，粗 %d → %d' % s)
    print('合计：红 %d → %d，粗 %d → %d' % (sum(s[1] for s in stats), sum(s[2] for s in stats), sum(s[3] for s in stats), sum(s[4] for s in stats)))
    if '--write' in sys.argv:
        io.open(qf, 'w', encoding='utf-8').write(new); print('已写入')

if __name__ == '__main__':
    main()
