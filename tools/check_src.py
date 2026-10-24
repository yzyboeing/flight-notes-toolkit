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
import json, io, os, re, subprocess, sys, glob, collections

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

SRC   = os.path.abspath(_arg('--src', 'notes_src'))
QUIET = '--quiet' in sys.argv

FM      = re.compile(r'\A---\n(.*?)\n---\n', re.S)
WIKI    = re.compile(r'\[\[([^\]|]+?)(?:\|([^\]]+))?\]\]')
BULLET  = re.compile(r'(?m)^\s*(?:[▪•]|[-*]\s)')
TAGS    = ('em', 'strong', 'td', 'th', 'tr', 'table')
TABLE   = re.compile(r'<table[^>]*>[\s\S]*?</table>', re.I)
TR      = re.compile(r'<tr([^>]*)>([\s\S]*?)</tr>')
CELL    = re.compile(r'<(td|th)([^>]*)>([\s\S]*?)</\1>')
COLSPAN = re.compile(r'colspan="(\d+)"')
FORBIDDEN_MARKS = re.compile(r'〔待补[^〕]{0,12}〕|\[待补[^\]]{0,12}\]|〔待确认〕|\[待确认\]|【缺失[^】]{0,20}】|待核实|待核对|待原件补充|待补充来源|部分内容待补来源|手册未写|手册中未见|原文未收录')   # SD-48 / SD-50：终稿不允许任何待补、待核类备注


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


MIX = []
def main():
    files = [f for f in glob.glob(os.path.join(SRC, '**', '*.md'), recursive=True)
             if not any(part.startswith('.') for part in os.path.relpath(f, SRC).split(os.sep))]
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
            for m2 in FORBIDDEN_MARKS.finditer(body):
                ln_no = body[:m2.start()].count('\n') + 1
                errors.append('SD-48 废弃标记残留 %s:%d 「%s」'
                              % (rel(f), ln_no, m2.group(0)))

        # 标签配对
        for tag in TAGS:
            o = len(re.findall(r'<%s[ >]' % tag, body))
            c = len(re.findall(r'</%s>' % tag, body))
            if o != c:
                errors.append('<%s> 开闭不配对 (%d 开 / %d 闭): %s' % (tag, o, c, rel(f)))
        # SD-204（2026-10-17 用户）：首列一律不加点——首列表头 col-bullet、行首格 cell-bullet 报错（并列对照表 parallel 除外；生成器也会忽略）
        for tm in re.finditer(r'<table([^>]*)>(.*?)</table>', body, re.S):
            if 'parallel' in tm.group(1): continue
            hm = re.search(r'<tr class="hdr"><th([^>]*)>', tm.group(2))
            occ = set(); bad = bool(hm and re.search(r'\bcol-(bullet|left)\b', hm.group(1)))
            for ri, rr in enumerate(re.findall(r'<tr(?![^>]*class="(?:hdr|premise|note)")[^>]*>(.*?)</tr>', tm.group(2), re.S)):
                k = 0
                for a2 in re.findall(r'<td([^>]*)>', rr):
                    while (ri, k) in occ: k += 1
                    cs = int((re.search(r'colspan="(\d+)"', a2) or [0, 1])[1]); rs = int((re.search(r'rowspan="(\d+)"', a2) or [0, 1])[1])
                    if k == 0 and cs == 1 and 'cell-bullet' in a2: bad = True
                    for dr in range(rs):
                        for dc in range(cs): occ.add((ri + dr, k + dc))
                    k += cs
            if bad: errors.append('SD-204 首列加点标注 %s 「%s」' % (rel(f), re.sub(r'<[^>]+>', '', tm.group(2))[:24].strip()))
        # SD-138（2026-10-05 用户）：全书不写来源、出处行；图注不写「（来源：……）」
        for ln_no, line in enumerate(body.split('\n'), 1):
            if re.match(r'^[ \t]*(?:>[ \t]*)?(?:\*\*)?(来源|出处)[：:]', line) or '（来源：' in line:
                errors.append('SD-138 来源 / 出处 %s:%d' % (rel(f), ln_no))
            # SD-213（2026-10-20 用户第 87 条「条目标题后的星标两本一并去掉」）：标题不再标 {★} / {★★}
            if re.match(r'^#{1,4}\s.*\{★+\}\s*$', line):
                errors.append('SD-213 标题带星级标记 %s:%d' % (rel(f), ln_no))
            # SD-212（2026-10-20 用户「图注里的手册出处都去掉，按『引用只写名字』的原则统一处理」）：图注不写手册出处与「依据……」
            cm = re.match(r'%%FIG(?:SIDE|RIGHT)?%%[^|]*\|([^|]*)', line)
            if cm and re.search(r'(?:FCOM|FCTM|AFM|QRH) [\d.]+|AC-\d+-\d+ 图|技术提示汇总|依据', cm.group(1)):
                errors.append('SD-212 图注写了手册出处 %s:%d「%s」' % (rel(f), ln_no, cm.group(1).strip()[:40]))
        # SD-140（2026-10-05 用户）：全书不写「详见」——「本身目录和前面的序号都会改变」
        for ln_no, line in enumerate(body.split('\n'), 1):
            if '详见' in line and not line.lstrip().startswith('<!--'):
                errors.append('SD-140 详见 %s:%d' % (rel(f), ln_no))
        # 2026-10-07 P1-L023/L036：补「见 A 块」「见下方 / 上方」「按上述」
        # SD-137（2026-10-04 用户）：表格内不出现交叉引用（见 x.y / 第 N 条 / 见 A-1 / 见上表 等），需要的写进表后「注：……——详见 x.y」；节首不写「本节为……」导语
        intbl = False
        for ln_no, ln in enumerate(t.split('\n'), 1):
            if '<table' in ln: intbl = True
            if intbl and not ln.startswith('<!--') and 'class="note"' not in ln:
                plain = re.sub(r'<[^>]+>', '', ln)
                m = re.search(r'(?:详见|参见|另见|(?<![可显明预少罕所常意])见)\s*(?:\[\[|\d\.\d{1,2}|第\s*\d+\s*条|[A-H]-\d+|[A-H]\s*块|上表|下表|上行|下行|上方|下方|本节)', plain) or re.search(r'(?:按|同)\s*第\s*\d+\s*条|按上述|同上述', plain)
                if m:
                    errors.append('SD-137 表格内交叉引用 %s:%d 「%s」' % (rel(f), ln_no, plain[max(0, m.start() - 15):m.end() + 10].strip()))
            if '</table>' in ln: intbl = False
            if re.match(r'^(?:<strong>)?本节(?:为|是|只|按)', ln):
                errors.append('SD-137 节首说明性导语 %s:%d' % (rel(f), ln_no))

        # 表格列数一致性：build_docx.js 只按首行算 nCols，首行写错会静默毁掉整表列宽
        # 按 <table>…</table> 切分，与 build_docx.js 界定表格的方式一致；
        # 不按空行切，否则表内出现空行会被拆成两个假表而误报
        for mt in TABLE.finditer(body):
            w = table_widths(mt.group(0))
            if len(w) < 2:
                continue
            if w[0] != max(w):
                ln_no = body[:mt.start()].count('\n') + 1
                errors.append('表格首行列宽 %d，表内最大 %d（nCols 会被算成 %d）: %s:%d'
                              % (w[0], max(w), w[0], rel(f), ln_no))

        # SD-170：两型共有的知识点用对比表，不在格内 / 段落里写「【737-NG】…／【737-8】…」（建议级，只计数；SRC_MIX=1 列出明细）
        for seg in re.findall(r'<td\b[^>]*>.*?</td>', body, re.S) + [l for l in re.sub(r'<table\b.*?</table>', '', body, flags=re.S).split('\n')]:
            if '【737-NG' in seg and '【737-8' in seg:
                MIX.append('%s:%d「%s」' % (rel(f), body[:body.find(seg)].count('\n') + 1 if seg in body else 0, re.sub(r'<[^>]+>|\s+', '', seg)[:24]))

        # 2026-10-10 用户（速查第 35 页 LE FLAPS TRANSIT 表）：对比表首列必须是「项目」，表头第一格就是机型名会被当成标签列排窄加粗
        for mh in re.finditer(r'<table\b[^>]*>\s*(?:<tr class="premise">.*?</tr>\s*)?<tr class="hdr">\s*<th[^>]*>(.*?)</th>', body, re.S):
            if re.match(r'\s*(?:<[^>]+>)*\s*737', mh.group(1)):
                errors.append('对比表缺项目列（表头首格是机型名）%s:%d' % (rel(f), body[:mh.start()].count('\n') + 1))

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

    qsrc = os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '速查', '速查源.md')   # 2026-10-09 用户：速查也按 SD-170 查（此前漏查）
    if os.path.exists(qsrc):
        qt = io.open(qsrc, encoding='utf-8').read()
        for mh in re.finditer(r'<table\b[^>]*>\s*(?:<tr class="premise">.*?</tr>\s*)?<tr class="hdr">\s*<th[^>]*>(.*?)</th>', qt, re.S):
            if re.match(r'\s*(?:<[^>]+>)*\s*737', mh.group(1)):
                errors.append('对比表缺项目列（表头首格是机型名）速查/速查源.md:%d' % (qt[:mh.start()].count('\n') + 1))
        for seg in re.findall(r'<td\b[^>]*>.*?</td>', qt, re.S) + [l for l in re.sub(r'<table\b.*?</table>', '', qt, flags=re.S).split('\n') if not l.startswith(('<!--', '#'))]:
            if '【737-NG' in seg and '【737-8' in seg:
                MIX.append('速查/速查源.md:%d「%s」' % (qt[:qt.find(seg)].count('\n') + 1, re.sub(r'<[^>]+>|\s+', '', seg)[:24]))
    # SD-196：速查主题 JSON 的 order / see / exclude / short 引用必须能前缀匹配到某条速查标题（失效引用曾掩盖条目被删）
    _qj = os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '速查', '速查按主题查.json')
    if os.path.exists(qsrc) and os.path.exists(_qj):
        _T = [l[4:].strip() for l in qt.split('\n') if l.startswith('### ')]
        _q = json.load(io.open(_qj, encoding='utf-8'))
        _refs = [('short', k) for k in _q.get('short', {})] + [('%s.%s' % (th.get('theme'), k), x) for d in _q.get('dims', []) for th in d.get('themes', []) for k in ('order', 'see', 'exclude') for x in th.get(k, [])]
        for _w, _x in _refs:
            if not any(y.startswith(_x) for y in _T):
                errors.append('速查主题 JSON 失效引用（%s）「%s」：没有以此开头的速查条目（SD-196）' % (_w, _x))
    # ---- SD-209（2026-10-19 整体优化方案批 1）：S6 速查同名、S7 相对上一基线的丢失 / 重复、S8 节粒度、T36 标题规范、T37 速查顺序、L6 插图 ----
    ROOT = os.path.dirname(os.path.abspath(SRC.rstrip('/')))
    if os.path.exists(qsrc):
        import subprocess, collections as _co
        _ents = [(m.group(1).strip(), m.group(2), m.group(3)) for m in re.finditer(r'(?ms)^### (.+?)\n(<!-- 详见.*?-->\n)?(.*?)(?=^### |^## |^%%PART%%|\Z)', qt)]
        for _t, _n in _co.Counter(e[0] for e in _ents).items():
            if _n > 1: errors.append('S6 速查同名条目「%s」×%d' % (_t, _n))
        # K-T39（SD-209 批 3 Q08）：速查每条第一行「%%KEY%% 要点」，宽度 ≤ 30 字（半角算半字）
        for _t, _, _b in _ents:
            _f = next((x for x in _b.split('\n') if x.strip()), '')
            if not _f.startswith('%%KEY%%'): errors.append('T39 速查条目「%s」缺「要点」行（Q08）' % _t); continue
            _k = _f[7:].strip().replace(' ', '')
            if sum(0.5 if ord(c) < 128 else 1 for c in _k) > 30: errors.append('T39 速查条目「%s」要点超过 30 字（Q08）' % _t)
        # S7：相对上一个 baseline/* 标签，速查每条原有表行是否还在；同条目内块与块重复
        _P = lambda x: re.sub(r'[^\w]', '', re.sub(r'<[^>]+>', '', re.sub(r'<!--.*?-->', '', x, flags=re.S)))
        try:
            _tags = subprocess.run(['git', '-C', ROOT, 'tag', '-l', 'baseline/*'], capture_output=True, text=True).stdout.split()
            _base = sorted(_tags)[-1] if _tags else ''
            _old = subprocess.run(['git', '-C', ROOT, 'show', '%s:速查/速查源.md' % _base], capture_output=True, text=True).stdout if _base else ''
        except Exception: _old = ''
        if _old and not os.environ.get('SRC_NO_S7'):
            _nn = _P(qt); _NS = set(_nn[i:i + 4] for i in range(len(_nn) - 3))
            def _rc(x):
                x = _P(x); sh = [x[i:i + 4] for i in range(len(x) - 3)]
                return sum(1 for y in sh if y in _NS) / max(1, len(sh))
            _lost = 0
            try: _DEL = set(json.load(open(os.path.join(ROOT, '速查删行台账.json'), encoding='utf-8')).get('行', {}))
            except (OSError, ValueError): _DEL = set()
            for m in re.finditer(r'(?ms)^### (.+?)\n(.*?)(?=^### |^## |^%%PART%%|\Z)', _old):
                for r in re.findall(r'<tr.*?</tr>', m.group(2), re.S):
                    if len(_P(r)) >= 8 and _rc(r) < 0.6 and _P(r)[:30] not in _DEL:   # 2026-10-21：《速查删行台账.json》登记的已确认删除不报
                        _lost += 1
                        if _lost <= 20: errors.append('S7 速查「%s」原有表行在新文件里找不到（相对 %s）：「%s」' % (m.group(1)[:20], _base, _P(r)[:30]))
            if _lost > 20: errors.append('S7 …共 %d 行' % _lost)
        for _t, _, _b in _ents:
            _blocks = re.findall(r'<table.*?</table>', _b, re.S) + [l for l in re.sub(r'<table.*?</table>', '', _b, flags=re.S).split('\n') if len(_P(l)) >= 10 and not l.startswith('%%KEY%%')]   # 要点行是全条摘要，不算重复块
            _sh = [set(_P(x)[i:i + 4] for i in range(len(_P(x)) - 3)) for x in _blocks]
            for i in range(len(_blocks)):
                for j in range(len(_blocks)):
                    if i != j and _sh[i] and len(_sh[i]) <= len(_sh[j]) and len(_sh[i] & _sh[j]) / len(_sh[i]) >= 0.8 and len(_P(_blocks[i])) >= 30:
                        errors.append('S7 速查「%s」块内重复：「%s…」与「%s…」' % (_t[:20], _P(_blocks[i])[:16], _P(_blocks[j])[:16])); break
        # T36 标题规范（W05 H6）：vs / 破折号 / 问号 / 末尾标点
        for _t, _, _ in _ents:
            if re.search(r'\svs\s|——|？|\?|[。；，]$', _t): errors.append('T36 速查标题「%s」含 vs / 破折号 / 问号 / 末尾标点（W05）' % _t)
        # T37 速查同一绑定节的条目按正文条目号排（提示）
        # 2026-10-19 批 3：只比同一「## 主题」内、绑定同一节的条目；绑 ≥ 3 条的汇总条目（限制总表、差异汇编、要点）不比
        _seq = {}
        for _m in re.finditer(r'(?ms)^## (.+?)\n(.*?)(?=^## |^%%PART%%|\Z)', qt):
            for _t, _bind, _ in [(m.group(1).strip(), m.group(2), '') for m in re.finditer(r'(?ms)^### (.+?)\n(<!-- 详见.*?-->\n)?', _m.group(2))]:
                _bs = re.findall(r'\|(\d+\.\d+) 第 (\d+) 条\]\]', _bind or '')
                if not _bs or len(_bs) >= 3: continue
                _seq.setdefault('%s·%s' % (_m.group(1).strip(), _bs[0][0]), []).append((int(_bs[0][1]), _t))
        _bad = [(k, v) for k, v in _seq.items() if len(v) >= 3 and any(v[i][0] > v[i + 1][0] for i in range(len(v) - 1))]
        if _bad: warns.append('T37 速查条目未按正文条目号排（N1）：%s' % '；'.join('%s（%s）' % (k, ' → '.join(str(x[0]) for x in v)) for k, v in _bad[:12]))
    # S8 节粒度（N2：5～25 条）；T36 完整版标题；L6 插图文件与图注
    for f in files:
        b = os.path.basename(f)
        if not re.match(r'[1-6]\.\d+ ', b): continue
        t = io.open(f, encoding='utf-8').read()
        n = len(re.findall(r'(?m)^### \d+\. ', t))
        if n and (n < 5 or (n > 25 and not b.startswith('1.'))) and not b.startswith(('4.', '5.', '6.')): warns.append('S8 %s 共 %d 条（N2：一节 5～25 条；第四章技术提示按通告分节、第五章按 QRH 分节，不计）' % (b[:-3], n))
        for m in re.finditer(r'(?m)^### \d+\. (.+?)(?:\s*\{★+\})?$', t):
            if re.search(r'\svs\s|——|？|\?|[。；，]$', m.group(1)): errors.append('T36 %s 标题「%s」含 vs / 破折号 / 问号 / 末尾标点（W05）' % (b[:-3], m.group(1)))
        for m in re.finditer(r'(?m)^%%FIG(?:SIDE|RIGHT)?%% ([^|\n]+?)\s*\|\s*([^|\n]*?)\s*(?:\||$)', t):
            fp = os.path.join(SRC, m.group(1).strip())
            if not os.path.exists(fp): errors.append('L6 %s 插图文件不存在：%s' % (b[:-3], m.group(1).strip()))
            if not m.group(2).strip(): errors.append('L6 %s 插图缺图注：%s' % (b[:-3], m.group(1).strip()))
    # W13 手册用语（SD-177 / 215，2026-10-21 方案 A2）：私有库《手册用语表.json》里的禁用说法，两本去标签后逐字查（建议级）
    _wf = os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '手册用语表.json')
    if os.path.exists(_wf):
        _ban = [(x['禁用'], x['改用']) for x in json.load(io.open(_wf, encoding='utf-8')).get('条目', [])]
        _qf = os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '速查', '速查源.md')
        for f in list(files) + ([_qf] if os.path.exists(_qf) else []):
            for _i, _ln in enumerate(io.open(f, encoding='utf-8'), 1):
                _p = re.sub(r'<[^>]+>', '', _ln)
                for _b, _r in _ban:
                    if _b in _p: warns.append('W13 %s:%d 用了手册没有的说法「%s」——改用「%s」（手册用语表）' % (os.path.basename(f)[:-3], _i, _b, _r))
    # S9 主题挂靠（2026-10-19 1019R4 起；2026-10-23 SD-233 改为「速查目录」）：完整版每条正文在速查目录里恰好作为直接项出现一次
    #   ——由 qr_directory.py 自检（速查源绑定 + EXTRA / 第五章规则），漏挂即报错
    import subprocess as _sp, tempfile as _tf
    _r = _sp.run([sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'qr_directory.py'), '--repo', os.path.dirname(os.path.abspath(SRC.rstrip('/'))),
                  '--out', os.path.join(_tf.mkdtemp(), 'd.json')], capture_output=True, text=True)
    if _r.returncode: errors.append('S9 ' + ((_r.stderr or _r.stdout).strip().splitlines() or ['速查目录生成失败'])[-1])
    # S10 引用的图都已入库（2026-10-21：figures/nd/ 10 张 ND 符号图被 .gitignore 的 *.png 挡掉，bundle 与 GitHub 备份缺图，换机 / 噜噜重建会失败；10-14 同类问题第二次出现）
    _root = os.path.dirname(os.path.abspath(SRC.rstrip('/')))
    try:
        _trk = set(subprocess.run(['git', '-C', _root, 'ls-files', '-z'], capture_output=True, text=True).stdout.split('\0'))
    except Exception:
        _trk = None
    if _trk and len(_trk) > 1:
        _refs = set()
        for f in files + [os.path.join(_root, '速查', '速查源.md')]:
            if os.path.exists(f):
                _refs |= {'notes_src/' + m for m in re.findall(r'figures/[^\s|)\]"\'<>]+\.(?:png|jpe?g|svg)', io.open(f, encoding='utf-8').read())}
        _un = sorted(r for r in _refs if r not in _trk)
        if _un: errors.append('S10 %d 张引用的图没有入库（git add -f，并检查 .gitignore 是否挡掉了子目录）：%s' % (len(_un), '、'.join(_un[:10])))
    # S11 公司名不进成品（2026-10-22 用户：「整个笔记中不要出现××等信息，不同标准可用遵照公司实际要求」（××为公司简称；名单只放私有库《成品禁用词.json》，公开库不写））：
    # 航空公司名称 / 简称、航班号、无线电呼号写成「公司」「合格证持有人」「（航班呼号）」或「遵照公司实际要求」。
    # 2026-10-24：词表（正则）只放私有库根目录《成品禁用词.json》——公开工具链不写公司名（推送泄漏扫描）；词表缺失或读不了报错，不悄悄放过
    _ban = []
    _banf = os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '成品禁用词.json')
    try: _ban = json.load(io.open(_banf, encoding='utf-8'))
    except (OSError, ValueError) as _e: errors.append('S11 读不到私有库《成品禁用词.json》（%s）——公司名检查没做，不算通过' % _e)
    _hit = []
    for f in files + [os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '速查', '速查源.md')]:
        if not os.path.exists(f): continue
        for i, l in enumerate(io.open(f, encoding='utf-8'), 1):
            for p in _ban:
                if re.search(p, l): _hit.append('%s:%d「%s」' % (os.path.basename(f), i, re.search(p, l).group(0)))
    if _hit: errors.append('S11 成品里出现公司名 / 航班号 %d 处（写「公司」「合格证持有人」「（航班呼号）」或「遵照公司实际要求」）：%s' % (len(_hit), '、'.join(_hit[:8])))
    # 约定保留（私有库根目录《两型混写保留.json》，2026-10-09）：按「文件 + 去标签前 24 字」匹配，只计数不列为待处理
    keepf = os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '两型混写保留.json')
    keep = set()
    if os.path.exists(keepf):
        keep = {(x['文件'], x['片段']) for x in json.load(io.open(keepf, encoding='utf-8')).get('条目', [])}
    kept = [x for x in MIX if (x.split(':')[0], x[x.find('「') + 1:-1]) in keep]
    MIX[:] = [x for x in MIX if x not in kept]
    if MIX:
        warns.append('两型混写 %d 处（SD-170：两型共有的知识点改成对比表；SRC_MIX=1 列出明细）' % len(MIX))
        if os.environ.get('SRC_MIX'):
            warns.extend('  两型混写 ' + x for x in MIX)
    if kept:
        print('  两型混写约定保留 %d 处（两型混写保留.json）' % len(kept))
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
