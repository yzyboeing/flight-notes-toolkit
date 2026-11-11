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
from html import unescape

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


def find_type_mixtures(text):
    """同格或同段显式混写，两种标记必须完整出现；返回准确原始偏移。"""
    masked = list(text)
    spans = []
    def hide(start, end):
        for i in range(start, end):
            if masked[i] != '\n': masked[i] = ' '
    for m in re.finditer(r'<!--.*?-->', text, re.S): hide(m.start(), m.end())
    clean = ''.join(masked)
    for m in re.finditer(r'<t[dh]\b[^>]*>.*?</t[dh]>', clean, re.S | re.I):
        spans.append(('cell', m.start(), m.end()))
    for m in re.finditer(r'<table\b.*?</table>', clean, re.S | re.I): hide(m.start(), m.end())
    outside = ''.join(masked)
    for m in re.finditer(r'(?m)(?:^[^\S\n]*\S[^\n]*(?:\n(?![ \t]*\n)[^\n]*)*)', outside):
        if m.group().lstrip().startswith(('#', '%%FIG')): continue
        spans.append(('paragraph', m.start(), m.end()))
    out = []
    for kind, start, end in sorted(spans, key=lambda x: x[1]):
        block = clean[start:end]
        if '【737-NG】' in block and '【737-8】' in block:
            plain = re.sub(r'<[^>]+>|\s+', '', block)
            out.append({'kind': kind, 'offset': start, 'line': text.count('\n', 0, start) + 1, 'fragment': plain[:24]})
    return out

NOTE_LINE = re.compile(r'^[ \t]*注[：:](.*)$', re.M)
NOTE_META = re.compile(r'中文\s*(?:QRH|FCOM|FCTM|AFM).{0,12}(?:译作|译为|措辞|写作)|'
                       r'(?:属|仅为|只是)(?:中英文)?翻译差异|本表以.{0,24}为准|本机队|'
                       r'两型.{0,12}(?:只是|仅)措辞不同')


def check_note_lines(body, filename, offset=0):
    notes = [[], [], []]
    headings = list(re.finditer(r'^### (.+)$', body, re.M))
    for index, heading in enumerate(headings):
        end = headings[index + 1].start() if index + 1 < len(headings) else len(body)
        start = heading.end()
        matches = list(NOTE_LINE.finditer(body, start, end))
        line = offset + body.count('\n', 0, heading.start()) + 1
        if len(matches) >= 4:
            notes[0].append(f'{filename}:{line}「{heading[1]}」{len(matches)} 行')
        for match in matches:
            line = offset + body.count('\n', 0, match.start()) + 1
            text = re.sub(r'<[^>]*>', '', match[1]).strip()
            if not text:
                notes[1].append(f'{filename}:{line}')
            if NOTE_META.search(text):
                notes[2].append(f'{filename}:{line}「{NOTE_META.search(text)[0]}」')
    return notes


# SD-239 / SD-240: old labels may appear at the start of a line or inside a table cell.
RETIRED_NOTE = re.compile(r'(?:公司差异|解释|出处|来源)\s*[：:]')
NAMED_MANUAL = re.compile(
    r'FCOM\s*(?:限制|用(?!于))|大纲\s*用(?!于)|QRH\s*[：:]|'
    r'(?:运行手册|SOP|FCOM|FCTM|QRH|AFM|大纲)\s*(?:的\s*)?口径|'
    r'公司手册|(?:运行手册|SOP|FCOM|FCTM|QRH|AFM)\s*(?:只写|仅写|写作|另写|写为|只要求|写(?=[「“])|为[「“])'
)
ABBR_NOTE = re.compile(r'^\s*注[：:]\s*缩写[—–-]+')
ABBR_DEFINITION = re.compile(r'(?:^|[；;])\s*([A-Z][A-Z0-9 /.-]*?)\s*[：:]')


def common_abbreviations():
    # Read the maintained W06 list to keep the checker and the rule in sync.
    from pathlib import Path
    rules = Path(__file__).resolve().parents[1] / 'prompt/layout-checklist.md'
    w06 = next(line for line in rules.read_text(encoding='utf-8').splitlines() if '| **W06** |' in line)
    match = re.search(r'常用缩写（([^（]+)（SD-250））', w06)
    return set(match[1].removesuffix(' 等').split('、'))


def scan_editorial_markers(source, filename, common, offset=0):
    retired, named, glossary = [], [], []
    # Mask comments, keeping line numbers stable.
    source = re.sub(r'<!--.*?-->', lambda m: '\n' * m[0].count('\n'), source, flags=re.S)
    for line, raw in enumerate(source.splitlines(), offset + 1):
        text = re.sub(r'<[^>]*>', '', raw).strip()
        where = f'{filename}:{line}'
        for piece in re.split(r'<br\s*/?>|<t[dh]\b[^>]*>|[。；;]', raw, flags=re.I):
            visible = re.sub(r'<[^>]*>', '', piece).strip().lstrip('#* ').strip()
            if hit := RETIRED_NOTE.match(visible):
                retired.append(f'{where}「{hit[0]}」')
        for hit in NAMED_MANUAL.finditer(text):
            named.append(f'{where}「{hit[0]}」')
        match = ABBR_NOTE.match(text)
        if match:
            for item in ABBR_DEFINITION.findall(text[match.end():]):
                if item in common:
                    glossary.append(f'{where}「{item}」')
    return retired, named, glossary


# S21: a cell shared by the two model columns must contain only common material.
MODEL_ONLY = re.compile(
    r'【\s*737-(?:NG|8)(?=[\s】])[^】]*】|'
    r'部分\s*737-(?:NG|8)\b|\b737-(?:NG|8)\s*及', re.I
)
MODEL_ATTRIBUTE = re.compile(r'data-aircraft=["\'](737-(?:NG|8))["\']', re.I)


def scan_model_colspans(source, filename, offset=0):
    def span(attrs, name):
        match = re.search(rf'\b{name}=["\'](\d+)["\']', attrs)
        return int(match[1]) if match else 1

    hits = []
    for table in re.finditer(r'<table\b[^>]*>.*?</table>', source, re.S):
        if source.rfind('<!--', 0, table.start()) > source.rfind('-->', 0, table.start()):
            continue
        occupied = set()
        model_columns, merged = {}, []
        for index, row in enumerate(re.finditer(r'<tr\b[^>]*>.*?</tr>', table[0], re.S)):
            column = 0
            for cell in CELL.finditer(row[0]):
                while (index, column) in occupied:
                    column += 1
                tag, attrs, body = cell.groups()
                width, height = span(attrs, 'colspan'), span(attrs, 'rowspan')
                columns = range(column, column + width)
                occupied.update((index + step, col) for step in range(1, height) for col in columns)
                visible = re.sub(r'<[^>]+>', '', body)
                heading = re.sub(r'[【】\s]', '', visible).upper()
                if tag == 'th' and heading in ('737-NG', '737-8'):
                    model_columns[heading] = columns
                if width > 1:
                    location = table.start() + row.start() + cell.start()
                    merged.append((columns, attrs, visible, location))
                column += width
        if len(model_columns) != 2:
            continue
        for columns, attrs, visible, location in merged:
            if not all(set(columns).intersection(model) for model in model_columns.values()):
                continue
            marker = MODEL_ONLY.search(visible) or MODEL_ATTRIBUTE.search(attrs)
            if marker:
                line = offset + source.count('\n', 0, location) + 1
                hits.append(f'{filename}:{line}「{marker[0]}」')
    return hits


_CELL_CLASS = re.compile(r"\bclass\s*=\s*([\"'])(.*?)\1", re.S)
_CELL_LINE_BREAK = re.compile(r'<br\s*/?>|</(?:p|li)>', re.I)


def _alignment_segments(body):
    text = re.sub(r'<[^>]+>', '', _CELL_LINE_BREAK.sub('\n', body.replace('\n', ' ')))
    return [re.sub(r'[\s\u200b\u2060\u00ad]+', '', unescape(part))
            for part in text.split('\n') if re.sub(r'\s+', '', unescape(part))]


def _alignment_length(text):
    # 与 App 的短格口径一致，中文计一字宽，ASCII 计半字宽。
    return sum(0.5 if ord(ch) < 256 else 1 for ch in text)


_MODEL_HEADER = re.compile(r'737|两型|仅', re.I)
_PIN_MARKS = {'居中': 'col-center', '靠左加点': 'col-bullet'}
_PIN_BOOKS = {'两本': ('完整版', '速查'), '完整版': ('完整版',), '速查': ('速查',)}


def alignment_tables(source):
    """Yield source table positions with their logical column spans."""
    masked = re.sub(r'<!--.*?-->', lambda m: re.sub(r'[^\n]', ' ', m[0]), source, flags=re.S)
    for table in TABLE.finditer(masked):
        occupied, cells = set(), []
        for row_no, row in enumerate(TR.finditer(table[0])):
            col = 0
            for cell in CELL.finditer(row[2]):
                while (row_no, col) in occupied:
                    col += 1
                tag, attrs, body = cell.groups()
                colspan = int((re.search(r"""\bcolspan\s*=\s*["'](\d+)["']""", attrs) or [None, 1])[1])
                rowspan = int((re.search(r"""\browspan\s*=\s*["'](\d+)["']""", attrs) or [None, 1])[1])
                occupied.update((row_no + dy, col + dx) for dy in range(rowspan) for dx in range(colspan))
                classes = _CELL_CLASS.search(attrs)
                attr_start = table.start() + row.start(2) + cell.start(2)
                cells.append({
                    'row': row_no, 'col': col, 'span': colspan, 'tag': tag,
                    'segments': _alignment_segments(body), 'classes': classes[2].split() if classes else [],
                    'offset': attr_start + classes.start() if classes else attr_start + len(attrs),
                })
                col += colspan
        header_row = next((cell['row'] for cell in cells if cell['tag'] == 'th'), None)
        if header_row is not None:
            headers = [cell for cell in cells if cell['row'] == header_row and cell['tag'] == 'th']
            first_row = next((cell['row'] for cell in cells if cell['row'] > header_row and cell['tag'] == 'td'), None)
            yield {'headers': headers, 'cells': cells, 'offset': table.start(),
                   'first': [''.join(c['segments']) for c in cells if c['row'] == first_row]}


def find_alignment_pins(documents, ledger):
    """Resolve each active pin in each specified book and report unmatched or ambiguous pins."""
    tables = collections.defaultdict(list)
    for book, filename, source, offset, table in documents:
        headers = tuple(''.join(cell['segments']) for cell in table['headers'])
        tables[book].append((headers, filename, source, offset, table))
    found, problems = [], []
    for pin in ledger:
        if pin.get('停用'):
            continue
        headers = tuple(re.sub(r'\s+', '', label) for label in pin['表头'])
        label = re.sub(r'\s+', '', pin['列'])
        first = re.sub(r'\s+', '', pin.get('首行', ''))
        for book in _PIN_BOOKS[pin['书']]:
            candidates = [entry for entry in tables[book]
                          if not first or any(first in value for value in entry[4]['first'])]
            matches = [entry for entry in candidates if entry[0] == headers]
            if not matches:
                # A table may acquire a trailing column while keeping the pinned headers.
                extended = [entry for entry in candidates if entry[0][:len(headers)] == headers]
                if len(extended) == 1:
                    matches = extended
            if len(matches) != 1:
                problems.append(f"{book}「{'｜'.join(headers)}」/「{label}」匹配 {len(matches)} 张表")
                continue
            _, filename, source, offset, table = matches[0]
            index = headers.index(label)
            cell = table['headers'][index]
            found.append((filename, source, offset, table, cell, pin))
    return found, problems


def scan_center_alignment(table, exempt):
    for header in table['headers']:
        label = ''.join(header['segments'])
        if (header['offset'] in exempt or 'col-center' not in header['classes']
                or not _MODEL_HEADER.search(label)):
            continue
        if any(_alignment_length(segment) >= 15 for cell in table['cells']
               if cell['row'] > header['row'] and cell['tag'] == 'td' and cell['span'] == 1
               and header['col'] <= cell['col'] < header['col'] + header['span']
               for segment in cell['segments']):
            yield header


MIX_DETAIL = []

MIX = []
def main():
    files = [f for f in glob.glob(os.path.join(SRC, '**', '*.md'), recursive=True)
             if not any(part.startswith('.') for part in os.path.relpath(f, SRC).split(os.sep))]
    if not files:
        sys.exit('未在 %s 下找到任何 .md' % SRC)

    names = {os.path.splitext(os.path.basename(f))[0] for f in files}
    ids   = collections.defaultdict(list)
    notes = 0
    note_hints = [[], [], []]
    common = common_abbreviations()
    editorial = [[], [], []]
    alignment_documents = []

    def collect_editorial(text, filename, offset=0):
        for bucket, hits in zip(editorial, scan_editorial_markers(text, filename, common, offset)):
            bucket.extend(hits)
        errors.extend('K-S21 两型共有格含单型内容 '+hit
                      for hit in scan_model_colspans(text, filename, offset))
        book = '速查' if filename == '速查/速查源.md' else '完整版'
        alignment_documents.extend((book, filename, text, offset, table) for table in alignment_tables(text))


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

        body_start = m.end()
        body = t[body_start:]
        collect_editorial(body, rel(f), t[:body_start].count("\n"))
        if is_note:
            for bucket, hits in zip(note_hints, check_note_lines(body, rel(f), t[:body_start].count('\n'))):
                bucket.extend(hits)

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

        # SD-170：提示级；同格/完整多行段落，保留准确全文行号供私有进度表。
        for hit in find_type_mixtures(body):
            line = t[:body_start].count('\n') + hit['line']
            previous = list(re.finditer(r'(?m)^### (\d+)\.', body[:hit['offset']]))
            key = sid + ' ' + previous[-1].group(1) if previous and sid else ''
            MIX_DETAIL.append({**hit, 'file': rel(f), 'section': sid or '', 'key': key, 'line': line})
            MIX.append('%s:%d「%s」' % (rel(f), line, hit['fragment']))

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
        collect_editorial(qt, '速查/速查源.md')
        for bucket, hits in zip(note_hints, check_note_lines(qt, '速查/速查源.md')):
            bucket.extend(hits)
        for mh in re.finditer(r'<table\b[^>]*>\s*(?:<tr class="premise">.*?</tr>\s*)?<tr class="hdr">\s*<th[^>]*>(.*?)</th>', qt, re.S):
            if re.match(r'\s*(?:<[^>]+>)*\s*737', mh.group(1)):
                errors.append('对比表缺项目列（表头首格是机型名）速查/速查源.md:%d' % (qt[:mh.start()].count('\n') + 1))
        for seg in re.findall(r'<td\b[^>]*>.*?</td>', qt, re.S) + [l for l in re.sub(r'<table\b.*?</table>', '', qt, flags=re.S).split('\n') if not l.startswith(('<!--', '#'))]:
            if '【737-NG' in seg and '【737-8' in seg:
                MIX.append('速查/速查源.md:%d「%s」' % (qt[:qt.find(seg)].count('\n') + 1, re.sub(r'<[^>]+>|\s+', '', seg)[:24]))
    ledger_path = os.path.join(os.path.dirname(SRC), '用户点名排版台账.json')
    ledger = []
    if os.path.isfile(ledger_path):
        with open(ledger_path, encoding='utf-8') as ledger_file:
            ledger = json.load(ledger_file)
    pins, pin_problems = find_alignment_pins(alignment_documents, ledger)
    errors.extend('K-USERPIN ' + problem for problem in pin_problems)
    exempt = collections.defaultdict(set)
    for filename, source, offset, _, cell, pin in pins:
        exempt[filename].add(cell['offset'])
        mark = _PIN_MARKS[pin['要求']]
        if mark not in cell['classes']:
            line = offset + source.count('\n', 0, cell['offset']) + 1
            errors.append(f"K-USERPIN {filename}:{line}「{pin['列']}」缺少 {mark}")
    for _, filename, source, offset, table in alignment_documents:
        for cell in scan_center_alignment(table, exempt[filename]):
            line = offset + source.count('\n', 0, cell['offset']) + 1
            errors.append(f"K-S22 {filename}:{line} 机型长句列标了 col-center")

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
        if n and (n < 5 or (n > 25 and not b.startswith('1.'))) and not b.startswith(('4.', '5.', '6.')): warns.append('S8 %s 共 %d 条（N2：一节 5～25 条；第四、五章按专题分节，不计）' % (b[:-3], n))
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
    for f in files + ([os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '速查', '速查源.md')] if os.environ.get('BUILD_QR') else []):   # 旧速查源只在出速查版时查（SD-240）
        if not os.path.exists(f): continue
        _fm = False
        for i, l in enumerate(io.open(f, encoding='utf-8'), 1):
            if i == 1 and l.strip() == '---': _fm = True; continue
            if _fm:   # 文件头只查进成品的字段（章首简介 blurb、节名 title_cn），source 等不进成品、不查
                _fm = l.strip() != '---'
                if not re.match(r'\s*(blurb|title_cn)\s*:', l): continue
            for p in _ban:
                if re.search(p, l): _hit.append('%s:%d「%s」' % (os.path.basename(f), i, re.search(p, l).group(0)))
    if _hit: errors.append('S11 成品里出现公司名 / 公司化描述 / 内部机构文号 %d 处（SD-240：公司规定写「部分航空公司规定」，各家不同写「按航空公司规定」，规章责任主体写「合格证持有人」，通话写「（航班呼号）」）：%s' % (len(_hit), '、'.join(_hit[:8])))
    # S12 不写具体机号（SD-247，2026-11-05 用户：「所有的笔记中也不要出现具体的机号」「不要写部分机号」）：
    # 注册号（B-1234 / B-206M）与「机号」一词都报错；notes_src 与速查源都查（速查源去机号是 SD-242 只读的一次性例外）
    _reg = re.compile(r'B-\d{3}[0-9A-Z]|机号')
    _rh = []
    for f in files + [os.path.join(os.path.dirname(os.path.abspath(SRC.rstrip('/'))), '速查', '速查源.md')]:
        if not os.path.exists(f): continue
        _fm = False
        for i, l in enumerate(io.open(f, encoding='utf-8'), 1):
            if i == 1 and l.strip() == '---': _fm = True; continue
            if _fm:
                _fm = l.strip() != '---'
                if not re.match(r'\s*(blurb|title_cn)\s*:', l): continue
            for m in _reg.finditer(l): _rh.append('%s:%d「%s」' % (os.path.basename(f), i, m.group(0)))
    if _rh: errors.append('S12 成品里出现具体机号或「机号」一词 %d 处（SD-247：写「部分 737-NG 飞机」「部分 737-8 飞机」，数值不同的几组都保留、只换标签）：%s' % (len(_rh), '、'.join(_rh[:8])))
    # S13 成品不写审核口吻（SD-248，2026-11-05 用户：「按你的建议做提速」）：校对理由写进卡片和台账，不进正文
    _voice = re.compile(r'不能写成|不套用|本次所核|通用许可|推定|不把持续|不要把这一分支|本卡|未获原页|原页没有')
    _vh = []
    for f in files:
        for i, l in enumerate(io.open(f, encoding='utf-8'), 1):
            for m in _voice.finditer(l): _vh.append('%s:%d「%s」' % (os.path.basename(f), i, m.group(0)))
    if _vh: errors.append('S13 成品里出现审核口吻 %d 处（SD-248：校对理由写进知识点卡和台账，正文只写手册内容）：%s' % (len(_vh), '、'.join(_vh[:8])))
    # S14 两型对比表逐字相同未合并（B13，SD-248）：表头含 737-NG 与 737-8 的表，同一行两列内容逐字相同（「—」除外）应写成跨两列的共有格
    _ih = []
    for f in files:
        _s = io.open(f, encoding='utf-8').read()
        for _tm in re.finditer(r'<table[^>]*>(.*?)</table>', _s, re.S):
            _ng = _e8 = None; _grid = set()
            for _ri, (_attr, _r) in enumerate(re.findall(r'<tr([^>]*)>(.*?)</tr>', _tm.group(1), re.S)):
                _cells = re.findall(r'<t[hd]([^>]*)>(.*?)</t[hd]>', _r, re.S)
                if 'hdr' in _attr:
                    _nm = [re.sub(r'<[^>]+>|[【】\s]', '', c[1]) for c in _cells]
                    if '737-NG' in _nm and '737-8' in _nm: _ng, _e8 = _nm.index('737-NG'), _nm.index('737-8')
                    continue
                if _ng is None: continue
                _col = 0; _pos = {}
                for _a, _t in _cells:
                    while (_ri, _col) in _grid: _col += 1
                    _cs = int((re.search(r'colspan="(\d+)"', _a) or [0, 1])[1]); _rs = int((re.search(r'rowspan="(\d+)"', _a) or [0, 1])[1])
                    for _k in range(1, _rs):
                        for _c in range(_col, _col + _cs): _grid.add((_ri + _k, _c))
                    if _cs == 1: _pos[_col] = _t
                    _col += _cs
                _x, _y = _pos.get(_ng), _pos.get(_e8)
                if _x is not None and _y is not None and _x.strip() and _x.strip() not in ('—', '-', '–') and _x.strip() == _y.strip():
                    _ih.append('%s:%d「%s」' % (os.path.basename(f), _s[:_tm.start()].count('\n') + 1, re.sub(r'<[^>]+>', '', _x)[:20]))
    if _ih: errors.append('S14 两型对比表逐字相同未合并 %d 处（B13 / SD-248：写成 colspan="2" data-aircraft="both" 的共有格）：%s' % (len(_ih), '、'.join(_ih[:8])))
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
    if '--mix-report' in sys.argv:
        target = _arg('--mix-report', '')
        counts = collections.Counter(x['section'] for x in MIX_DETAIL)
        data = {'scope': 'current notes_src only', 'level': 'warning',
                'count': len(MIX_DETAIL), 'sections': dict(sorted(counts.items())),
                'locations': [{**x, 'agreedKeep': (x['file'], x['fragment']) in keep} for x in MIX_DETAIL]}
        with io.open(target, 'w', encoding='utf-8') as h:
            json.dump(data, h, ensure_ascii=False, indent=2)
            h.write('\n')
    # S18 applies to retired labels regardless of their source container.
    errors.extend('K-S18 停用标签 '+hit for hit in editorial[0])
    for label, hits in zip(('S19 点名手册', 'S20 W06 常用缩写注'), editorial[1:]):
        for hit in hits:
            print(f'  提示  K-{label} {hit}')
    # K-S21（N7，用户 2026-11-09）：第五章故障条目按「现象 → 分析 → 处置 → 注意事项」四段写——缺「现象：」导语表或「注：注意事项——」的报提示（进度跟踪用）
    s21 = []
    # 判断方法、指示判读、操作技术、数值由来类条目不是单一故障，不套四段（顾问 11-09 裁定）
    S21_EXEMPT = {'5.3.1', '5.3.2', '5.3.3', '5.3.4', '5.3.7', '5.3.8', '5.3.9', '5.3.10', '5.3.11', '5.3.15',
                  '5.5.2', '5.5.3', '5.5.4', '5.5.5',   # 5.5.2～5 是单发按飞行阶段的处置流程
                  '5.7.3', '5.7.7',   # 限坡度进近技术、非正常形态进近速度汇总（1110 第四轮 O2 调序后 5.7.6 → 5.7.7）
                  '5.9.5',   # 液压失效识别与决策（现象与判断已在 5.9.6）
                  '5.14.3', '5.14.4'}   # 水上迫降检查单的后续阶段（现象与分析在 5.14.2）
    for f in sorted(glob.glob(os.path.join(SRC, '5 故障处置', '5.*.md'))):
        sec = os.path.basename(f).split(' ')[0]
        if sec not in ('5.3', '5.4', '5.5', '5.7', '5.9', '5.14'):
            continue
        parts = re.split(r'\n### (\d+)\. ', io.open(f, encoding='utf-8').read())
        for k in range(1, len(parts), 2):
            body = parts[k + 1]
            summ = '现象与处置：' in body   # 汇总条目：一行一张检查单，不要求处置思路
            if re.search(r'(?m)^判断与决策参考：\s*$', body):   # 第五轮（11-10）新骨架：判断与决策参考 → 背景知识（可无）→ 结构说明（可无）→ 检查单与重要项目说明 → 相关故障
                miss = [n for n, ok in (('检查单与重要项目说明', re.search(r'(?m)^检查单与重要项目说明：\s*$', body)), ('相关故障', '注：相关故障——' in body)) if not ok]
                if miss and '%s.%s' % (sec, parts[k]) not in S21_EXEMPT:
                    s21.append('%s.%s（新骨架）缺%s' % (sec, parts[k], '、'.join(miss)))
                continue
            miss = [n for n, ok in (('现象', re.search(r'(?m)^现象(与处置)?：\s*$', body)),
                                    ('处置思路', summ or re.search(r'(?m)^处置思路：\s*$', body)),
                                    ('注意事项', '注：注意事项——' in body),
                                    ('相关故障', '注：相关故障——' in body)) if not ok]
            if miss and '%s.%s' % (sec, parts[k]) not in S21_EXEMPT:
                s21.append('%s.%s 缺%s' % (sec, parts[k], '、'.join(miss)))
    if s21 or not QUIET:
        print(f'  提示  K-S21 故障条目五段结构 {len(s21)} 条未完成' + ('：' + '、'.join(s21[:8]) + ('…' if len(s21) > 8 else '') if s21 else ''))
    # K-S22（顾问 11-09：差异真源在知识点处、第六章作对照视图）：第五章「注：两型检查单差异——」每一段都要能在 6.2 正文里找到
    s22 = []
    p62 = glob.glob(os.path.join(SRC, '6 机型差异', '6.2 *.md'))
    if p62:
        plain = lambda s: re.sub(r'\s+', '', re.sub(r'<[^>]+>', '', s.replace('<br>', '；'))).replace('&lt;', '<').replace('&gt;', '>')
        T62 = plain(io.open(p62[0], encoding='utf-8').read())
        for f in sorted(glob.glob(os.path.join(SRC, '5 故障处置', '5.*.md'))):
            _t5 = io.open(f, encoding='utf-8').read()
            # 第五轮（11-10）：差异由「注：两型检查单差异——」大段注改为「两型检查单差异：」导语 + 表格；取表格里除首列（检查单 / 差异点）、机型列外的各格
            for m in re.finditer(r'(?m)^两型检查单差异：\n\n((?:<table.*?</table>\n\n?)+)', _t5, re.S):
                for row in re.findall(r'(?s)<tr><td>.*?</tr>', m.group(1)):
                    cells = re.findall(r'(?s)<td[^>]*>(.*?)</td>', row)[1:]
                    cells = [c for c in cells if c not in ('737-8', '737-NG')]
                    for seg in re.split(r'<br>|：|；|。', '；'.join(cells)):
                        s = plain(seg).lstrip('-')
                        if len(s) >= 6 and s != '无此项' and s not in T62:
                            s22.append('%s「%s」' % (os.path.basename(f).split(' ')[0], seg.strip()[:24]))
            if '注：两型检查单差异——' in _t5:
                s22.append('%s 仍有旧式「注：两型检查单差异——」大段注（应改为表格）' % os.path.basename(f).split(' ')[0])
    # K-S23：第一章「〈本节系统名〉两型差异」条目里的表，须与 6.1 对应表逐字一致（6.1 为对照视图）
    s23 = []
    p61 = glob.glob(os.path.join(SRC, '6 机型差异', '6.1 *.md'))
    if p61:
        norm = lambda s: re.sub(r'\s+', '', s)
        T61 = {norm(x) for x in re.findall(r'(?s)<table.*?</table>', io.open(p61[0], encoding='utf-8').read())}
        for f in sorted(glob.glob(os.path.join(SRC, '1 系统描述', '1.*.md'))):
            txt = io.open(f, encoding='utf-8').read(); sysn = (re.search(r'(?m)^title_cn:\s*"([^"]+)"', txt) or [None, '\x00'])[1]
            for m in re.finditer(r'(?ms)^### (\d+)\. %s ?两型差异\n(.*?)(?=^### |\Z)' % re.escape(sysn), txt):
                for tb in re.findall(r'(?s)<table.*?</table>', m.group(2)):
                    if norm(tb) not in T61:
                        s23.append('%s.%s' % (os.path.basename(f).split(' ')[0], m.group(1)))
    # 规则号 S24 S25。K-S24（第四轮 F0）：同一节内插图编号「图 x.y-N」须按出现顺序 1、2、3…（导图追加后忘了重编号会乱序）
    # K-S25（第四轮第一批）：正文不得有连续重复的中文标点「；；」「，，」「。。」（转换脚本拼接时最易产生）
    s24, s25 = [], []
    for f in sorted(glob.glob(os.path.join(SRC, '[1-6] *', '[1-6].*.md'))):
        sec = os.path.basename(f).split(' ')[0]; nums = []
        for ln, line in enumerate(io.open(f, encoding='utf-8').read().split('\n'), 1):
            m = re.match(r'%%FIG[A-Z]*%%[^|]*\|\s*图\s*' + re.escape(sec) + r'-(\d+)', line)
            if m: nums.append(int(m.group(1)))
            if not line.startswith('%%') and re.search(r'；；|，，|。。', line):
                s25.append('%s:%d' % (os.path.basename(f), ln))
        if nums and nums != list(range(1, len(nums) + 1)):
            s24.append('%s 图号顺序 %s' % (sec, '、'.join(map(str, nums))))
    # 规则号 S26。K-S26（第五轮，用户 11-10：「这种太复杂了，看一眼就不想看下去了」）：注的任一行（<br> 分隔）≥ 200 字且含 2 段以上并列（顶层「；」）须分行；第五章两型检查单差异须用表格（K-S22 同时把关）
    s26 = []
    def _top_semis(s):
        d = 0; n = 0
        for ch in s:
            if ch in '（(': d += 1
            elif ch in '）)': d -= 1
            elif ch == '；' and d == 0: n += 1
        return n
    for f in sorted(glob.glob(os.path.join(SRC, '[1-6] *', '[1-6].*.md'))):
        for ln, line in enumerate(io.open(f, encoding='utf-8').read().split('\n'), 1):
            if not line.startswith('注：'): continue
            for seg in line.split('<br>'):
                ps = re.sub(r'<[^>]+>', '', seg)
                if len(ps) >= 200 and _top_semis(ps) >= 1:
                    s26.append('%s:%d（%d 字）' % (os.path.basename(f), ln, len(ps)))
    errors.extend('K-S26 长注未分行（≥ 200 字的并列内容按「；」逐条换行）：' + x for x in s26)
    # 规则号 S27。K-S27（第五轮，用户 11-10「知识点用超链接引到所在位置，减少重复」）：行内知识点链接 <a href="x.y 第 N 条">文字</a> 的目标条目须存在、不指向本条、文字非空
    s27 = []; _items = set()
    for f in sorted(glob.glob(os.path.join(SRC, '[1-6] *', '[1-6].*.md'))):
        _sec = os.path.basename(f).split(' ')[0]
        for _n in re.findall(r'(?m)^### (\d+)\.', io.open(f, encoding='utf-8').read()): _items.add('%s 第 %s 条' % (_sec, _n))
    for f in sorted(glob.glob(os.path.join(SRC, '[1-6] *', '[1-6].*.md'))):
        _sec = os.path.basename(f).split(' ')[0]; _cur = None
        for ln, line in enumerate(io.open(f, encoding='utf-8').read().split('\n'), 1):
            _m = re.match(r'### (\d+)\.', line)
            if _m: _cur = '%s 第 %s 条' % (_sec, _m.group(1))
            for _h, _txt in re.findall(r'<a href="([^"]*)">(.*?)</a>', line):
                if _h not in _items: s27.append('%s:%d 链接目标不存在「%s」' % (os.path.basename(f), ln, _h))
                elif _h == _cur: s27.append('%s:%d 链接指向本条「%s」' % (os.path.basename(f), ln, _h))
                if not re.sub(r'<[^>]+>', '', _txt).strip(): s27.append('%s:%d 链接文字为空' % (os.path.basename(f), ln))
    errors.extend('K-S27 知识点链接：' + x for x in s27)
    # 规则号 S28。K-S28（第五轮，用户点名「下方附两型下降率表」）：<!-- copy:x.y 第 N 条 --> … <!-- /copy --> 里的表须与原条目第一张表逐字一致（只许照抄，改只改原条目）
    s28 = []; _ent = {}
    for f in sorted(glob.glob(os.path.join(SRC, '[1-6] *', '[1-6].*.md'))):
        _sec = os.path.basename(f).split(' ')[0]
        for _m in re.finditer(r'(?ms)^### (\d+)\.[^\n]*\n(.*?)(?=^### |\Z)', io.open(f, encoding='utf-8').read()):
            _tb = re.search(r'(?s)<table.*?</table>', _m.group(2)); _ent['%s 第 %s 条' % (_sec, _m.group(1))] = _tb.group(0) if _tb else None
    for f in sorted(glob.glob(os.path.join(SRC, '[1-6] *', '[1-6].*.md'))):
        for _m in re.finditer(r'(?s)<!-- copy:(\S+ 第 \d+ 条) -->\s*(.*?)\s*<!-- /copy -->', io.open(f, encoding='utf-8').read()):
            _src = _ent.get(_m.group(1))
            if _src is None: s28.append('%s 照抄来源「%s」不存在或没有表' % (os.path.basename(f), _m.group(1)))
            elif re.sub(r'\s+', '', _src) != re.sub(r'\s+', '', _m.group(2)): s28.append('%s 照抄「%s」的表与原表不一致' % (os.path.basename(f), _m.group(1)))
    errors.extend('K-S28 照抄表：' + x for x in s28)
    errors.extend('K-S24 插图编号未按出现顺序：' + x for x in s24)
    errors.extend('K-S25 连续重复标点：' + x for x in s25)
    if s23:
        warns.append('K-S23 第一章两型差异表与 6.1 不一致 %d 处：%s' % (len(s23), '、'.join(s23)))
    if s22:
        warns.append('K-S22 两型检查单差异注与 6.2 不一致 %d 处：%s' % (len(s22), '、'.join(s22[:6])))
    for label, hits in zip(('S15 同条注 ≥ 4', 'S16 空注', 'S17 元注关键词'), note_hints):
        if hits or not QUIET:
            print(f'  提示  K-{label} {len(hits)} 处' + ('：' + '、'.join(hits[:8]) if hits else ''))
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
