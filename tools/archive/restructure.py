#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""restructure.py —— 表格与长文字的协调（SD-27），只移动位置、不改文字

用法：python3 tools/restructure.py <计划.json> [--src notes_src] [--write]
计划每项：[节号, 条目号, 动作, 表序号(默认 0)]
  note_out     表末通栏注释行 → 表后段落（按 <br> 分段）
  premise_out  表首通栏前提行 → 表前段落
  table_paras  两列「标签｜内容」表 → 「粗体标签」+ 段落（推理 / 举例类整表）
  split_rows   两列「类型｜处置」表里按粗体引导词（「现象：」「处置：」「若不能恢复：」…）拆成子行，
               变成「类型｜环节｜内容」三列；没有引导词的行保持一格
写入前自检：去掉标签与空白后，改前改后的字符序列（按段落 / 格子切分后排序）完全相同。
"""
import io, os, re, sys, glob, json

def _arg(flag, default):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default
SRC = os.path.abspath(_arg('--src', 'notes_src'))
plan = json.load(io.open(sys.argv[1], encoding='utf-8'))

def lines_of(inner):
    return [s.strip() for s in re.split(r'<br\s*/?>', inner.strip()) if s.strip()]

def signature(s):
    s = re.sub(r'(?s)<tr class="hdr">.*?</tr>', '', s)       # 表头是结构，不算正文
    s = re.sub(r'<[^>]+>', '', s).replace('：', '')
    return sorted(re.sub(r'\s', '', s))

def act(tb, how):
    rows = list(re.finditer(r'(?s)<tr([^>]*)>(.*?)</tr>\n?', tb))
    if how == 'note_out':
        r = [x for x in rows if 'class="note"' in x.group(1)][-1]
        inner = re.search(r'(?s)<td[^>]*>(.*)</td>', r.group(2)).group(1)
        return '', tb[:r.start()] + tb[r.end():], '\n\n'.join(lines_of(inner))
    if how == 'premise_out':
        r = [x for x in rows if 'class="premise"' in x.group(1)][0]
        inner = re.search(r'(?s)<td[^>]*>(.*)</td>', r.group(2)).group(1)
        return '\n\n'.join(lines_of(inner)), tb[:r.start()] + tb[r.end():], ''
    if how == 'table_paras':
        out = []
        for r in rows:
            if 'hdr' in r.group(1): continue
            cells = re.findall(r'(?s)<td([^>]*)>(.*?)</td>', r.group(2))
            if len(cells) == 2:
                out.append('<strong>%s</strong>' % re.sub(r'<br\s*/?>', ' ', re.sub(r'</?strong>', '', cells[0][1])).strip())
                cells = cells[1:]
            for _, inner in cells:
                out += lines_of(inner)
        hdr = [x for x in rows if 'hdr' in x.group(1)]
        return '', '', '\n\n'.join(out)
    if how == 'split_rows':
        out = []
        for r in rows:
            if 'hdr' in r.group(1):
                ths = re.findall(r'(?s)<th[^>]*>.*?</th>', r.group(2))
                out.append('<tr%s>%s<th>环节</th>%s</tr>' % (r.group(1), ths[0], ''.join(ths[1:]))); continue
            cells = re.findall(r'(?s)<td([^>]*)>(.*?)</td>', r.group(2))
            if len(cells) != 2 or re.search(r'premise|note|warn', r.group(1)):
                out.append(r.group(0).rstrip('\n').replace('colspan="2"', 'colspan="3"')); continue
            groups = []
            for seg in lines_of(cells[1][1]):
                m = re.match(r'<strong>([^<]{1,24})：</strong>\s*(.*)$', seg, re.S)
                if m:
                    groups.append([m.group(1), [m.group(2)] if m.group(2).strip() else []])
                elif groups: groups[-1][1].append(seg)
                else: groups.append(['', [seg]])
            if len(groups) <= 1:
                out.append('<tr><td>%s</td><td colspan="2">%s</td></tr>' % (cells[0][1], cells[1][1])); continue
            for gi, (lab, segs) in enumerate(groups):
                head = '<td rowspan="%d">%s</td>' % (len(groups), cells[0][1]) if gi == 0 else ''
                out.append('<tr>%s<td>%s</td><td>%s</td></tr>' % (head, lab, '<br>'.join(segs)))
        tbl_open = re.match(r'<table[^>]*>', tb).group(0)
        return '', tbl_open + '\n' + '\n'.join(out) + '\n</table>', ''
    raise SystemExit('未知动作 ' + how)

files = {}
for sec, item, how, *rest in plan:
    idx = rest[0] if rest else 0
    f = [x for x in glob.glob(os.path.join(SRC, '[1-5]*', '*.md')) if os.path.basename(x).startswith(sec + ' ')][0]
    t = files.get(f) or io.open(f, encoding='utf-8').read()
    m = re.search(r'(?ms)^#{3,4} ' + re.escape(item) + r'　[^\n]*\n(.*?)(?=^#{3,4} |\Z)', t)
    body = m.group(1)
    tbm = list(re.finditer(r'(?s)<table\b.*?</table>', body))[idx]
    before, newtb, after = act(tbm.group(0), how)
    if how == 'table_paras':
        rep = after
        # 表头文字不入正文：检查时把表头从原文里去掉
        orig = re.sub(r'(?s)<tr class="hdr">.*?</tr>', '', tbm.group(0))
    else:
        rep = (before + '\n\n' if before else '') + newtb + ('\n\n' + after if after else '')
        orig = tbm.group(0)
    assert signature(orig) == signature(rep), '文字发生变化：%s %s %s' % (sec, item, how)
    body2 = body[:tbm.start()] + rep + body[tbm.end():]
    t = t[:m.start(1)] + body2 + t[m.end(1):]
    files[f] = t
    print('OK', sec, item, how)
if '--write' in sys.argv:
    for f, t in files.items(): io.open(f, 'w', encoding='utf-8').write(t)
    print('已写入 %d 个文件' % len(files))
