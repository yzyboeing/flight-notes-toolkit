#!/usr/bin/env python3
"""SD-119 / SD-120 排版规则审查（2026-10-03）：读 build_docx.js 的 CELL_DUMP（每格分条后的文字，含内部标记），
逐格按规则查圆点、父子层级、编号、混合标记；再扫 notes_src 查表后注释格式。只报告，不改文件。

用法：python3 audit_rules.py cells.jsonl notes_src 输出.md
标记：\\uE001「•」并列项  \\uE002「–」子项（\\uE002\\uE002 源文件「- 」子项）  \\uE003 引语  \\uE004 单句 / 续句（不加点、对齐）
      \\uE005 编号子项  \\uE006 粗体小标题段（不加点、悬挂）
"""
import json, re, sys, glob, os, collections

B, C, P, Q, N, L = '', '', '', '', '', ''
MK = B + C + P + Q + N + L
NUM = re.compile(r'^\s*([①-⑳]|\d{1,2}[.、)）])')
PH = re.compile(r'^[—－\-–/／无空×✕✓√?？…（）()\s]*$')


def show(t):
    return (t.replace(L, '[标]').replace(B, '• ').replace(C, '– ').replace(P, '[引]').replace(Q, '').replace(N, '')
            .replace('\n', ' ⏎ '))


def mk(x):
    m = re.match('^[' + MK + ']+', x)
    return m.group(0) if m else ''


def body(x):
    return x[len(mk(x)):].strip()


def is_child(x):
    m = mk(x)
    return m.startswith(C) or m == N


def is_parent(x):
    return bool(re.search(r'[：:]\s*$', body(x))) and len(body(x)) <= 80


def audit_cells(path):
    out = collections.defaultdict(list)
    seen = set()
    for ln in open(path, encoding='utf-8'):
        tb = json.loads(ln)
        key = json.dumps(tb, ensure_ascii=False)
        if key in seen:
            continue
        seen.add(key)
        hdr = ' | '.join(tb['hdr'])[:40]
        cols = collections.defaultdict(list)
        for r in tb['rows']:
            if re.search('hdr|note|premise|warn', r['cls']):
                continue
            first = next((c['t'] for c in r['cells'] if c['col'] == 0), '')[:14].replace('\n', ' ')
            where = '%s ／ %s' % (hdr, first)
            for c in r['cells']:
                if c['head'] or c['col'] == 0:
                    continue
                t = c['t']
                ls = [x for x in t.split('\n') if x.strip()]
                if not ls:
                    continue
                if c['span'] == 1:
                    cols[c['col']].append(t)
                tops = [x for x in ls if mk(x) == B]
                kids = [x for x in ls if is_child(x)]
                marked = any(mk(x) for x in ls)
                # R1 单条圆点：整格只有 1 个「•」且没有子项（SD-120：单句不加点）
                if len(tops) == 1 and not kids and len(ls) == 1:
                    out['R1 单句加了圆点'].append((where, show(t)[:70]))
                # R2 编号子项前带短线
                # （渲染时编号子项自动不加短线，见 build_docx.js prefix2，这里不再报）
                # R3 父项下只有一个子项（应合并）
                for i, x in enumerate(ls):
                    if not is_child(x) and is_parent(x) and i + 1 < len(ls) and is_child(ls[i + 1]) and not (i + 2 < len(ls) and is_child(ls[i + 2])):
                        out['R3 父项下只有一个子项'].append((where, show(x)[:40] + ' ⏎ ' + show(ls[i + 1])[:30]))
                # R4 子项没有父项
                for i, x in enumerate(ls):
                    if is_child(x):
                        j = i - 1
                        while j >= 0 and is_child(ls[j]):
                            j -= 1
                        if j < 0 or not is_parent(ls[j]):
                            out['R4 子项上面没有以「：」结尾的父项'].append((where, (show(ls[j])[:30] if j >= 0 else '（格首）') + ' ⏎ ' + show(x)[:30]))
                            break
                # R5 父项（以「：」结尾）后面紧跟的是「•」同级项，层级没分开
                for i, x in enumerate(ls[:-1]):
                    if mk(x) == B and is_parent(x) and mk(ls[i + 1]) == B:
                        out['R5 父项与子项同级（都是「•」）'].append((where, show(x)[:40] + ' ⏎ ' + show(ls[i + 1])[:30]))
                # R6 同格混合：有标记项又有无标记的普通行（「注：」行除外）
                if marked and len(ls) >= 2:
                    plain = [x for x in ls if not mk(x) and not re.match(r'^注[：:]', x.strip())]
                    if plain and (tops or kids):
                        out['R6 同格里有圆点项又有无标记行'].append((where, show(t)[:90]))
                # R7 无标记多行格：≥ 2 行、各 ≥ 6 字、不是编号 / 机型 / 注，却没有任何标记（可能漏分条，也可能是隐式父子层级）
                if not marked and len(ls) >= 2 and all(len(x) >= 10 for x in ls) and any(len(x) >= 20 for x in ls) and not any(NUM.match(x) or x.startswith(('【', '注', '（', '→', '但', '即', '且', '或')) for x in ls):
                    if any(is_parent(x) for x in ls[:-1]):
                        out['R7a 隐式父子层级（渲染时按「：」自动分层，核对层级是否正确）'].append((where, show(t)[:90]))
                    else:
                        out['R7b 多行并列却没加点'].append((where, show(t)[:90]))
                # R8 小标题段：同一格里一部分是粗体小标题段（[标]），一部分是普通圆点
                if any(mk(x) == L for x in ls) and tops:
                    out['R8 小标题段与圆点项混用'].append((where, show(t)[:90]))
                # R9 编号不连续
                nums = [re.match(r'^\s*([①-⑳])', body(x)) for x in ls]
                nums = [ord(m.group(1)) - 0x245F for m in nums if m]
                if len(nums) >= 2 and nums != list(range(nums[0], nums[0] + len(nums))):
                    out['R9 ①②编号不连续'].append((where, str(nums)))
        # 列级：同一列里，有的格有多条却没加点、有的格加了点
        for k, cs in cols.items():
            b = [t for t in cs if B in t]
            nb = [t for t in cs if B not in t and L not in t and not PH.match(t)
                  and len([x for x in t.split('\n') if len(x.strip()) >= 6]) >= 2]
            if b and nb:
                out['C1 同列不统一（有的格多条加点、有的格多条没加点）'].append((hdr + ' 第 %d 列' % (k + 1), show(nb[0])[:60]))
    return out


def audit_notes(src):
    out = collections.defaultdict(list)
    for f in sorted(glob.glob(os.path.join(src, '*', '*.md'))):
        rel = os.path.relpath(f, src)
        L_ = open(f, encoding='utf-8').read().split('\n')
        for i, t in enumerate(L_):
            s = t.strip()
            p = re.sub(r'<[^>]+>', '', s)
            if re.match(r'^注\s*:|^注 +：', p):
                out['N1 「注」后用了半角冒号或多了空格'].append((rel, i + 1, p[:50]))
            m = re.match(r'^注：([^：，。；——]{2,14})：', p)
            if m and '——' not in p[:30]:
                out['N2 「注：小标题：」应写成「注：小标题——」'].append((rel, i + 1, p[:50]))
            if re.match(r'^(说明|备注|提示|附注)[：:]', p):
                out['N3 用了「说明 / 备注 / 提示」而不是「注：」'].append((rel, i + 1, p[:50]))
        # 表后不带「注：」的补充说明（引导句、公司差异、警告、来源、详见除外）
        for i, t in enumerate(L_):
            if t.strip() != '</table>':
                continue
            j = i + 1
            while j < len(L_):
                s = L_[j].strip()
                if not s:
                    j += 1
                    continue
                if re.match(r'(#|<table|%%|<!--|注|来源|出处|详见|---|\|)', s):
                    break
                p = re.sub(r'<[^>]+>', '', s)
                k = j + 1
                while k < len(L_) and not L_[k].strip():
                    k += 1
                lead = p.endswith(('：', ':')) and k < len(L_) and L_[k].strip().startswith(('<table', '-'))
                if not lead and not re.match(r'(公司差异|警告|注意)[：:]', p):
                    out['N4 表后补充说明没有「注：」'].append((rel, j + 1, p[:60]))
                j += 1
    return out


if __name__ == '__main__':
    cells, src, rep = sys.argv[1], sys.argv[2], sys.argv[3]
    a, n = audit_cells(cells), audit_notes(src)
    lines = ['# 排版规则审查（SD-119 / SD-120）', '', '自动审查结果：每类先给条数，再列全部条目。「R7a」「R8」「N4」这类需要人工判断。', '']
    for title, d in (('一、表格单元格', a), ('二、注释', n)):
        lines += ['## ' + title, '']
        for k in sorted(d):
            lines.append('### %s（%d 条）' % (k, len(d[k])))
            lines += ['- ' + ' ｜ '.join(str(x) for x in e) for e in d[k]]
            lines.append('')
    open(rep, 'w', encoding='utf-8').write('\n'.join(lines))
    for d in (a, n):
        for k in sorted(d):
            print('%-40s %d' % (k, len(d[k])))
