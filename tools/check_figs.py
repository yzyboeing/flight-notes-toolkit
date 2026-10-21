#!/usr/bin/env python3
"""K-L7 自绘图数值核对（P05；2026-10-21 规则重组轮方案 A1）。

用法：check_figs.py [笔记库路径]   ——问题打印「FIGS 错误 …」，退出码 1；全部通过退出码 0

读私有库 figures_src/figs.py（只做语法解析，不执行、不出图）：
1. 每个 @fig('图号') 函数里所有字符串中的数值（数字 + 单位）都要在所在节的两本正文里找得到
   （完整版该节文件 + 速查里绑定到该节的条目）；只作示意的数值登记在 FIGVALS_OK 并写理由。
2. TYPEVALS 登记的两型数值：正文里该数值必须属于同一机型（两型对比表的机型列，或【737-NG】/【737-8】标注之后），
   属于另一机型 → 错误（教训：1.7-6 起动时序图曾把两型 N2 标反）；两边都认不出 → 错误，请人工看。
"""
import ast, glob, os, re, sys
from pathlib import Path

NOTES = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path(__file__).resolve().parents[2] / 'gh-private'
FIGS = NOTES / 'figures_src' / 'figs.py'
VAL = re.compile(r'(?<![\w.])(\d+(?:\.\d+)?)\s*(%|kt|ft|fpm|psi|nm|NM|km|m|℃|°C|°|s|min|V|Hz|hPa|mb)(?![A-Za-z])')
TYPES = ('737-NG', '737-8')


def plain(t):
    return re.sub(r'\s+', '', re.sub(r'<[^>]+>', '', t).replace('&lt;', '<').replace('&gt;', '>'))


def strip(t):
    return re.sub(r'<[^>]+>', '', t).replace('&lt;', '<').replace('&gt;', '>')


def vals(s):
    return {m.group(1) + m.group(2) for m in VAL.finditer(s)}


def type_facts(text):
    """正文里属于某一机型的数值：两型对比表的机型列 + 【机型】标注之后到下一个【或格尾。"""
    facts = set()
    for tb in re.findall(r'(?s)<table.*?</table>', text):
        cols = None
        for tr in re.findall(r'(?s)<tr[^>]*>.*?</tr>', tb):
            cells, i = [], 0
            for m in re.finditer(r'(?s)<t([dh])([^>]*)>(.*?)</t[dh]>', tr):
                span = int((re.search(r'colspan="(\d+)"', m.group(2)) or [0, 1])[1])
                cells.append((i, span, m.group(3))); i += span
            if 'class="hdr"' in tr:
                cols = {}
                for i0, sp, c in cells:
                    for t in TYPES:
                        if re.search(r'%s(?![\d-])' % re.escape(t), plain(c)): cols[i0] = t
                continue
            if not cols: continue
            for i0, sp, c in cells:
                if sp == 1 and i0 in cols:
                    facts |= {(cols[i0], v) for v in vals(strip(c))}
    for m in re.finditer(r'【(737-NG|737-8)】([^【\n<]*)', re.sub(r'</?(?:em|b|strong)>', '', text)):
        facts |= {(m.group(1), v) for v in vals(m.group(2))}
    return facts


def main():
    errs = []
    if not FIGS.is_file():
        print('FIGS 提示 没有 figures_src/figs.py，跳过'); return 0
    tree = ast.parse(FIGS.read_text(encoding='utf-8'))
    figs, ok, typevals = {}, {}, {}
    for n in tree.body:
        if isinstance(n, ast.Assign) and isinstance(n.targets[0], ast.Name) and n.targets[0].id in ('FIGVALS_OK', 'TYPEVALS'):
            (ok if n.targets[0].id == 'FIGVALS_OK' else typevals).update(ast.literal_eval(n.value))
        if isinstance(n, ast.FunctionDef):
            for d in n.decorator_list:
                if isinstance(d, ast.Call) and getattr(d.func, 'id', '') == 'fig':
                    figs[d.args[0].value] = [c.value for c in ast.walk(n) if isinstance(c, ast.Constant) and isinstance(c.value, str)]
    qsrc = (NOTES / '速查' / '速查源.md').read_text(encoding='utf-8')
    qents = re.split(r'(?m)^### ', qsrc)
    for k, ss in sorted(figs.items()):
        sec = k.split('-')[0]
        full = ''.join(open(f, encoding='utf-8').read() for f in glob.glob(str(NOTES / 'notes_src' / '*' / (sec + ' *.md'))))
        quick = ''.join(e for e in qents if re.search(r'\[\[%s ' % re.escape(sec), e.split('\n', 2)[1] if '\n' in e else ''))
        if not full and not quick:
            errs.append('图 %s：找不到第 %s 节的正文' % (k, sec)); continue
        text = plain(full + quick)
        for v in sorted(set().union(*[vals(s) for s in ss]) if ss else ()):
            if v not in text and v not in ok.get(k, {}):
                errs.append('图 %s 里的数值「%s」在 %s 节两本正文里都找不到——图中数值要取自正文（只作示意的登记 FIGVALS_OK 并写理由）' % (k, v, sec))
        facts = type_facts(full + quick)
        for t, vs in typevals.get(k, {}).items():
            other = TYPES[1 - TYPES.index(t)]
            for v in vs:
                if v not in set().union(*[vals(x) for x in ss]): errs.append('图 %s 的 TYPEVALS 登记了「%s」，图里没有这个数值——登记与图不符' % (k, v))
                if (t, v) in facts: continue
                if (other, v) in facts: errs.append('图 %s 把「%s」标成 %s，正文里它属于 %s——两型标反' % (k, v, t, other))
                else: errs.append('图 %s 的 %s 数值「%s」在正文里认不出属于哪一型——人工核对' % (k, t, v))
    for e in errs: print('FIGS 错误', e)
    print('自绘图核对：%d 张图，问题 %d' % (len(figs), len(errs)))
    return 1 if errs else 0


if __name__ == '__main__':
    sys.exit(main())
