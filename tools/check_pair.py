#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_pair.py —— 两本同表对比（2026-10-16 用户：「速查版和完整版相同的内容保持一致，我说改一处那另外一处也要统一」；待改清单第 6 / 8 / 16 条）

读生成器排版导出 build/tbl_dump_*.jsonl（sync.sh 设 TBL_DUMP=1，每格记对齐、各段前缀、列宽占比），
把速查版每张表与完整版内容最接近的表配对（格文字重合 ≥ 50%），逐项比：
  P1 列数 / 表头文字不同                                  —— 错误
  P2 同一格文字相同、排法不同（居中 / 靠左、加不加「•」）    —— 错误
  P3 两本同一行文字不同（速查精简过的除外：速查格的字全在完整版格里）—— 提示
  P4 两本各有对方没有的行                                  —— 提示（速查可少收行；完整版少了速查的行要查）
  P5 手工 <br> 两段合起来按当前格宽一行排得下（有意版式除外）  —— 提示（第 7 条）
报告：build/两本同表对照报告.md。用法：python3 check_pair.py [--repo 路径]
"""
import sys, os, re, json

REPO = os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1]) if '--repo' in sys.argv else os.getcwd()
B = os.path.join(REPO, 'build')
FULL = os.path.join(B, 'tbl_dump_B737机型理论知识笔记.jsonl')
QR = os.path.join(B, 'tbl_dump_B737机型理论知识速查.jsonl')
OUT = os.path.join(B, '两本同表对照报告.md')

def load(p):
    return [json.loads(l) for l in open(p, encoding='utf-8') if l.strip()]
def norm(t): return re.sub(r'[\s（）()，,；;。：:、/·–\-•]', '', t or '')
def grid(tb):
    rows = {}
    for c in tb['cells']: rows.setdefault(c['r'], []).append(c)
    return [rows[k] for k in sorted(rows)]
def sig(tb): return {norm(c['t'])[:24] for c in tb['cells'] if not c['h'] and len(norm(c['t'])) >= 3 and not re.fullmatch(r'[\u2460-\u2473\d]+', norm(c['t']))}   # 序号格、单字格不参与配对
def hdr(tb): return [c['t'] for r in grid(tb) for c in r if c['h']]
def body(tb): return [r for r in grid(tb) if not any(c['h'] for c in r)]
def look(c):
    """排法：每段对齐 + 前缀，去掉重复（一格多段同排法记一次）"""
    out = []
    for x in c['p']:
        if x not in out: out.append(x)
    return '/'.join(out)
def bigr(t):
    t = re.sub(r'[^一-鿿A-Za-z0-9]', '', t); return {t[i:i + 2] for i in range(len(t) - 1)}

def main():
    if not (os.path.exists(FULL) and os.path.exists(QR)):
        print('两本同表对比：缺排版导出（TBL_DUMP=1 全量后才有），跳过'); return 0
    F, Q = load(FULL), load(QR)
    fsig = [sig(t) for t in F]
    errs, tips = [], []
    npair = 0
    for qt in Q:
        qs = sig(qt)
        if len(qs) < 3: continue
        best, bj = 0, -1
        for j, fs in enumerate(fsig):
            if not fs: continue
            k = len(qs & fs); s = k / min(len(qs), len(fs)) if k >= 3 else 0
            if s > best: best, bj = s, j
        if best < 0.5: continue
        ft = F[bj]
        # 速查汇编表（如「发动机限制差异」把完整版几张表合成一张）不是同一张表：速查行有 ≥ 3 行、≥ 30% 在这张完整版表里找不到 → 不比
        _fk = {norm(r[0]['t'])[:20] for r in body(ft) if r}; _qk = [norm(r[0]['t'])[:20] for r in body(qt) if r]
        _miss = sum(1 for k in _qk if k not in _fk)
        if _miss >= 3 and _miss >= 0.3 * len(_qk): continue
        npair += 1
        name = (body(qt)[0][0]['t'] if body(qt) and body(qt)[0] else '')[:16]
        tag = '速查表 %d / 完整版表 %d「%s」' % (qt['t'], ft['t'], name)
        hq, hf = hdr(qt), hdr(ft)
        nq = sum(c['cs'] for r in grid(qt)[:1] for c in r); nf = sum(c['cs'] for r in grid(ft)[:1] for c in r)
        if nq != nf: errs.append(('P1', tag, '列数 速查 %d ／ 完整版 %d' % (nq, nf)))
        elif [norm(x) for x in hq] != [norm(x) for x in hf]: errs.append(('P1', tag, '表头 速查「%s」／ 完整版「%s」' % (' | '.join(hq), ' | '.join(hf))))
        # 格级：同列、同文字的格排法要一样
        fcell = {}
        for c in ft['cells']:
            if not c['h']: fcell.setdefault((c['c'], c['cs'], norm(c['t'])), c)
        bad = []
        for c in qt['cells']:
            if c['h'] or not norm(c['t']): continue
            m = fcell.get((c['c'], c['cs'], norm(c['t'])))
            if m and look(m) != look(c): bad.append('「%s」速查 %s ／ 完整版 %s' % (c['t'][:14], look(c), look(m)))
        if bad: errs.append(('P2', tag, '；'.join(bad[:4]) + ('（共 %d 格）' % len(bad) if len(bad) > 4 else '')))
        # 行级
        fk = {norm(r[0]['t'])[:20]: r for r in body(ft) if r}
        qk = {norm(r[0]['t'])[:20]: r for r in body(qt) if r}
        diffrow = []
        for k, r in qk.items():
            if k not in fk: continue
            qtxt = ''.join(c['t'] for c in r); ftxt = ''.join(c['t'] for c in fk[k])
            if norm(qtxt) != norm(ftxt):
                b = bigr(qtxt); cov = len(b & bigr(ftxt)) / len(b) if b else 1
                if cov < 0.98: diffrow.append('%s（速查字面在完整版里 %d%%）' % (r[0]['t'][:12], cov * 100))
        if diffrow: tips.append(('P3', tag, '；'.join(diffrow[:4])))
        onlyQ = [qk[k][0]['t'][:12] for k in qk if k not in fk]
        if onlyQ: tips.append(('P4', tag, '完整版没有速查的行：' + '、'.join(onlyQ[:6])))
    # P5（待改清单第 7 条）：手工 <br> 两段合起来一行排得下——有意版式除外：缩写 / 全称上下两行、名称 /【机型】、数值 + 括注、两个参数上下并列、引语「…：」
    brs = []
    for nm, D in (('完整版', F), ('速查', Q)):
        for tb in D:
            for c in tb['cells']:
                a = c.get('br')
                if not a: continue
                x, y = a
                asc = lambda t: re.fullmatch(r'[A-Za-z0-9 /\-.&+]+', t) is not None
                if y.startswith(('（', '(', '【')) or x.startswith('【') or x.endswith(('：', ':')): continue
                if asc(x) != asc(y): continue                                   # 缩写 / 全称
                if re.search(r'\d', x) and re.search(r'\d', y): continue        # 两个参数上下并列
                if re.search(r'[；;。，,、]$', x): continue                          # 两条并列句 / 分句，本来就分条
                if re.fullmatch(r'[A-Z0-9/\-]{2,8}', x) or re.fullmatch(r'[A-Z0-9/\-]{2,8}', y): continue   # 缩写 + 全称
                vw = sum(2 if '\u2e80' <= ch <= '\u9fff' else 1 for ch in x + y)
                if vw > 26: continue                                                # 只查短标签（约 13 个汉字以内）
                brs.append((nm, tb['t'], x + '⏎' + y))
    tips += [('P5', '%s 表 %d' % (nm, t), '「%s」两段合起来一行排得下，手工换行可能多余' % txt) for nm, t, txt in brs]
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('# 两本同表对照报告\n\n配对 %d 张；错误 %d 条、提示 %d 条。\n\n## 错误（P1 列数 / 表头，P2 同格排法）\n' % (npair, len(errs), len(tips)))
        for e in errs: f.write('- [%s] %s：%s\n' % e)
        f.write('\n## 提示（P3 同行文字不同，P4 行不对应，P5 多余手工换行候选）\n')
        for e in tips: f.write('- [%s] %s：%s\n' % e)
    print('两本同表对比：配对 %d 张，错误 %d 条、提示 %d 条 → %s' % (npair, len(errs), len(tips), OUT))
    return 1 if errs else 0

if __name__ == '__main__':
    sys.exit(main())
