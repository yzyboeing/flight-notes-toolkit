#!/usr/bin/env python3
import re
# -*- coding: utf-8 -*-
"""fit_fix.py —— 多遍排版：一页放得下的表绝不拆开；略超一页的表逐级压缩后整表同页（SD-79 / SD-80，2026-09-30 用户）
用法：python3 fit_fix.py <输入.md> <输出.docx>      （环境变量照常传给 build_docx.js）
做法：
  1. build_docx.js 出 docx（TBL_DUMP 记下每张表的签名），LibreOffice 转 PDF，check_splits.py --json 找断表；
  2. 按表内文字（字符二元组覆盖率）把断表匹配回签名，写进 build/keep_force_<输出名>.txt，重出，直到稳定（最多 7 遍）：
     · 断开但放得下一页（连同条目标题、表前说明）→ 强制整表同页（行：签名）；
     · 放不下但超出不多（≤ 1.2 页，且不是跨 3 页的大表）→ 逐级压缩（行：~N:签名）：
         1 级 9pt 不变、收紧行距与单元格边距；2 级 8.5pt；3 级 8pt（底线）；
       按超出比例直接从合适的一级起试，仍断开就升一级；
     · 3 级仍放不下 → 放弃（行：!签名），照常行间分页 + 防孤行。
  · 末行孤字 / 短格折行（layout_measure.py 实测）→ 这一列加宽（行：W:列:DXA:签名），从最宽列匀出；加宽两次仍不行就放弃（WB:）。
清单跨次保留，下次构建第一遍就生效；表格内容改动后签名对不上，自动撤出。
是否放得下以 PDF 实测为准，不看估算。"""
import sys, os, json, subprocess, tempfile, shutil
T = os.path.dirname(os.path.abspath(__file__))
if len(sys.argv) < 3: print(__doc__); sys.exit(1)
MD, OUT = sys.argv[1], sys.argv[2]
name = os.path.splitext(os.path.basename(OUT))[0]
KF = os.path.join(os.path.dirname(OUT) or '.', 'keep_force_%s.txt' % name)
SOF = shutil.which('soffice') or '/Applications/LibreOffice.app/Contents/MacOS/soffice'
PROF = os.path.join(os.environ.get('TMPDIR', '/tmp'), 'lo-sync-profile')
MAXLV = 3

def build(dump):
    env = dict(os.environ, TBL_DUMP=dump, KEEP_FORCE=KF)
    r = subprocess.run(['node', os.path.join(T, 'build_docx.js'), MD, OUT], env=env)
    if r.returncode: sys.exit('build_docx.js 失败')

def lo_fonts(prof):
    # SD-107：LibreOffice 不读 ~/Library/Fonts，把思源字体复制进 profile 的 user/fonts（已存在且大小相同就跳过）
    import glob
    fd = os.path.join(prof, 'user', 'fonts'); os.makedirs(fd, exist_ok=True)
    for f in glob.glob(os.path.expanduser('~/Library/Fonts/SourceHan*.otf')):
        t = os.path.join(fd, os.path.basename(f))
        if not os.path.exists(t) or os.path.getsize(t) != os.path.getsize(f): shutil.copy2(f, t)

def to_pdf():
    lo_fonts(PROF)
    d = tempfile.mkdtemp()
    subprocess.run([SOF, '-env:UserInstallation=file://' + PROF, '--headless', '--convert-to', 'pdf', '--outdir', d, OUT],
                   capture_output=True)
    p = os.path.join(d, name + '.pdf')
    if not os.path.exists(p): sys.exit('LibreOffice 没能生成 PDF')
    return p

# 状态：lv[sig] = 0（强制整表）/ 1～3（压缩级）；block＝放弃
lv, block, wfix, wblock, splitok, pbreak, nwfix = {}, set(), {}, set(), set(), set(), {}
figw = {}   # 1020R1 图宽实测：图文件名 → 宽度 mm（FW:）
bw, bwblock, bwtried = {}, set(), {}   # 2026-10-15 借宽：bw[(签名, 收列, 借出列)] = DXA；bwblock＝借过 3 次仍不行的 (签名, 收列)
ffix = {}   # SD-148 实测拉满：F:列号:加宽DXA:签名（生成器落实，目标列做过不增行收窄时转给同表最挤的列）
nlbase = {}   # (表, 列) → 收窄前的总行数
nwlock = set()   # 收窄后又折行的 (表, 列)：本次运行不再收窄（2026-10-06 防拉锯）
ncond = {}   # SD-130 表外「注：」段落末行孤字：ncond[注文前 12 字] = 1（-0.3pt）/ 2（-0.5pt）
condense = {}   # 2026-10-03 末行孤字兜底：加宽无效的格收紧字距，condense[(sig, 首行前 10 字)] = 1（-0.3pt）/ 2（-0.5pt）   # pbreak＝另起一页的条目标题（P:，SD-97 孤行兜底）   # splitok＝标题被留下的块索引表：允许按块分页（S:）     # wfix[(sig, 列)] = 加宽 DXA；wblock＝加宽也没用的格，不再试
for l in (open(KF, encoding='utf-8') if os.path.exists(KF) else []):
    l = l.strip()
    if not l: continue
    if l.startswith('W:'):
        _, k, dd, sg = l.split(':', 3); wfix[(sg, int(k))] = int(dd)
    elif l.startswith('S:'):
        splitok.add(l[2:])
    elif l.startswith('P:'):
        pbreak.add(l[2:])
    elif l.startswith('WB:'):
        pass   # 1020R1（用户 2026-10-20 第 290 页「右侧有很大空间，为什么不增加表格宽度以减少行数」）：WB（加宽放弃）不再跨运行保留——旧记录让 56 张窄表永远不加宽；本次运行内仍防拉锯
    elif l.startswith('NW:'):
        _, k, dd, sg = l.split(':', 3); nwfix[(sg, int(k))] = int(dd)
    elif l.startswith('NWL:'):
        pass   # 1020R1（用户 2026-10-20 第 9 张截图「这一列空白区域过大」）：NWL 不再跨运行保留——旧记录在内容 / 分页变了以后照样锁住收窄、并让 T13 不报；每次全量重新实测
    elif l.startswith('BW:'):
        _, k, j, dd, sg = l.split(':', 4); bw[(sg, int(k), int(j))] = int(dd)
    elif l.startswith('BWB:'):
        pass   # 1020R1：BWB（借宽放弃）同样只在本次运行内有效
    elif l.startswith('F:'):
        _, k, dd, sg = l.split(':', 3); ffix[(sg, int(k))] = int(dd)
    elif l.startswith('C:'):
        _, cl, rest = l.split(':', 2); sg, fk = rest.rsplit('|', 1); condense[(sg, fk)] = int(cl)
    elif l.startswith('N:'):
        _, cl, fk = l.split(':', 2); ncond[fk] = int(cl)
    elif l.startswith('FW:'):
        _, mm, fn = l.split(':', 2); figw[fn] = float(mm)
    elif l.startswith('!'): block.add(l[1:])
    elif l.startswith('~'): lv[l[3:]] = int(l[1])
    else: lv[l] = 0
def save():
    with open(KF, 'w', encoding='utf-8') as f:
        for s, v in sorted(lv.items()): f.write((('~%d:' % v) if v else '') + s + '\n')
        for s in sorted(block): f.write('!' + s + '\n')
        for (sg, k), dd in sorted(wfix.items()): f.write('W:%d:%d:%s\n' % (k, dd, sg))
        for (sg, k) in sorted(wblock): f.write('WB:%d:%s\n' % (k, sg))
        for (sg, k), dd in sorted(nwfix.items()): f.write('NW:%d:%d:%s\n' % (k, dd, sg))
        for (sg, k), dd in sorted(ffix.items()): f.write('F:%d:%d:%s\n' % (k, dd, sg))
        for (sg, k, j), dd in sorted(bw.items()): f.write('BW:%d:%d:%d:%s\n' % (k, j, dd, sg))
        for (sg, k) in sorted(bwblock): f.write('BWB:%d:%s\n' % (k, sg))
        for (sg, k) in sorted(nwlock): f.write('NWL:%d:%s\n' % (k, sg))
        for sg in sorted(splitok): f.write('S:' + sg + '\n')
        for h in sorted(pbreak): f.write('P:' + h + '\n')
        for (sg, fk), cl in sorted(condense.items()): f.write('C:%d:%s|%s\n' % (cl, sg, fk))
        for fk, cl in sorted(ncond.items()): f.write('N:%d:%s\n' % (cl, fk))
        for fn, mm in sorted(figw.items()): f.write('FW:%.1f:%s\n' % (mm, fn))

bg = lambda x: {x[k:k + 2] for k in range(len(x) - 1)}
def match(text, tbls):
    """返回最匹配的表签名集合（同文表分数相近的一并返回）；匹配不上返回空集。"""
    want = bg(text)
    if not want: return set()
    score = sorted(((len(want & t['bg']) / len(want), t['sig']) for t in tbls), reverse=True)
    if not score: return set()
    if score[0][0] < 0.75:
        # 跨页表、紧贴表后注的表：PDF 上截到的文字只是源表的一部分，覆盖率偏低。最高分 ≥ 0.5 且明显高于第二名（≥ 0.2）也算匹配上（2026-10-05）
        second = score[1][0] if len(score) > 1 else 0
        return {score[0][1]} if score[0][0] >= 0.5 and score[0][0] - second >= 0.2 else set()
    return {s for v, s in score if v >= max(0.75, score[0][0] - 0.1)}

dump = os.path.join(tempfile.mkdtemp(), 'dump.jsonl')
# 2026-10-12（用户：固定来回跳的列宽）：逐遍记录各表实际列宽与成品取哪一遍，供 fixosc.py 找来回跳的表并定宽
TRACE = os.path.join(os.path.dirname(OUT) or '.', 'fitfix_trace_%s.jsonl' % name)
open(TRACE, 'w').close()
def trace(obj):
    with open(TRACE, 'a', encoding='utf-8') as tf: tf.write(json.dumps(obj, ensure_ascii=False) + '\n')
tried = {}
import time, hashlib
MAXPASS = int(os.environ.get('FITFIX_MAXPASS', '10'))   # 预览（preview.py）用 4 遍
seen_state = {}   # 2026-10-07 提前停止：记录每遍开工前的状态，状态重复＝来回拉锯，再跑也只是循环（用户：「为什么每次都这么久」）
last_pdf = None
prev_calm = False
hist = []
snap = {}
for n in range(1, MAXPASS + 1):   # 2026-10-06：7 → 10 遍（列宽规则多了，收敛要多几遍）
    save()
    st = hashlib.md5(open(KF, 'rb').read()).hexdigest()
    if st in seen_state:
        print('fit_fix：第 %d 遍开工前的状态与第 %d 遍相同（列宽来回拉锯），停止，成品取上一遍' % (n, seen_state[st])); trace({'best': n - 1, 'why': 'state-repeat'}); break
    seen_state[st] = n
    t0 = time.time()
    build(dump)
    tbls = [json.loads(l) for l in open(dump, encoding='utf-8') if l.strip()]
    for t in tbls: t['bg'] = bg(t['text'])
    trace({'pass': n, 'W': {t['sig']: t.get('W') for t in tbls}})
    _full = {t['sig'] for t in tbls if t.get('fixed') == 'full'}; _narrow = {t['sig'] for t in tbls if t.get('fixed') == 'narrow'}
    for _d in (wfix, nwfix, ffix):   # 定宽表名下生成器不认的旧记录撤掉（满宽：W / NW / F；窄：NW / F）
        for _k in [k2 for k2 in _d if k2[0] in _full or (_d is not wfix and k2[0] in _narrow)]: del _d[_k]
    sigs = {t['sig'] for t in tbls}
    onep = {t['sig'] for t in tbls if t.get('onepage')}
    changed = [s for s in list(lv) if s not in sigs]            # 表已改动：撤出
    for s in changed: del lv[s]
    block &= sigs
    for key in [k2 for k2 in wfix if k2[0] not in sigs]: del wfix[key]
    for key in [k2 for k2 in nwfix if k2[0] not in sigs]: del nwfix[key]
    for key in [k2 for k2 in ffix if k2[0] not in sigs]: del ffix[key]
    for key in [k2 for k2 in bw if k2[0] not in sigs]: del bw[key]
    bwblock = {k2 for k2 in bwblock if k2[0] in sigs}
    nwlock = {k2 for k2 in nwlock if k2[0] in sigs}
    wblock = {k2 for k2 in wblock if k2[0] in sigs}
    pdf = to_pdf(); last_pdf = os.path.splitext(OUT)[0] + '.fitfix.pdf'; shutil.copy2(pdf, last_pdf)   # 测量后临时目录会删，先留一份
    # 2026-10-14 提速：断表检查与下面的实测并行跑，到下面删临时目录前再收
    _cs = subprocess.Popen([sys.executable, os.path.join(T, 'check_splits.py'), pdf, '--json'],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    # SD-85 实测加宽：末行孤字、短格折行 → 这张表这一列加宽，下一遍重排（原文 <br> 主动换行的不算）
    sys.path.insert(0, T); from layout_measure import measure
    wadd, wgive, now, cadd, nwcand, nwwrap, grown_sg, nlnow = 0, 0, set(), 0, {}, set(), set(), {}
    _xs = measure(pdf, srcmd=MD)
    # 2026-10-15 借宽：先按 (表, 源列) 汇总可借出宽度（跨页的表取各段最小值），再处理孤字 / 省行
    def _srccol(x, tb):
        Wt = tb.get('W') or []
        if not Wt: return None
        pos, acc, col = x['cx'] * sum(Wt) / max(x['tw'], 1), 0, len(Wt) - 1
        for q, wq in enumerate(Wt):
            acc += wq
            if pos <= acc: return q
        return col
    lendm = {}
    for x in _xs:
        if not x.get('nw') or 'lend' not in x or x['col'] < 0: continue
        for sg in match(x['table'], tbls):
            tb = next(t for t in tbls if t['sig'] == sg); c = _srccol(x, tb)
            if c is None: continue
            lendm[(sg, c)] = min(lendm.get((sg, c), 10 ** 9), int(x['lend'] * 20))
    TWO = re.compile(r'737-NG|737-8|两型')
    bw_now = set()   # 本遍已借宽的表（每表每遍一次，免得几列同时借、互相挤）
    def borrow(sg, tb, k, need, why):
        if (sg, k) in bwblock or sg in bw_now: return False
        eq, hdr = set(tb.get('eq') or []), tb.get('hdr') or []
        if k in eq or any(s2 == sg and jj == k for (s2, _, jj) in bw): return False   # 等宽列、已在借出的列不收宽（防来回借）
        donors = [(lendm.get((sg, j), 0), j) for j in range(len(tb.get('W') or [])) if j != k and j not in eq
                  and not any(kk == j for (s2, kk, _) in bw if s2 == sg)]   # 等宽列（col-eq）不借出；已在收宽的列不借出（SD-192 起两型对比列也可让出）
        # SD-192：单列让得出就向让得最多的一列借；单列不够、几列合起来够，就从可让最多的几列依次各借一部分（每列都不超过其可让宽度，不增行）
        donors = sorted([d for d in donors if d[0] >= 40], reverse=True)   # 可让不足 2pt 的列不动
        if not donors or sum(d[0] for d in donors) < need: return False
        take, left = [], need
        for lv0, j in donors:
            if left <= 0: break
            t = min(lv0, left); take.append((lv0, j, t)); left -= t
        bwtried[(sg, k)] = bwtried.get((sg, k), 0) + 1
        if bwtried[(sg, k)] > 3: bwblock.add((sg, k)); [bw.pop(kk) for kk in [kk for kk in bw if kk[0] == sg and kk[1] == k]]; return False
        for lv0, j, t in take:
            bw[(sg, k, j)] = bw.get((sg, k, j), 0) + t; lendm[(sg, j)] = lv0 - t
            trace({'borrow': sg[:20], 'to': k, 'from': j, 'dxa': t, 'why': why})
        bw_now.add(sg)
        return True
    nbw = 0
    # 借多了的退回：收宽列孤字已消、实测仍可让出 ≥ 10pt 时，退回多出的部分，只留约 6pt（与 T13 同口径；可让出宽度已保证退回后不增行、不再出孤字）
    # SD-192（2026-10-16）：孤字借宽不再「借多了退回」——退回会让收宽列孤字复现、借出列反借回来，两列来回跳（第 452 页 4.11 安定面表 10 遍不收敛）。
    # 改为：借出列本遍出现孤字 → 撤销这次借宽并锁定（BWB），收宽列改走收紧字距 / 接受
    _orph_cols = set()
    for x0 in _xs:
        if x0.get('orphan') and not x0.get('nw'):
            for sg0 in match(x0['table'], tbls):
                tb0 = next(t for t in tbls if t['sig'] == sg0); c0 = _srccol(x0, tb0)
                if c0 is not None: _orph_cols.add((sg0, c0))
    for (sg, k, j), dd in list(bw.items()):
        if (sg, j) in _orph_cols:
            for kk in [kk for kk in bw if kk[0] == sg and kk[1] == k]: bw.pop(kk)
            bwblock.add((sg, k)); bw_now.add(sg); nbw += 1
            trace({'borrow_undo': sg[:20], 'to': k, 'from': j, 'dxa': dd, 'why': 'donor-orphan'})
    # 2026-10-15 首次试运行：省行借宽（growb）与 T14「空位让给折行列」互相打架、整表来回挪，借宽只用于消末行孤字；省行仍由 SD-144 / T14 处理
    for x in [x0 for x0 in _xs if x0.get('orphan') and not x0.get('nw')]:   # 借宽：定宽满宽表的孤字，以及普通加宽已放弃（WB）的孤字
        if True:
            for sg in match(x['table'], tbls):
                tb = next(t for t in tbls if t['sig'] == sg); k = _srccol(x, tb)
                if k is None or x.get('span'): continue
                if x.get('growb') and (tb.get('fixed') != 'full' and (sg, k) not in wblock): continue   # 省行借宽只给定宽满宽表和普通加宽已放弃的列（其余表由 SD-144 加宽 / 拉满处理）
                # SD-192（2026-10-16 用户「可在不增加行数的情况下适当收窄其余列宽度以减少本列的孤字问题」）：孤字一律先借宽（同表别列可让出、不增行；表总宽不变，不与加宽 / 拉满互相拉锯），借不到才走加宽 / 收紧字距
                if (x.get('prev') and x['prev'] + '|' in tb.get('br', '')): continue   # 止于原文 <br> 的是主动换行
                need = int(x['extra_pt'] * 20) + 20
                if borrow(sg, tb, k, need, 'orphan' if x.get('orphan') else 'growb'): nbw += 1
    for x in _xs:
        for sg in match(x['table'], tbls):           # 速查区与正文同文的表一并加宽
            tb = next(t for t in tbls if t['sig'] == sg)
            fx = tb.get('fixed') or ''   # 2026-10-12：定宽表（w-NN）——满宽的生成器不吃任何实测调整，窄的只吃加宽（W:）；不再记账，免得每遍空转
            if fx == 'full' or (fx == 'narrow' and (x.get('nw') or x.get('fill'))): continue   # SD-204：origin（定宽起始）照常收窄 / 拉满
            # 列号按格子中心的横向位置换算回源表列（PDF 识别合并单元格时会多切列，列号不可靠）
            Wt = tb.get('W') or []
            if not Wt: continue
            pos, acc, col = x['cx'] * sum(Wt) / max(x['tw'], 1), 0, len(Wt) - 1
            for q, wq in enumerate(Wt):
                acc += wq
                if pos <= acc: col = q; break
            key = (sg, col)
            if x.get('nw'):   # 实测收窄（SD-144）：先按 (表, 列) 收集，跨页的表取各段空白的最小值（0 是否决票），循环后再入账
                if x['col'] < 0:
                    for q in range(len(Wt)): nwcand[(sg, q)] = 0
                else:
                    ex = x['extra_pt'] if (key not in nwfix and fx != 'origin') else x.get('extra_strict', x['extra_pt'])   # SD-204：定宽起始表只按「本格行数不增」收窄，不借行高（实测会把实际折行的列收窄）   # 借行高收窄只给一次：已有收窄记录的列只按「本格行数不增」继续收（防一遍遍往下收）
                    if fx == 'origin': ex = max(0, ex - 14)   # SD-204：定宽起始表收窄多留 14pt（粗体 / 句末标点实宽偏大，收到最长行 + 6pt 会让该行下一遍折行、再拉回，来回拉锯）
                    nwcand[key] = min(nwcand.get(key, 10 ** 9), int(ex * 20))
                    nlnow[key] = nlnow.get(key, 0) + x.get('nl', 0)   # 本遍这一列的总行数（跨页各段相加）
                continue
            if key in nwfix and (x.get('orphan') or x.get('grow')): del nwfix[key]   # 同一列又需要加宽：撤销收窄，省行优先
            if key in now or (x.get('prev') and x['prev'] + '|' in tb.get('br', '')) or any(l and l + '|' in tb.get('br', '') for l in x.get('brs', [])): continue   # brs：SD-144 加宽项里各行，任一行止于原文 <br> 就不是折行
            if x.get('span') and not x.get('orphan'): continue   # 跨列格的短格折行不处理
            if x.get('fill'):   # SD-148 实测拉满：本表这遍没有落实的省行加宽时，右侧实测空余整块给最挤的列（不受加宽上限、次数与放弃名单限制）
                if sg not in grown_sg: ffix[key] = ffix.get(key, 0) + int(x['extra_pt'] * 20); wadd += 1
                continue
            # 跨列格的孤字也先试加宽（加宽所跨的那一列即加宽整格，捐出列可能在格外）；加宽无效再收紧字距（2026-10-05）
            if sg in bw_now and any(kk[0] == sg and kk[1] == col for kk in bw): now.add(key); continue   # 本遍刚借宽（含 SD-192 加宽后借宽），先看效果，不叠加加宽 / 收紧字距
            if key in wblock:   # 加宽无效（表已占满版面、邻列太窄）（表已占满版面、邻列太窄）：只在末行孤字时收紧这一格的字距，最多两级
                fk = (x.get('first') or '')[:10]
                if fk and x.get('orphan') and condense.get((sg, fk), 0) < 4: condense[(sg, fk)] = condense.get((sg, fk), 0) + 1; cadd += 1
                now.add(key); continue   # 2026-10-03：与 check_layout T8 同口径——只有倒数第二行止于原文 <br> 才算主动换行（原来看第一行，带 <br> 的格一律跳过）
            now.add(key)
            new = wfix.get(key, 0) + int(x['extra_pt'] * 20) + 20
            cap = wfix.get(key, 0) + int(x['extra_pt'] * 20) + 400 if x.get('grow') else 1500   # SD-144 加宽项只用右侧空余（实测已保证放得下），不受 75pt 上限；1020R1：上限按「已加宽 + 本遍需求」算（原来只按本遍需求，第二遍必然超限放弃）
            if new > cap or tried.get(key, 0) >= (6 if x.get('grow') else 3): wblock.add(key); wfix.pop(key, None); wgive += 1; continue   # 加宽 3 次（省行加宽 6 次）或超过上限仍不行：放弃
            wfix[key] = new; tried[key] = tried.get(key, 0) + 1; wadd += 1
            if x.get('grow'): grown_sg.add(sg)   # 本表这遍已实测加宽，拉满留到下一遍
    # 过时记录按本遍实测撤销（2026-10-06，第 12 页 C-4：早期记下的「加宽」「收窄」互相卡住，宽度分错了列）：
    #   收窄过的列现在折行了 → 撤掉收窄；加宽过的列现在没折行且右侧空 ≥ 10pt → 加宽减回去
    # 有阻尼，防来回拉锯：收窄后又折行的列退回 15pt 并锁定（本次运行不再收窄）；加宽过的列只有空白超过加宽余量（≥ 30pt）时才减，且只减多出余量的部分
    # 收窄过头 ＝ 收窄后这一列行数变多了（2026-10-06：原来看「某行排满且下面还有一行」，收窄到正好放下最长行时最长行必然排满，误判后锁住、收到一半停下）
    for key, nl_ in nlnow.items():
        if key in nwfix and key in nlbase and nl_ > nlbase[key]:
            nwfix[key] -= 300; nwlock.add(key); wadd += 1
            if nwfix[key] <= 0: del nwfix[key]
        if key not in nwfix or key not in nlbase: nlbase[key] = nl_   # 收窄前的行数作基准（还没收窄、或撤销干净时更新）
    for key in nwwrap: pass
    for key, dd in nwcand.items():
        if key in ffix and key not in nwwrap and dd >= 200:   # 2026-10-06：空白 ≥ 10pt 就撤，撤到只留约 6pt（与 T13 的 10pt 同口径）
            ffix[key] -= dd - 120; wadd += 1
            if ffix[key] <= 0: del ffix[key]
    for key, dd in nwcand.items():
        if key in wfix and key not in nwwrap and dd >= 200:   # 2026-10-06：空白 ≥ 10pt 就撤，撤到只留约 6pt（与 T13 的 10pt 同口径）
            wfix[key] -= dd - 120; wadd += 1
            if wfix[key] <= 0: del wfix[key]
    for key, dd in nwcand.items():   # 实测收窄入账：加宽过的列不收；每遍按当前成品的实测空白累加，收到位后空白 < 10pt 自然停
        if key in wfix or key in ffix or key in nwlock or dd < 200: continue   # 不到 10pt 的空白不折腾（2026-10-06 用户：表宽以文字成行为标准，原 18pt）
        if nwfix.get(key, 0) < 30000: nwfix[key] = min(nwfix.get(key, 0) + dd, 30000); wadd += 1   # 2026-10-06：上限由 6000 放开（初始分宽过大的列收不到位，第 12、60 页）
    # 1020R1 图宽实测（用户 2026-10-20：「图片的下面和左右空间都很大，为什么把图片放这么小」「尽量利用空白区域」）：
    #   独立插图（%%FIG%%）下方到页脚还有空白（图后可有注 / 段落） → 按空白放大（留 18pt 余量），上限：版心宽、原图 150dpi
    import fitz as _fz
    fadd = 0
    try:
        _figsrc = {}; _figbase = {}; _figfixed = set()
        for _l in open(MD, encoding='utf-8'):
            _m = re.match(r'%%FIG%%\s*([^|]+)\|\s*(图\s*[\d.]+-\d+)[^|]*\|\s*([\d.]+)', _l)
            if _m: _figsrc[re.sub(r'\s', '', _m.group(2))] = os.path.basename(_m.group(1).strip()); _figbase[os.path.basename(_m.group(1).strip())] = float(_m.group(3))
            if _m and re.search(r'\|\s*定\s*(\||$)', _l): _figfixed.add(os.path.basename(_m.group(1).strip()))   # 第 96 / 103 条：图行标「定」＝用户定过的宽度，不再按版面放大 / 缩小
        _figdir = os.environ.get('DOC_FIG_DIR') or os.path.join(os.getcwd(), 'notes_src')
        _prev = None   # 上一页：(页脚上缘, 正文最低处)
        for _pg in _fz.open(pdf):
            _H = _pg.rect.height; _Wd = _pg.rect.width
            _lines = []
            for _b in _pg.get_text('dict')['blocks']:
                for _ln in _b.get('lines', []):
                    _t = ''.join(_sp['text'] for _sp in _ln['spans']).strip()
                    if _t: _lines.append((_ln['bbox'], _t))
            _foot = min([_bb[1] for _bb, _t in _lines if re.search(r'第\s*\d+\s*页', _t) and _bb[1] > _H * 0.85] or [_H - 40])
            _imgs = [_fz.Rect(_im['bbox']) for _im in _pg.get_image_info() if _fz.Rect(_im['bbox']).width >= 120]
            _bot = max([_bb[3] for _bb, _t in _lines if _bb[3] < _foot - 2] + [r.y1 for r in _imgs] + [0])
            for _ib in _imgs:
                _cap = [(_bb, _t) for _bb, _t in _lines if _t.startswith('图') and 0 <= _bb[1] - _ib.y1 <= 30 and _bb[0] < (_ib.x0 + _ib.x1) / 2 < _bb[2] + 200]
                if not _cap: continue
                _cb, _ct = _cap[0]
                _mm0 = re.match(r'(图\s*[\d.]+-\d+)', _ct)
                _fn = _figsrc.get(re.sub(r'\s', '', _mm0.group(1))) if _mm0 else None
                if not _fn or _fn in _figfixed: continue
                if any(_bb[3] > _ib.y0 and _bb[1] < _ib.y1 and (_bb[2] < _ib.x0 - 4 or _bb[0] > _ib.x1 + 4) for _bb, _t in _lines): continue   # 图旁有文字（并排）
                _cur = _ib.width / 72 * 25.4
                # ① 被挤到本页顶上、上一页底部空着：缩到上一页能放下（不小于原来的 60%）
                if _prev and not any(_bb[3] <= _ib.y0 + 1 for _bb, _t in _lines) and _ib.y0 < 90:
                    _pfree = _prev[0] - 12 - _prev[1] - (_cb[3] - _ib.y1) - 30
                    _base = _figbase.get(_fn, _cur)
                    if _pfree >= 0.8 * _ib.height and figw.get(_fn, _base) >= _base - 0.5:   # 1020R1：只缩一次、不小于源文件宽度的 80%（用户：图不要太小）
                        _mm = round(max(_base * 0.8, _cur * min(1.0, _pfree / _ib.height) * 0.97), 1)
                        if _cur - _mm >= 3: figw[_fn] = _mm; fadd += 1; continue
                # ② 下方到页脚还空：按空白放大。1020R1：图下只有注 / 段落、整页最低处离页脚还空时也放大（图后的注随图下移，正好填满）；一页多张图只放大最下面那张
                if _ib.y1 < max(r.y1 for r in _imgs) - 1: continue
                _below = [_bb[3] for _bb, _t in _lines if _bb[1] > _cb[3] + 2 and _bb[3] < _foot - 2]
                _free = _foot - 12 - (max(_below) + 6 if _below else _cb[3]) - 18
                if _free < 28: continue
                _nh = _ib.height + _free; _nw = _ib.width * _nh / _ib.height
                try:
                    _raw = open(os.path.join(_figdir, 'figures', _fn), 'rb').read(32); _pxw = int.from_bytes(_raw[16:20], 'big')
                except Exception: _pxw = 0
                _nw = min(_nw, _Wd - 72, (_pxw / 150 * 72) if _pxw else _Wd - 72)
                _mm = round(_nw / 72 * 25.4, 1)
                if _mm - _cur >= 5 and _mm > figw.get(_fn, 0) + 1:
                    figw[_fn] = _mm; fadd += 1
            _prev = (_foot, _bot)
    except Exception as _e:
        print('fit_fix：图宽实测跳过（%s）' % _e)
    if fadd: print('fit_fix 第 %d 遍：插图按版面空白放大 %d 张' % (n, fadd))
    # SD-130 表外「注：」段落末行只剩一两个字：这一段收紧字距，最多两级
    import fitz
    nseen = set()
    for pg in fitz.open(pdf):   # PDF 里注段落每行各成一块：按淡蓝底色块取行，再按「注：」分条
        for dr in pg.get_drawings():
            fl = dr.get('fill')
            if not fl or '%02X%02X%02X' % tuple(round(v * 255) for v in fl) != 'EEF4FB': continue
            lines = []
            for b in pg.get_text('dict', clip=dr['rect'])['blocks']:
                for ln in b.get('lines', []): lines.append((ln['bbox'][1], ''.join(sp['text'] for sp in ln['spans'])))
            notes = []
            for _, t in sorted(lines):
                if t.lstrip().startswith('注：'): notes.append([t])
                elif notes: notes[-1].append(t)
            for ls in notes:
                if len(ls) < 2: continue
                last = re.sub(r'[^\w]', '', ls[-1])
                if 0 < len(last) <= 2:
                    fk = re.sub(r'[^\w]', '', ''.join(ls).lstrip()[2:])[:12]
                    if fk and fk not in nseen and ncond.get(fk, 0) < 2: nseen.add(fk); ncond[fk] = ncond.get(fk, 0) + 1; cadd += 1
    # SD-87 标题被留下的近空白页：块索引表允许按块分页；其他表按超出比例压缩，让标题 / 导语与表同页
    from layout_measure import lonely
    ladd = 0
    for x in lonely(pdf):
        if x['kind'] != 'lead': continue
        for sg in match(x['table'], tbls):
            if sg.startswith('块主题条目'):
                if sg not in splitok: splitok.add(sg); lv.pop(sg, None); ladd += 1
            elif sg not in block and sg not in splitok:
                need = 1 if x['over'] <= 1.04 else 2 if x['over'] <= 1.10 else 3 if x['over'] <= 1.2 else 0
                if need and lv.get(sg, 0) < need: lv[sg] = max(need, lv.get(sg, 0) + (1 if sg in lv else 0)); ladd += 1
    # SD-89 只有一两行的页：上一页最后一张表收紧一级（最多到 3 级），把这一两行拉回上一页
    from layout_measure import sparse
    for x in sparse(pdf):
        sgs = match(x['table'], tbls) if x['table'] else set()
        live = [sg for sg in sgs if sg not in block and sg not in splitok]
        for sg in live:
            if lv.get(sg, 0) < 3: lv[sg] = lv.get(sg, 0) + 1; ladd += 1
            else: block.add(sg); lv.pop(sg, None)
        # 表已收紧到底（或上一页没有可收紧的表）仍只剩一两行：上一页最后一个条目标题另起一页
        if not live and x.get('head') and x['head'] not in pbreak: pbreak.add(x['head']); ladd += 1
    splitok &= sigs
    # 2026-10-15（噜噜虚拟机自检报「JSON 解析错」）：不再丢掉 check_splits 的 stderr；取最后一行 JSON（PyMuPDF 旧版会把警告打到 stdout），解析不了就把真正原因报出来
    _out, _err = _cs.communicate()   # 先收断表结果再删临时目录
    _js = [l for l in _out.splitlines() if l.startswith('{')]
    try: res = json.loads(_js[-1])
    except (IndexError, ValueError):
        sys.exit('fit_fix：断表检查 check_splits.py 没有给出结果（退出码 %s）。\n  stdout 末尾：%s\n  stderr 末尾：%s'
                 % (_cs.returncode, _out.strip()[-500:] or '（空）', _err.strip()[-1500:] or '（空）'))
    shutil.rmtree(os.path.dirname(pdf), ignore_errors=True)
    add, up, gave, miss = 0, 0, 0, []
    for sp in res['splits']:
        cand = match(sp['text'], tbls)
        if not cand:
            if sp['fits']: miss.append('第 %d 页（%s…）' % (sp['page'], sp['text'][:16]))
            continue
        for s in cand - block - splitok:   # 允许按块分页的块索引表不再强制 / 压缩
            one = s in onep   # 源文件标 one-page 的超大表：可压到 4 级（7.5pt），不论超出多少都先试
            if s not in lv:
                if sp['fits']: lv[s] = 0; add += 1
                elif one: lv[s] = 1 if sp['over'] <= 1.04 else 2 if sp['over'] <= 1.10 else 3 if sp['over'] <= 1.2 else 4 if sp['over'] <= 1.35 else 5; up += 1   # 第 108 条（2026-10-20）：从够用的最低级起压，不够再逐级升（原直接 3～5 级，速查 285 页常用速度表压到约 8pt、页下空 120pt）
                elif sp['over'] <= 1.2 and not sp['multi']:
                    lv[s] = 1 if sp['over'] <= 1.04 else 2 if sp['over'] <= 1.10 else 3; up += 1
            elif lv[s] < (5 if one else MAXLV):
                lv[s] = max(lv[s] + 1, 1); up += 1                 # 强制 / 压缩后仍断开：升一级
            else:
                del lv[s]; block.add(s); gave += 1                  # 8pt 仍放不下：放弃，照常分页
    print('fit_fix 第 %d 遍：%d 页，断表 %d 处（本可整页 %d）；新增整表 %d、压缩升级 %d、放弃 %d、撤出 %d；当前压缩 %d 张；列宽调整 ±%d、放弃 %d；标题孤页处理 %d%s' % (
        n, res['pages'], len(res['splits']), sum(s['fits'] for s in res['splits']), add, up, gave, len(changed),
        sum(1 for v in lv.values() if v), wadd, wgive, ladd, ('；匹配不到：' + '、'.join(miss)) if miss else ''))
    if cadd: print('fit_fix 第 %d 遍：末行孤字收紧字距 %d 格' % (n, cadd))
    if nbw: print('fit_fix 第 %d 遍：借宽 %d 处（孤字 / 省行列向本表可借出宽度最多的列借）' % (n, nbw))
    print('fit_fix 第 %d 遍用时 %ds' % (n, time.time() - t0))
    if not (add or up or gave or changed or wadd or wgive or ladd or cadd or nbw or fadd): break
    hist.append((res['pages'], len(res['splits']), wadd + wgive, add or up or gave or changed or ladd or cadd or nbw))   # 2026-10-16：借宽的一遍不算「状态重复」
    _sd = os.path.join(os.path.dirname(dump), 'p%d' % n); os.makedirs(_sd, exist_ok=True)   # 每遍留一份 docx / PDF，循环时挑好的那遍
    snap[n] = (wadd + wgive + cadd, shutil.copy2(OUT, os.path.join(_sd, 'o.docx')), shutil.copy2(last_pdf, os.path.join(_sd, 'o.pdf')))
    _td = os.path.join(os.path.dirname(OUT), 'tbl_dump_' + os.path.splitext(os.path.basename(OUT))[0] + '.jsonl')   # 2026-10-16：表格排版导出跟着这一遍留一份
    if os.path.exists(_td): shutil.copy2(_td, os.path.join(_sd, 'tbl.jsonl'))
    if len(hist) >= 3 and not hist[-1][3] and hist[-1][:2] == hist[-3][:2] and abs(hist[-1][2] - hist[-3][2]) <= max(3, hist[-3][2] // 10):
        best = min((n - 1, n), key=lambda q: snap[q][0])   # 两种状态取待调整更少的一遍作成品
        if best != n:
            shutil.copy2(snap[best][1], OUT); shutil.copy2(snap[best][2], last_pdf)
            if os.path.exists(os.path.join(os.path.dirname(snap[best][1]), 'tbl.jsonl')): shutil.copy2(os.path.join(os.path.dirname(snap[best][1]), 'tbl.jsonl'), _td)
        print('fit_fix：第 %d 遍与第 %d 遍结果相同（页数、断表、列宽调整量），列宽在两种状态间来回跳，提前停止；成品取第 %d 遍（待调整 %d 处）' % (n, n - 2, best, snap[best][0]))
        trace({'best': best, 'why': 'two-state'}); break   # 2026-10-07 实测：±60 / ±25 两遍一循环，跑满 10 遍也一样
    calm = not (add or up or gave or changed or ladd or cadd or nbw) and wadd + wgive <= 3   # 只剩零星列宽微调（借宽的一遍不算平稳）
    if calm and prev_calm:
        print('fit_fix：连续两遍只有零星列宽微调（≤ 3 处），提前停止'); break
    prev_calm = calm
    if n == 7: print('fit_fix：7 遍仍未稳定，docx 保持本遍结果，交 check_layout 报告')
save()
# 交出最后一遍的 PDF（与 docx 同一遍生成），sync.sh / quickref.py 见到标记就不再重转一次（2026-10-07，省一次整本转换）
if last_pdf and os.path.exists(last_pdf):
    os.replace(last_pdf, os.path.splitext(OUT)[0] + '.pdf')
    open(OUT + '.pdf-ok', 'w').write(str(int(os.path.getmtime(OUT))))
