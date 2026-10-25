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
PROF = os.environ.get('LO_PROFILE') or os.path.join(os.environ.get('TMPDIR', '/tmp'), 'lo-sync-profile')   # E6（2026-10-21）：两本并行构建时各用一个 LibreOffice profile
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
        pass   # 2026-10-22（SD-223）：条目标题另起一页只在本次运行内有效——旧记录在内容 / 分页变了以后成了多余分页、留白
    elif l.startswith('B:') or l.startswith('SB:'):
        pass   # SD-218：整表下移 / 大表拆开只在本次运行内有效（分页位置随内容变）
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
    elif l.startswith('!'): pass   # 2026-10-22（SD-223）：「已放弃整页」只在本次运行内有效——旧记录让表在新规则下被跳过、卡在拆开状态（1021R4 六张）
    elif l.startswith('WZ:'): pass   # 2026-10-23（SD-232 ⑨）：列宽冻结只在本次运行内有效
    elif l.startswith('Z~'): pass   # 2026-10-21：一页放得下的表的收紧级只在本次运行内有效
    elif l.startswith('~'): lv[l[3:]] = 0   # 2026-10-22（SD-229）：点名整页表的压缩级也只在本次运行内有效（旧 ~5 让 1.2 自动增压表一开始就压到 7.5pt、和新插图同页拆开）
    else: lv[l] = 0
MINR = int(os.environ.get('SPLIT_MINROWS', '3'))   # 2026-10-22：大表续排每截至少 3 行（取代 SD-218 的 1 行）
BIGT = float(os.environ.get('SPLIT_BIG', '0.5'))
CLVMAX = int(os.environ.get('SHRINK_MAX', '6'))   # 2026-10-22 用户：「表格的压缩底线压到 7pt」——6 级 = 7pt   # 2026-10-22 用户：「适当缩小大表格以及字号和行间距，尽量放在同一页」——下截只差一点的表收紧（1 级行距边距、2 级 8.5pt、3 级 8pt 底线）   # SD-218：每截至少几行数据（用户定 1）
brk, sbk = set(), {}   # SD-218：本次运行内的整表下移（B:）与大表拆开位置（SB:）
clv = {}   # 2026-10-21：一页放得下、下截只差一点的表——本次运行内收紧级（1 = 收行距边距，2 = 8.5pt），仍放不下就整表移页
SHRINK_F = [1.0, 0.90, 0.85, 0.77, 0.68, 0.57, 0.50]   # 2026-10-23：4 级改 8pt 收行距（原 7.5pt 0.64），实测后校准   # 2026-10-22：各压缩级表高约为正文字号时的比例（0 级 9pt → 6 级 7pt），实测估计
def lvl_for(r):
    """整表要缩到原高的 r 倍才放得下：返回最轻的那一级（留 3% 余量）；7pt 也不够返回 None"""
    for L in range(1, CLVMAX + 1):
        if SHRINK_F[L] <= r * 0.97: return L
    return None
trial, locked, dtried = {}, set(), set()   # 2026-10-22：压缩后试退一级（trial：签名 → 原级）；退了又拆开的锁定原级
pulled = set()   # 2026-10-22：移页复查拉回过的表（每张只拉回一次）
bigsplit = set()   # 2026-10-22：压到 7.5pt 仍放不下一页的特大表（本次运行内只整行续排）
repulled = set()   # 2026-10-22：再拉回过一次的表
ipdone, ipfail = {}, set()   # 2026-10-22 条目拉回：已试过的表（签名 → 试前压缩级）、拉回失败已恢复的表
c4 = set()   # 2026-10-22 方案 C：按大表续排、每截 ≥ 4 行的表
_lvprev = {}   # 2026-10-22：上一遍各表压缩级（级别变了就作废收窄锁定）
pbfree = set()   # 2026-10-22：本次运行撤销过「另起一页」的条目标题（不再加回，防来回）
roomy = set()   # 2026-10-22：正文字号连注宽裕拉回的表（注掉页时不再整表移走）
notepull = set()   # 2026-10-22：为把表后一两行拉回本页而收紧的表（2026-10-23 起同样参加每遍试退一级）
nobrk = set()   # 2026-10-22：撤销过整表移页的表（空白页自愈），本次运行不再强制移页
sbfree = set()   # 2026-10-21：本次运行里撤销过强制拆开的表——不再对它强制拆开（防来回）
def save():
    with open(KF, 'w', encoding='utf-8') as f:
        for sg in sorted(brk): f.write('B:' + sg + '\n')
        for sg, ks in sorted(sbk.items()):
            for k in sorted(ks): f.write('SB:%d:%s\n' % (k, sg))
        # 2026-10-22：点名整页表的 lv 与本次运行的 clv 同一张表都有时取较重的一级（原来后写的 Z~ 盖掉 lv，one-page 表压不下去被放弃）
        for s, v in sorted(lv.items()): v = max(v, clv.get(s, 0)); f.write((('~%d:' % v) if v else '') + s + '\n')
        for s, v in sorted(clv.items()):
            if s in bigsplit: continue   # 2026-10-23（SD-232）：特大表＝恢复 9pt 续排，不得同时带压缩级（1023R9「引气故障」表 8pt 被拆开）
            if s not in lv: f.write('Z~%d:%s\n' % (v, s))
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
        for (sg, k) in sorted(wfrozen): f.write('WZ:%d:%s\n' % (k, sg))   # SD-232 ⑨：列宽冻结（检查器据此把 T13 / T8 降为提示）

bg = lambda x: {x[k:k + 2] for k in range(len(x) - 1)}
def match(text, tbls):
    """返回最匹配的表签名集合（同文表分数相近的一并返回）；匹配不上返回空集。"""
    want = bg(text)
    if not want: return set()
    score = sorted(((len(want & t['bg']) / len(want), t['sig']) for t in tbls), reverse=True)
    if not score: return set()
    if score[0][0] < 0.75:
        # 2026-10-22：PDF 上把紧跟表后的注框也认成表的一部分时，截到的文字比源表多、覆盖率偏低——改看「源表文字有多少出现在截到的文字里」（速查 RVR 表）
        rev = sorted(((len(want & t['bg']) / max(1, len(t['bg'])), t['sig']) for t in tbls), reverse=True)
        if rev and rev[0][0] >= 0.9 and (len(rev) < 2 or rev[0][0] - rev[1][0] >= 0.2): return {rev[0][1]}
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
MAXPASS = int(os.environ.get('FITFIX_MAXPASS', '14'))   # 2026-10-22 SD-223：整表同页 + 移页复查需要多几遍收敛（10 → 14）   # 预览（preview.py）用 4 遍
seen_state = {}   # 2026-10-07 提前停止：记录每遍开工前的状态，状态重复＝来回拉锯，再跑也只是循环（用户：「为什么每次都这么久」）
last_pdf = None
prev_calm = False
split_pending = False
hist = []
snap = {}
nspl = {}   # 每遍断表数（两种状态时挑成品用）
_clvprev = {}   # 2026-10-23（SD-232）：上一遍开工时的压缩级
# 2026-10-23（SD-232 ⑨，用户：「不要让排版来回跳」；1023R12 完整版 1.10「系统」列：第 6 遍加宽 400 → 第 7 遍收窄 954 → 第 8 遍借入 230 → 第 9 遍收窄 784……）：
#   加宽 W / F、收窄 NW、借宽 BW 各有各的放弃记录，彼此不知道对方做过什么，轮流改同一列、后一遍推翻前一遍。按列记方向：一列的净调整量换向两次（真来回跳），
#   就冻结在当时的记录，此后每遍开工前强制恢复，三种机制都不再动它；冻结写进 keep_force（WZ:）与《列宽冻结记录》，检查器对冻结列的 T13 / T8 只提示
wfrozen, wseq = {}, {}
def _wnet():
    d = {}
    for (sg, k), dd in list(wfix.items()) + list(ffix.items()): d[(sg, k)] = d.get((sg, k), 0) + dd
    for (sg, k), dd in nwfix.items(): d[(sg, k)] = d.get((sg, k), 0) - dd
    for (sg, k, j), dd in bw.items(): d[(sg, k)] = d.get((sg, k), 0) + dd; d[(sg, j)] = d.get((sg, j), 0) - dd
    return d
def _wimpose(sg, k, sn):
    for _d, _n in ((wfix, 'W'), (nwfix, 'NW'), (ffix, 'F')):
        if sn[_n] is None: _d.pop((sg, k), None)
        else: _d[(sg, k)] = sn[_n]
    for key in [key for key in bw if key[0] == sg and k in key[1:] and key not in sn['BW']]: del bw[key]
    bw.update(sn['BW'])
    wblock.add((sg, k)); nwlock.add((sg, k)); bwblock.add((sg, k))
_wprev = _wnet()   # 起点＝读进来的记录（不是 0：否则第一次清掉过时记录也算一次换方向）
_wsigs = None   # 上一遍实际排出的表（只跟踪这些表的列）
kfmap, shipped = {}, None   # 2026-10-23（SD-232）：每遍开工时的 keep_force 副本（＝排出这一遍成品的记录）；shipped＝成品取的是哪一遍
for n in range(1, MAXPASS + 1):   # 2026-10-06：7 → 10 遍（列宽规则多了，收敛要多几遍）
    # 2026-10-23（SD-232，用户：「不要让排版来回跳」）：试退过的表上一遍又被压回（拆表 / 注掉页 / 别处版面变了），锁定在新级别——
    #   每张表的压缩级在一次运行里最多转一次向，迭代必然收敛；锁定后版面变了能否再退，由收尾定稿退级逐张实测
    for _s in bigsplit: clv.pop(_s, None); trial.pop(_s, None)   # 特大表不带压缩级、不留试退记录（防别处用旧试退记录把它恢复成压缩级）
    for _s in dtried:
        if clv.get(_s, 0) > _clvprev.get(_s, 0): locked.add(_s)
    _clvprev = dict(clv)
    _wn = _wnet()
    for _key in set(_wn) | set(_wprev):
        if _key in wfrozen or _wn.get(_key, 0) == _wprev.get(_key, 0) or (_wsigs is not None and _key[0] not in _wsigs): continue
        wseq.setdefault(_key, [_wprev.get(_key, 0)]).append(_wn.get(_key, 0))
        _sg = [1 if y > x else -1 for x, y in zip(wseq[_key], wseq[_key][1:]) if y != x]
        # 换向 ≥ 2 次才算来回跳（A→B→A 是「试了又撤回」，原有的 NWL / WB / BWB 已管；1023R13 一换向就冻结，冻了 154 + 135 列，210 列是这种一次性试探）
        # 冻结在「不做调整」的状态（定宽表＝源文件宽度，其余＝生成器自己的分配）——冻结在中间状态＝把一次掷硬币的结果定住（1023R14 航向道 / 下滑道表 9/47/44 定宽被冻成 9/58/33）
        if sum(1 for x, y in zip(_sg, _sg[1:]) if x != y) >= 2: wfrozen[_key] = {'W': None, 'NW': None, 'F': None, 'BW': {}}; print('fit_fix 第 %d 遍：列宽来回调整、冻结——%s… 第 %d 列（净调整 %s DXA）' % (n, _key[0][:16], _key[1] + 1, ' → '.join(str(x) for x in wseq[_key])))
    for _key, _sn in wfrozen.items(): _wimpose(_key[0], _key[1], _sn)
    _wprev = _wnet()
    save()
    kfmap[n] = shutil.copy2(KF, os.path.join(os.path.dirname(dump), 'kf_p%d.txt' % n))
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
    _cmpd = {s0 for s0 in set(lv) | set(clv) if max(lv.get(s0, 0), clv.get(s0, 0)) > 0}   # 2026-10-22：压缩中的表——定宽表也按实测收窄（生成器压缩后套用 NW）
    for _d in (wfix, nwfix, ffix):   # 定宽表名下生成器不认的旧记录撤掉（满宽：W / NW / F；窄：NW / F）
        for _k in [k2 for k2 in _d if (k2[0] in _full or (_d is not wfix and k2[0] in _narrow)) and not (_d is nwfix and k2[0] in _cmpd)]: del _d[_k]
    # 2026-10-22：本遍压缩级与上一遍不同的表——字号变了，原来的「收窄后又折行」锁定与行数基准作废（压缩表右侧大片空白收不掉，1021R4 T13 二十余处）
    _lvnow = {s: max(lv.get(s, 0), clv.get(s, 0)) for s in set(lv) | set(clv)}
    for _s in {s for s in set(_lvnow) | set(_lvprev) if _lvnow.get(s, 0) != _lvprev.get(s, 0)}:
        for _k in [k2 for k2 in nwlock if k2[0] == _s]: nwlock.discard(_k)
        for _k in [k2 for k2 in nlbase if k2[0] == _s]: del nlbase[_k]
    _lvprev = _lvnow
    sigs = {t['sig'] for t in tbls}; _wsigs = sigs
    onep = {t['sig'] for t in tbls if t.get('onepage')}
    changed = [s for s in list(lv) if s not in sigs]            # 表已改动：撤出
    for s in changed: del lv[s]
    for s in [s for s in list(lv) if s not in onep]: del lv[s]   # SD-218：只有点名整页的表压缩；旧的整表 / 压缩记录撤掉
    block &= onep
    block &= sigs
    for key in [k2 for k2 in wfix if k2[0] not in sigs]: del wfix[key]
    for key in [k2 for k2 in nwfix if k2[0] not in sigs]: del nwfix[key]
    for key in [k2 for k2 in ffix if k2[0] not in sigs]: del ffix[key]
    for key in [k2 for k2 in bw if k2[0] not in sigs]: del bw[key]
    bwblock = {k2 for k2 in bwblock if k2[0] in sigs}
    nwlock = {k2 for k2 in nwlock if k2[0] in sigs}
    wblock = {k2 for k2 in wblock if k2[0] in sigs}
    pdf = to_pdf(); last_pdf = os.path.splitext(OUT)[0] + '.fitfix.pdf'; shutil.copy2(pdf, last_pdf); shipped = n   # 测量后临时目录会删，先留一份
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
            if (fx == 'full' or (fx == 'narrow' and (x.get('nw') or x.get('fill')))) and not (x.get('nw') and sg in _cmpd): continue   # SD-204：origin（定宽起始）照常收窄 / 拉满
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
        _SELFDRAWN = {}   # 自绘图（figures_src/figs.py 的 @fig 图号）→ 字号系数 ZOOM
        try:
            _ft = open(os.path.join(os.getcwd(), 'figures_src', 'figs.py'), encoding='utf-8').read()
            _zm = dict((k, float(v)) for k, v in re.findall(r"'(\d+\.\d+-\d+)':\s*([\d.]+)", (re.search(r'ZOOM\s*=\s*\{[^}]*\}', _ft) or re.search('', '')).group(0)))
            for _k in re.findall(r"@fig\('(\d+\.\d+-\d+)'\)", _ft): _SELFDRAWN[_k] = _zm.get(_k, 1.0)
        except Exception: pass
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
                    _flo = 0.7 * _base   # 2026-10-23 SD-231（G5「P06 ↔ P05」）：自绘图缩到图中正文字（画布约 16px × 字号系数）不小于 9pt 为止
                    _fk = re.match(r'(\d+\.\d+-\d+)_', _fn)
                    if _fk and _fk.group(1) in _SELFDRAWN:
                        try:
                            _px = int.from_bytes(open(os.path.join(_figdir, 'figures', _fn), 'rb').read(32)[16:20], 'big')
                            if _px: _flo = max(_flo, 9 * (_px / 3) * 25.4 / (72 * 16 * _SELFDRAWN[_fk.group(1)]))
                        except Exception: pass
                    if _pfree >= (_flo / _base) * _ib.height and figw.get(_fn, _base) >= _base - 0.5 and _flo < _base:   # 1020R1：只缩一次；2026-10-23 用户（速查图 1.15-1「能不能缩小一点放在前一页下面」）：下限由 80% 放到 70%——只在图会独占下一页顶部时缩
                        _mm = round(max(_flo, _cur * min(1.0, _pfree / _ib.height) * 0.97), 1)
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
            elif sg in onep and sg not in block and sg not in splitok:
                need = 1 if x['over'] <= 1.04 else 2 if x['over'] <= 1.10 else 3 if x['over'] <= 1.2 else 0
                if False and need and lv.get(sg, 0) < need: lv[sg] = max(need, lv.get(sg, 0) + (1 if sg in lv else 0)); ladd += 1   # SD-229：不为让标题与表同页压缩（表连标题整体移页）
    # SD-89 只有一两行的页：上一页最后一张表收紧一级（最多到 3 级），把这一两行拉回上一页
    from layout_measure import sparse
    for x in sparse(pdf):
        sgs = match(x['table'], tbls) if x['table'] else set()
        live = [sg for sg in sgs if sg in onep and sg not in block and sg not in splitok]
        live = []   # SD-229：点名整页表也不为拉注压缩
        # SD-218：非点名整页的表——把它最后一行挪到下一页，和这一两行（通常是表后注）作伴，不留孤页
        nd = {t['sig']: len(t.get('rows') or []) or t.get('ndata', 0) for t in tbls}
        moved = False
        for sg in sgs - onep - splitok:
            if sg in trial: clv[sg] = trial.pop(sg); locked.add(sg); ladd += 1; moved = True; continue   # 试退一级后注又掉页：恢复原级并锁定
            if sg not in brk and sg not in bigsplit and sg not in nobrk and sg not in roomy: brk.add(sg); ladd += 1; moved = True   # SD-229：表后注跟表走——不压缩，表连注整体移到下一页
            elif sg in brk and clv.get(sg, 0) < CLVMAX and x.get('th', 0) / SHRINK_F[clv.get(sg, 0)] + x.get('sh', 0) + 24 >= 0.97 * x.get('area', 1): clv[sg] = clv.get(sg, 0) + 1; ladd += 1; moved = True; notepull.add(sg)   # SD-229：只限表＋注 9pt 整页放不下   # 表已在页首、注仍掉页：表＋注作一个整体（SD-229 注跟表走）——放不下一页就逐级压到刚好放得下（最低 7pt）   # 2026-10-22：同 G6 底线 7pt、逐级只到刚好拉回（8pt 仍拉不回两处）   # 2026-10-22：已整表移到页顶的表也可收紧（原来跳过，节末注成孤页三处）   # 表后注拉回只收到 8pt   # 2026-10-21：不再把表最后一行拆下去作伴（拆表），改为收紧这张表把一两行拉回本页
        # 表已收紧到底（或上一页没有可收紧的表）仍只剩一两行：上一页最后一个条目标题另起一页
        if not live and not moved and x.get('head') and not x.get('noteonly') and x['head'] not in pbreak and x['head'] not in pbfree: pbreak.add(x['head']); ladd += 1   # 2026-10-22：节末注掉页不把条目标题另起一页（会留更大的空）
    # 2026-10-22 条目拉回（用户：结构性减少留白「第四条选 7pt」）：上一页页底留白 ≥ 15%、下一页以条目（标题＋表）开头——
    #   把这条的第一张表逐级压到刚好放得进上一页的空位（最低 7pt），整条拉回；只压到刚好放得下，不参加「试退一级」
    ipull = 0
    try:
        if not os.environ.get('FITFIX_ITEMPULL'): raise StopIteration   # SD-229：不为填空白拉回压缩（停用）
        import pymupdf as _pm
        _dd = _pm.open(pdf); _H1 = _dd[0].rect.height; _T1, _B1 = 36, _H1 - 50
        _hre = re.compile(r'^\d{1,3}\.\s')
        def _bot(p):
            ys = [b[3] for b in p.get_text('blocks') if b[4].strip() and not re.search(r'第\s*\d+\s*页\s*$', b[4].strip())]
            ys += [x['bbox'][3] for x in p.get_image_info() if (x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 4000]
            return max(ys) if ys else _T1
        for _i in range(1, len(_dd) - 1):
            _gap = _B1 - 8 - _bot(_dd[_i])
            if _gap < 0.15 * (_B1 - _T1): continue
            _nx = _dd[_i + 1]
            _bl = sorted([b for b in _nx.get_text('blocks') if b[4].strip() and not re.search(r'第\s*\d+\s*页', b[4])], key=lambda b: b[1])
            if not _bl or not _hre.match(_bl[0][4].strip()): continue
            _tb = [t for t in _nx.find_tables().tables if t.bbox[1] > _bl[0][1] and t.bbox[3] - t.bbox[1] > 10]
            if not _tb: continue
            _t0 = _tb[0]
            if any(_hre.match(b[4].strip()) and _bl[0][1] < b[1] < _t0.bbox[1] for b in _bl[1:]): continue   # 表不属于这一条
            _hd = _t0.bbox[1] - _bl[0][1]
            for _sg in match(re.sub(r'\s+', '', _nx.get_text(clip=_t0.bbox)), tbls) - onep - splitok:
                if _sg in ipfail: continue
                if _sg in trial: clv[_sg] = trial.pop(_sg); locked.add(_sg); ipull += 1; continue   # 试退一级后这一条又掉回下一页：恢复原级并锁定
                if _sg in ipdone:   # 试过拉回、这一条仍在下一页页首：拉回失败——恢复原压缩级，不再试（不留压小了却没拉回的表）
                    _p0 = ipdone.pop(_sg)
                    if _p0: clv[_sg] = _p0
                    else: clv.pop(_sg, None)
                    notepull.discard(_sg); ipfail.add(_sg); ipull += 1; continue
                _cur = clv.get(_sg, 0)
                _b0 = (_t0.bbox[3] - _t0.bbox[1]) / SHRINK_F[_cur]
                if _gap * 0.97 - _hd <= 0: continue
                _L = lvl_for((_gap * 0.97 - _hd) / _b0) if _b0 > 0 else None
                if _L is not None and _L > _cur:
                    ipdone[_sg] = _cur; clv[_sg] = _L; brk.discard(_sg); notepull.add(_sg); ipull += 1
    except StopIteration: pass
    except Exception as e:
        print('fit_fix：条目拉回出错（跳过）：%s' % e)
    if ipull: print('fit_fix 第 %d 遍：条目拉回——上一页空位放得下（压到刚好，最低 7pt），拉回 %d 条' % (n, ipull))
    splitok &= sigs
    # 2026-10-15（噜噜虚拟机自检报「JSON 解析错」）：不再丢掉 check_splits 的 stderr；取最后一行 JSON（PyMuPDF 旧版会把警告打到 stdout），解析不了就把真正原因报出来
    _out, _err = _cs.communicate()   # 先收断表结果再删临时目录
    _js = [l for l in _out.splitlines() if l.startswith('{')]
    try: res = json.loads(_js[-1])
    except (IndexError, ValueError):
        sys.exit('fit_fix：断表检查 check_splits.py 没有给出结果（退出码 %s）。\n  stdout 末尾：%s\n  stderr 末尾：%s'
                 % (_cs.returncode, _out.strip()[-500:] or '（空）', _err.strip()[-1500:] or '（空）'))
    # 2026-10-21（用户：节末尾行收回上一页）：复查本次运行里早先做的强制拆开（SB）——后面几遍别处调整、内容前移后，
    # 当初的原因（页底放不下 / 表后注成孤页）可能已经没了，强制拆开反而留下只有表尾的半空页。
    # 拆开后下一页只剩这截表尾（下面没有别的内容），而上一页的空位放得下它（续页表头不算）→ 撤销，本次运行不再对这张表强制拆开。
    unsb = 0
    try:
        import pymupdf
        _d = pymupdf.open(pdf)
        def _span(pg):
            bl = [b for b in pg.get_text('blocks') if b[4].strip()]
            ft = [b for b in bl if re.search(r'第\s*\d+\s*页\s*$', b[4].strip())]
            fy = min((b[1] for b in ft), default=pg.rect.height - 36)
            body = [b for b in bl if b[3] < fy - 2]
            return (min((b[1] for b in body), default=fy), max((b[3] for b in body), default=0), fy)
        for sp in res['splits']:
            if sp['page'] >= len(_d): continue
            for s in (match(sp['text'], tbls) & set(sbk)) - sbfree:
                _, a_bot, a_fy = _span(_d[sp['page'] - 1]); b_top, b_bot, _ = _span(_d[sp['page']])
                tail = b_bot - b_top - 20                           # 续页表头约 20pt，收回后不再重复
                if b_bot < 0.45 * _d[0].rect.height and a_fy - 8 - a_bot >= tail:
                    sbk.pop(s, None); sbfree.add(s); unsb += 1
    except Exception as e:
        print('fit_fix：复查强制拆开出错（跳过）：%s' % e)
    if unsb: print('fit_fix 第 %d 遍：撤销过时的强制拆开 %d 处（表尾收回上一页）' % (n, unsb))
    # 2026-10-22（V3 全书第 253 页空白页）：整表移页的段前分页碰上「上一页正好排满」会多出一张空白页——空白页后的第一张表如果是强制移页的，撤销移页（它本来就落在新页顶），本次运行不再对它强制移页
    unbrk = 0
    try:
        for _i in range(1, len(_d) - 1):
            _b0 = [b for b in _d[_i].get_text('blocks') if b[4].strip() and not re.search(r'第\s*\d+\s*页\s*$', b[4].strip())]
            if _b0 or any((x['bbox'][2] - x['bbox'][0]) * (x['bbox'][3] - x['bbox'][1]) > 20000 for x in _d[_i].get_image_info()): continue
            try: _t1 = [t for t in _d[_i + 1].find_tables().tables if t.bbox[3] - t.bbox[1] > 10]
            except Exception: _t1 = []
            if not _t1: continue
            for _sg in match(re.sub(r'\s+', '', _d[_i + 1].get_text(clip=_t1[0].bbox)), tbls) & brk:
                brk.discard(_sg); nobrk.add(_sg); unbrk += 1
    except Exception as e:
        print('fit_fix：空白页自愈出错（跳过）：%s' % e)
    if unbrk: print('fit_fix 第 %d 遍：空白页自愈——撤销整表移页 %d 处' % (n, unbrk))
    # 2026-10-22（1021R4：留白页后面的表 8pt 以内本可拉回上一页，完整版 87 / 速查 69 处）：移页复查——早先遍次判「整表移页」时页面还没排定，
    # 后来前面的表压缩 / 移走空出了位置。每遍看被移到页顶的表：上一页空位按实测在 8pt 以内放得下（连标题）就撤销移页、原地收紧（每张只拉回一次）
    pull = 0
    try:
        _H0 = _d[0].rect.height; _TOP0, _BOT0 = 36, _H0 - 50
        # 2026-10-24（SD-235 排版来回跳治理 A）：留白与「开头那块」的高度改用 check_gaps.room()——与 K-B5 同一个量法、同一个余量 FIT，
        #   原来本段从页顶 36pt 量、要 3% 余量，check_gaps 从首块顶量、5% 余量，1024R6 第 459 页差 0.03pt：这里不拉回、那边报留白页
        from check_gaps import room as _room, FIT as _FIT
        for _i in range(1, len(_d) - 1):
            _rm = _room(_d, _i, 0.12)
            if not _rm: continue
            _gap, _needn, _tb = _rm
            # 2026-10-22：PDF 上表连表后注认成一块、文字带「●」，签名覆盖率不到 0.5 匹配不上（速查第 23 页等四处一直没拉回）——去圆点、截掉注再匹配
            _tx = re.sub(r'[●•]', '', re.sub(r'\s+', '', _d[_i + 1].get_text(clip=_tb)))
            _tx = re.split(r'注[:：]|公司差异[:：]', _tx)[0]
            # 下一页以条目标题开头、标题有「另起一页」记录（P:）而上一页放得下：撤销（记录过时，速查第 425 条）
            _hl = [b for b in _d[_i + 1].get_text('blocks') if b[4].strip() and b[1] >= _TOP0 - 4]
            _hd = re.sub(r'\s+', '', min(_hl, key=lambda b: b[1])[4].split('\n')[0]) if _hl else ''
            if _hd and _needn <= _FIT * _gap:
                for _h in [h for h in pbreak if _hd.startswith(h) or h.startswith(_hd)]:
                    pbreak.discard(_h); pbfree.add(_h); pull += 1
            for _sg in match(_tx, tbls) & brk:
                _roomy = _needn <= 0.9 * _gap and not clv.get(_sg, 0)   # 正文字号连注都宽裕放得下：不受「只拉回一两次」限制
                if _sg in pulled and not _roomy and (_gap < 0.5 * (_BOT0 - _TOP0) or _sg in repulled): continue   # 2026-10-22：上一页空过半时可再拉回（移页记录过时，速查第 355 条整页只剩三行）
                _r = _gap / max(1, _needn) * SHRINK_F[clv.get(_sg, 0)]
                _L = 0 if _r >= 1 / _FIT else lvl_for(_r)
                _cur = clv.get(_sg, 0)
                # SD-229：9pt 的表只在正文字号就放得下时拉回。2026-10-23（SD-232 ⑧，1023R11 第 543 页留白 406pt：4 级的表在下一页页顶、按现有级别放得进）：
                #   压缩过的表（级别由「一整页放得下」定）按现有级别实测放得进上一页空位就拉回，级别不变——不为填空白加压，也不按估算减压；与 K-B5 check_gaps 同一个「放得下」
                if (_cur == 0 and _L == 0) or (_cur > 0 and _needn <= _FIT * _gap):
                    if _sg in pulled: repulled.add(_sg)   # 2026-10-22：再拉回只给一次（防移页 / 拉回来回拉锯，1021R4 重建十三小表拆开）
                    brk.discard(_sg); pulled.add(_sg); pull += 1
                    if _roomy: roomy.add(_sg)   # 宽裕拉回的表本次运行不再因注掉页整表移走（真拆开了仍照 SD-229 整表移页）
    except Exception as e:
        print('fit_fix：移页复查出错（跳过）：%s' % e)
    if pull: print('fit_fix 第 %d 遍：移页复查——上一页空位放得下，拉回 %d 张表' % (n, pull))
    # 2026-10-22：压缩只到「刚好放得下」——本遍没拆开的已压缩表试退一级（每张只试一次），退了又拆开就恢复并锁定
    _splitsig = set()
    for _sp in res['splits']: _splitsig |= match(_sp['text'], tbls)
    # 2026-10-23（SD-232）：每遍留痕——各压缩表的状态与每处断表匹配到的表（键名不用 pass / W，fixosc.py 不读）
    trace({'state_pass': n, 'tables': {_s[:40]: {'lv': clv.get(_s, 0), 'brk': _s in brk, 'big': _s in bigsplit, 'trial': trial.get(_s), 'lock': _s in locked}
                                       for _s in set(clv) | bigsplit | set(trial) | locked},
           'splits': [[_sp['page'], _sp.get('csize'), sorted(_x[:40] for _x in match(_sp['text'], tbls))] for _sp in res['splits']]})
    tdown = 0
    # 2026-10-23（SD-232，用户：「不要让排版来回跳，否则每改动一次都会来回跳，就会一直做重复工」）：删去每遍「按估算直接退级」——
    #   它按比例估算、还把锁定过的表解锁再退，退了注掉页、下遍又压回，是来回跳的来源。每遍只留下面的真排试退一级（拆表 / 注掉页即恢复并锁定），
    #   锁定之后版面变了的表由收尾「定稿退级」按 scan_over 候选逐张实测重试
    for _s in [x for x in trial if x not in _splitsig]: trial.pop(_s)   # 2026-10-22：退一级没拆开＝成功，下遍可再退（压缩只到刚好放得下）
    for _s in list(clv):
        if _s in _splitsig or _s in locked or _s in trial: continue   # 2026-10-23（用户：速查第 329 页「为什么这一页还存在过度压缩的情况？明明下面还有空间」）：拉回表后注用的收紧也参加试退一级（原来排除，压到 7.5pt 退不回来）；退了注又掉页由拉回检查恢复并锁定   # 2026-10-22：拉回用的收紧也试退一级（速查第 25 页：空间早已腾出、表仍压着）；退了又掉页由拉回检查恢复并锁定   # 2026-10-22：为拉回表后注而收紧的表不试退（退了注又掉成孤页，1021R4 重建五三处）
        trial[_s] = clv[_s]; dtried.add(_s); tdown += 1
        if clv[_s] > 1: clv[_s] -= 1
        else: clv.pop(_s)
    if tdown: print('fit_fix 第 %d 遍：已压缩的表试退一级 %d 张（只压到刚好放得下）' % (n, tdown))
    if bigsplit: print('fit_fix 第 %d 遍：压到 7pt 仍放不下一页、整行续排的特大表 %d 张' % (n, len(bigsplit)))
    shutil.rmtree(os.path.dirname(pdf), ignore_errors=True)
    add, up, gave, miss, stuck = 0, 0, 0, [], 0
    for sp in res['splits']:
        one_hit = match(sp['text'], tbls) & onep
        if not sp.get('orphan') and not one_hit and sp.get('over', 9) > 1.6: continue   # 2026-10-22：超出一页不多（≤ 45%）的大表也处理——压缩后放得下就整表同页          # SD-218：合格的跨页断点不动；点名整页的表任何断开都要处理
        cand = match(sp['text'], tbls)
        if not cand:
            miss.append('第 %d 页（%s…）' % (sp['page'], sp['text'][:16])); continue
        for s in cand - splitok:
            if s in block and s in onep and s not in bigsplit: clv.pop(s, None); brk.discard(s); bigsplit.add(s)   # 2026-10-22：点名整页压到 7pt 仍放不下——按特大表整行续排处理断点（原来放弃后不再管，断在合并格中间，1021R4 2.3 进近稳定性表）
            elif s in block: continue
            if s in onep and s not in block:   # 点名整页：从页首开始，仍放不下就逐级压缩（最多 7pt）
                if s not in brk: brk.add(s); add += 1
                elif max(lv.get(s, 0), clv.get(s, 0)) < CLVMAX and ((sp.get('ttop', 0) - sp.get('above', 0)) <= 0.15 * sp.get('area', 1) or (sp.get('ha', 0) + sp.get('hb', 0) - sp.get('hh', 0)) / SHRINK_F[max(lv.get(s, 0), clv.get(s, 0))] >= 0.9 * sp.get('area', 1)): lv[s] = max(lv.get(s, 0), clv.get(s, 0)) + 1; up += 1   # 2026-10-22：底线 7pt（同 G6）；SD-229：只限 9pt 整页放不下的表
                elif max(lv.get(s, 0), clv.get(s, 0)) < CLVMAX: print('fit_fix 第 %d 遍：点名整页表已移页仍拆开（9pt 放得下一页，不压缩）：第 %d 页 %s…' % (n, sp['page'], s[:16])); stuck += 1
                else: lv.pop(s, None); block.add(s); gave += 1
                continue
            if s in bigsplit and not sp.get('orphan'): continue    # 压到底仍放不下一页的特大表：合格的整行续排断点不动
            da, db = sp.get('da', 0), sp.get('db', 0)
            _tf = db / max(1, da + db)
            _ov = sp.get('over', 1)
            # 2026-10-22 用户：「对于那种跨页的大表格，可以把间隔和字体大小的界限再小一点」「表格的压缩底线压到 7pt」
            # 「压缩的标准是只要这一页刚好能放下一整个表格内容就好，而不是压缩过度到整个页还有空白区域」——按实测高度算「刚好放得下」的最轻一级
            if s not in bigsplit:
                if s in trial: clv[s] = trial.pop(s); locked.add(s); up += 1; continue   # 退一级试排又拆开了：恢复原级并锁定
                _cur = clv.get(s, 0); _f0 = SHRINK_F[_cur]
                _gen = (sp.get('ttop', 0) - sp.get('above', 0)) <= 0.15 * sp.get('area', 1)   # 2026-10-22：表（连标题）已在页顶仍拆开＝真放不下；不在页顶＝移页没生效，不能靠压缩
                _need = (sp.get('ha', 0) + sp.get('hb', 0) - sp.get('hh', 0)) / _f0      # 正文字号时整表高
                _top = sp.get('area', 0) * 0.97 - (sp.get('above', 0) if s not in brk else 0)
                _tl = 0 if _need and _top >= _need else (lvl_for(_top / _need) if _need > 0 else None)   # 移到页顶放得下的最轻级（0＝9pt 就放得下）
                # SD-229（2026-10-22 用户）：能放进一页的表一律 9pt 整表（连标题）移到下一页页顶；只有 9pt 整页放不下才压到刚好放得下（最低 7pt）；
                #   不再为填页面空白原地压缩（撤 SD-223 ①、SD-224 方案 C、条目 / 注拉回的压缩）
                if s not in brk and s not in nobrk and _tl is not None:
                    if _tl: clv[s] = _tl
                    else: clv.pop(s, None)
                    brk.add(s); add += 1; continue
                if s in brk and _tl is not None and _tl > _cur:        # 已在页顶仍放不下：按实测再压到刚好放得下
                    clv[s] = _tl; up += 1; continue
                if s in brk and _cur < CLVMAX and (_tl is None or _tl <= _cur) and (_gen or _need >= 0.9 * sp.get('area', 0)):   # 估算偏乐观：再升一级，到 7pt 为止
                    clv[s] = _cur + 1; up += 1; continue   # 2026-10-22：只限正文字号整表接近一整页的表（完整版 MCAS 表 9pt 约半页，移页没生效时被逐级压到 7.5pt）
                if s in brk and not _gen and _need and _need < 0.9 * sp.get('area', 0):   # 9pt 放得下一页却仍拆开：移页没生效，不压缩，记下待查
                    print('fit_fix 第 %d 遍：已整表移页仍拆开（9pt 约 %.0f%% 页高，不压缩）：第 %d 页 %s…' % (n, 100 * _need / max(1, sp.get('area', 1)), sp['page'], s[:16])); stuck += 1; continue
            # 2026-10-23（SD-232）：估算说 7pt 也放不下时，不直接判特大表——按实测逐级升压，压到 7pt 仍拆开才续排（G6 原话「压到 7pt 仍放不下」）。
            #   估算对段落多的表严重偏大：1023R9「引气故障」表按 8pt 拆开的高度估「7pt 也放不下」判成特大表，实测 7.5pt 整表放得下
            if s not in bigsplit and 1 <= clv.get(s, 0) < CLVMAX:   # 只限已压缩的表（9pt 实测的高度是准的，不准的是从压缩表反推的估算）
                clv[s] = clv.get(s, 0) + 1
                if s not in brk and s not in nobrk: brk.add(s); add += 1
                up += 1; continue
            # ③ 压到底仍放不下：恢复正文字号，整行续排（每截 ≥ MINR 行、续页重复表头）
            if s not in bigsplit: clv.pop(s, None); brk.discard(s); bigsplit.add(s)
            _mr = 4 if s in c4 else MINR   # 方案 C 的大表每截 ≥ 4 行
            if da < MINR and sp.get('orphan') and s not in brk and (not sp.get('multi') or sp.get('above', 0) > 20):   # 2026-10-22：跨多页的表首截（上方有标题等）只剩一两行也整表下移（速查 ND 符号表）    # 大表页底只剩表头 / 一两行：从下一页开始
                brk.add(s); add += 1
            elif db < _mr or sp.get('group_tail') or sp.get('span_cut'):   # 页顶不足 / 页底是分组行 / 断在合并格：提前拆开
                rws = next((t.get('rows') for t in tbls if t['sig'] == s and t.get('rows')), [])
                fb = sp.get('first_b', '')
                b0 = next((j for j, r in enumerate(rws) if fb and r['t'] and (r['t'].startswith(fb[:12]) or fb.startswith(r['t'][:12]))), None)   # 下截第一行在整表里的行号
                if b0 is None:   # 合并格里的长文字续到了下一页、定位不到行：整张表从下一页开始
                    if s not in brk: brk.add(s); add += 1
                    else: stuck += 1
                    continue
                k = b0
                if sp.get('span_cut'):
                    while k > 0 and rws[k]['c']: k -= 1                 # 退到合并格那一组的首行
                elif sp.get('group_tail'): k = b0 - 1                   # 分组行挪到下一页
                else: k = b0 - (_mr - db)
                if s in sbfree: stuck += 1; continue          # 2026-10-21：撤销过强制拆开的表，本次运行不再拆
                ks = sbk.setdefault(s, set())
                if k >= _mr and k not in ks: ks.add(k); up += 1
                elif (sp['fits'] or k < _mr) and s not in brk: brk.add(s); add += 1   # 组从表首开始：整表从下一页开始（放不下的大表随后在合格处断开）
                else: stuck += 1
            else:
                stuck += 1
    print('fit_fix 第 %d 遍：%d 页，断表 %d 处（本可整页 %d）；新增整表 %d、压缩升级 %d、放弃 %d、撤出 %d；当前压缩 %d 张；列宽调整 ±%d、放弃 %d；标题孤页处理 %d%s' % (
        n, res['pages'], len(res['splits']), sum(s['fits'] for s in res['splits']), add, up, gave, len(changed),
        sum(1 for v in lv.values() if v), wadd, wgive, ladd, ('；匹配不到：' + '、'.join(miss)) if miss else ''))
    if cadd: print('fit_fix 第 %d 遍：末行孤字收紧字距 %d 格' % (n, cadd))
    if nbw: print('fit_fix 第 %d 遍：借宽 %d 处（孤字 / 省行列向本表可借出宽度最多的列借）' % (n, nbw))
    print('fit_fix 第 %d 遍用时 %ds' % (n, time.time() - t0))
    split_pending = bool(add or up or ladd or unsb or unbrk or tdown or pull or ipull)   # SD-218：这一遍新记的整表下移 / 拆开还没排进成品
    if not (add or up or gave or changed or wadd or wgive or ladd or cadd or nbw or fadd or unsb or unbrk or tdown or pull or ipull): break
    hist.append((res['pages'], len(res['splits']), wadd + wgive, add or up or gave or changed or ladd or cadd or nbw))   # 2026-10-16：借宽的一遍不算「状态重复」
    _sd = os.path.join(os.path.dirname(dump), 'p%d' % n); os.makedirs(_sd, exist_ok=True)   # 每遍留一份 docx / PDF，循环时挑好的那遍
    nspl[n] = len(res['splits'])
    snap[n] = (wadd + wgive + cadd, shutil.copy2(OUT, os.path.join(_sd, 'o.docx')), shutil.copy2(last_pdf, os.path.join(_sd, 'o.pdf')))
    _td = os.path.join(os.path.dirname(OUT), 'tbl_dump_' + os.path.splitext(os.path.basename(OUT))[0] + '.jsonl')   # 2026-10-16：表格排版导出跟着这一遍留一份
    if os.path.exists(_td): shutil.copy2(_td, os.path.join(_sd, 'tbl.jsonl'))
    if len(hist) >= 3 and not hist[-1][3] and hist[-1][:2] == hist[-3][:2] and abs(hist[-1][2] - hist[-3][2]) <= max(3, hist[-3][2] // 10):
        # 2026-10-23（SD-232 ⑨，1.10 单节：第 5 遍试退拆表、第 6 遍已恢复，却因第 5 遍待调整更少取了第 5 遍）：先比断表数，再比待调整数
        best = min((n - 1, n), key=lambda q: (nspl.get(q, 0), snap[q][0]))   # 两种状态取断表更少、待调整更少的一遍作成品
        if best != n:
            shutil.copy2(snap[best][1], OUT); shutil.copy2(snap[best][2], last_pdf)
            if os.path.exists(os.path.join(os.path.dirname(snap[best][1]), 'tbl.jsonl')): shutil.copy2(os.path.join(os.path.dirname(snap[best][1]), 'tbl.jsonl'), _td)
        print('fit_fix：第 %d 遍与第 %d 遍结果相同（页数、断表、列宽调整量），列宽在两种状态间来回跳，提前停止；成品取第 %d 遍（待调整 %d 处）' % (n, n - 2, best, snap[best][0]))
        shipped = best
        trace({'best': best, 'why': 'two-state'}); break   # 2026-10-07 实测：±60 / ±25 两遍一循环，跑满 10 遍也一样
    calm = not (add or up or gave or changed or ladd or cadd or nbw or ipull or pull or tdown or unsb or unbrk) and wadd + wgive <= 10   # 只剩零星列宽微调（借宽的一遍不算平稳）；2026-10-22：列宽收尾交给定稿后收窄 / 加宽（SD-226 / 227），≤ 10 处就算平稳
    if calm and prev_calm:
        print('fit_fix：连续两遍只有零星列宽微调（≤ 3 处），提前停止'); break
    prev_calm = calm
    if n == 7: print('fit_fix：7 遍仍未稳定，docx 保持本遍结果，交 check_layout 报告')
# SD-218：循环提前停下（两种列宽状态来回跳 / 到遍数上限）时，最后一遍新记的断点修正还没排进成品——按当前记录再排一遍
# 2026-10-23（SD-232，1023R10：迭代第 10 遍实测 552 页 / 断表问题 0，交出去的却是 553 页 / 问题 2——收尾重排用的是内存里最后一遍新做、没实测过的状态）：
#   收尾只从「实测过的那一遍」出发：先把它的 docx / PDF / keep_force 留作基准；补排、定稿退级的结果都要对照基准，断表问题、留白页、页数不能变多，变多就退回基准
_BASE = None
json.dump([{'table': sg[:40], 'col': k + 1, 'seq': wseq.get((sg, k)), 'frozen': {kk: (v if kk != 'BW' else {'%d←%d' % (a[1], a[2]): dd for a, dd in v.items()}) for kk, v in sn.items()}} for (sg, k), sn in sorted(wfrozen.items())],
          open(os.path.join(os.path.dirname(OUT) or '.', '列宽冻结记录_%s.json' % name), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
def _gate(pdfp):
    """成品质量：断表问题数、留白页数、页数（越小越好，逐项不许变多）"""
    import fitz as _fzg
    _np = len(_fzg.open(pdfp))
    try: _sp = len(json.loads(subprocess.run([sys.executable, os.path.join(T, 'check_splits.py'), pdfp, '--json'], capture_output=True, text=True, timeout=600).stdout)['problems'])
    except Exception: _sp = 0
    try: _gp = sum(1 for l in subprocess.run([sys.executable, os.path.join(T, 'check_gaps.py'), pdfp], capture_output=True, text=True, timeout=600).stdout.splitlines() if l.startswith('第 '))
    except Exception: _gp = 0
    return (_sp, _gp, _np)
def _keep_base(tag):
    _d = tempfile.mkdtemp()
    return (tag, shutil.copy2(OUT, os.path.join(_d, 'o.docx')), shutil.copy2(last_pdf, os.path.join(_d, 'o.pdf')), shutil.copy2(KF, os.path.join(_d, 'kf.txt')), _gate(last_pdf))
def _to_base(base, why):
    shutil.copy2(base[1], OUT); shutil.copy2(base[2], last_pdf); shutil.copy2(base[3], KF)
    print('fit_fix：%s——退回%s（断表问题 %d、留白页 %d、%d 页）' % (why, base[0], *base[4]))
if last_pdf and os.path.exists(last_pdf):
    if shipped in kfmap: shutil.copy2(kfmap[shipped], KF)   # keep_force 与成品同一遍（内存里已带着最后一遍新做的决定）
    _BASE = _keep_base('第 %s 遍实测成品' % shipped)
if split_pending and _BASE:   # SD-218：最后一遍新记的断点修正还没排进成品——补排一遍，但不比实测成品差才收
    save(); build(dump); _p = to_pdf(); _g = _gate(_p)
    if all(x <= y for x, y in zip(_g, _BASE[4])):
        shutil.copy2(_p, last_pdf); _BASE = _keep_base('补排后的成品'); print('fit_fix：补排一遍，把最后记下的断点修正排进成品')
    else: _to_base(_BASE, '补排后断表问题 / 留白页 / 页数变多（%d / %d / %d）' % _g)
    shutil.rmtree(os.path.dirname(_p), ignore_errors=True)
# 2026-10-23 定稿退级（SD-231；用户：「为什么这一页还存在过度压缩的情况？明明下面还有空间」）：
# 迭代里「退一级」和「为拉回表后注再压回」会来回跳。收尾时对成品里仍压过头的表统一退一级重排，逐张核对：表没拆开、表后紧跟的注仍和表同页、总页数不增——不满足的恢复原级
# 2026-10-23（SD-232，用户：「过度压缩为什么会反复出现」——检查器与本段各算各的，同一份 PDF 一个认 1 张、一个报 2 张）：
#   · 候选改用 check_overshrink.scan_over()，与检查器同一个定义；
#   · 每张候选的实测结果（退没退成、哪一条不满足）逐张写 build/压缩退级记录_<书名>.json，检查器读它：实测退不下来的只出提示；
#   · 整本页数增加不再连坐整批：批量试后，只因页数增加而失败的表逐张再试一次
_OSLOG = os.path.join(os.path.dirname(OUT) or '.', '压缩退级记录_%s.json' % name)
def _oslog(rows):
    json.dump(rows, open(_OSLOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
try:
    _oslog([])   # 先清空：本轮没跑到这一步时，检查器不会读到上一轮的旧结论
    import fitz as _fz3
    sys.path.insert(0, T)
    from check_overshrink import scan_over
    def _where(pdfp, key, near):
        _d3 = _fz3.open(pdfp)
        return [i for i in range(max(0, near - 2), min(len(_d3), near + 4)) if key and key in re.sub(r'\s+', '', _d3[i].get_text())]
    def _splits(pdfp):
        _sp = set()
        try:
            _so = subprocess.run([sys.executable, os.path.join(T, 'check_splits.py'), pdfp], capture_output=True, text=True, timeout=600).stdout
            for _m in re.finditer(r'第\s*(\d+)\s*→\s*(\d+)\s*页', _so): _sp.add(int(_m.group(1)) - 1)
        except Exception: pass
        return _sp
    def _judge(pdfp, keys, np0):
        """逐张核对退级后的成品：表找得到、没拆开、紧跟的注仍同页；另返回整本页数是否增加"""
        _np = len(_fz3.open(pdfp)); _sp = _splits(pdfp); _res = {}
        for _k in keys:
            _c = _over[_k]; _a = _where(pdfp, _c['head'], _c['page'] - 1); _r = []
            _h1 = None
            if not _a: _r.append('退级后找不到这张表')
            else:
                if _a[0] in _sp: _r.append('拆表')
                if _c['note'] and _a[0] not in _where(pdfp, _c['note'], _a[0]): _r.append('注「%s」掉到下一页' % _c['note'][:6])
                try:   # 退级后实测表高（以后按记录校准各级比例，不再手量）
                    _pg = _fz3.open(pdfp)[_a[0]]
                    _h1 = next((round(t.bbox[3] - t.bbox[1]) for t in _pg.find_tables().tables if re.sub(r'\s+', '', _pg.get_text(clip=t.bbox)).startswith(_c['head'])), None)
                except Exception: pass
            _res[_k] = (_r, _a[0] + 1 if _a else None, _h1)
        return _res, _np > np0, _np
    # SD-232：只改 keep_force 文件里这几张表的压缩级（Z?~N:签名），不 save() 内存状态——其余记录与实测成品完全相同
    def _setlv(_k, _from, _to):
        _ls = open(KF, encoding='utf-8').read().split('\n'); _sgs = set(_over[_k]['sig'])
        for _i, _l in enumerate(_ls):
            _m = re.match(r'(Z?~)(\d):(.+)$', _l)
            if _m and _m.group(3) in _sgs and int(_m.group(2)) == _from: _ls[_i] = '%s%d:%s' % (_m.group(1), _to, _m.group(3))
        open(KF, 'w', encoding='utf-8').write('\n'.join(_ls))
    def _down(_k): _setlv(_k, _over[_k]['level'], _over[_k]['level'] - 1)
    def _up(_k): _setlv(_k, _over[_k]['level'] - 1, _over[_k]['level'])
    def _rebuild():
        build(dump); return to_pdf()
    if last_pdf and os.path.exists(last_pdf) and _BASE:
        from check_overshrink import read_levels
        # SD-232 ⑧（1023R11 速查第 485 页：退成一张后页面变了，后面的表成了新候选，只扫一轮就漏了）：最多 3 轮，每轮在上一轮接受的成品上重新扫描；本次实测失败过的不再试
        _allrows, _tried = [], set()
        for _round in range(3):
            _lev = read_levels(KF)
            _over, _ = scan_over(last_pdf, [MD], _lev)
            _np0 = len(_fz3.open(last_pdf))
            _rows = [{'head': c['head'], 'page': c['page'], 'from': c['level'], 'to': c['level'] - 1, 'table_h': c['table_h'], 'extra_est': c.get('extra'), 'free': c['free'], 'note': c['note'],
                      'result': 'nomatch' if not c['sig'] else '', 'reason': '匹配不到压缩记录' if not c['sig'] else ''} for c in _over]
            _over = [c for c in _over if (c['head'], c['level']) not in _tried]
            _rows = [r for r in _rows if (r['head'], r['from']) not in _tried]
            _cand = [k for k, c in enumerate(_over) if c['sig']]
            if _cand:
                for _k in _cand: _down(_k)
                _p3 = _rebuild()
                _res, _grew, _np1 = _judge(_p3, _cand, _np0)
                _local = [k for k in _cand if _res[k][0]]
                _pageonly = [k for k in _cand if not _res[k][0] and _grew]
                for _k in _cand:
                    _rows[_k]['after_h'] = _res[_k][2]
                    if _k in _local: _rows[_k].update(result='fail', reason='；'.join(_res[_k][0]))
                    elif _k not in _pageonly: _rows[_k].update(result='ok', after_page=_res[_k][1])
                for _k in _local + _pageonly: _up(_k)
                if _local or _pageonly:
                    shutil.rmtree(os.path.dirname(_p3), ignore_errors=True)
                    if not any(r['result'] == 'ok' for r in _rows): _p3 = None; shutil.copy2(_BASE[1], OUT); shutil.copy2(_BASE[3], KF)   # 一张都没退成：不重排，成品就是基准
                    else: _p3 = _rebuild()
                    _npb = len(_fz3.open(_p3))
                    _npb = len(_fz3.open(_p3 or last_pdf))
                    for _k in _pageonly:   # 只因整本页数增加失败的：逐张单独再试
                        _docx0, _kf0 = shutil.copy2(OUT, OUT + '.try'), open(KF, encoding='utf-8').read()
                        _down(_k); _p4 = _rebuild()
                        _r4, _g4, _np4 = _judge(_p4, [_k], _npb); _rows[_k]['after_h'] = _r4[_k][2]
                        if _r4[_k][0] or _g4:
                            open(KF, 'w', encoding='utf-8').write(_kf0); shutil.copy2(_docx0, OUT)   # 恢复：docx 与 keep_force 回到试之前，不重排
                            _rows[_k].update(result='fail', reason='；'.join(_r4[_k][0] + (['整本多出 %d 页' % (_np4 - _npb)] if _g4 else [])))
                            shutil.rmtree(os.path.dirname(_p4), ignore_errors=True)
                        else:
                            _rows[_k].update(result='ok', after_page=_r4[_k][1])
                            if _p3: shutil.rmtree(os.path.dirname(_p3), ignore_errors=True)
                            _p3, _npb = _p4, _np4
                        os.remove(OUT + '.try')
                if _p3:
                    _g = _gate(_p3)
                    if all(x <= y for x, y in zip(_g, _BASE[4])): shutil.copy2(_p3, last_pdf)
                    else:
                        for r in _rows:
                            if r['result'] == 'ok': r.update(result='fail', reason='退级后整本断表问题 / 留白页 / 页数变多（%d / %d / %d），整体退回' % _g)
                        _to_base(_BASE, '定稿退级后整本变差')
                    shutil.rmtree(os.path.dirname(_p3), ignore_errors=True)
                else: shutil.copy2(_BASE[2], last_pdf)
            _allrows += _rows
            _tried |= {(r['head'], r['from']) for r in _rows}
            if not any(r['result'] == 'ok' for r in _rows): break
            _BASE = _keep_base('定稿退级第 %d 轮后的成品' % (_round + 1))
        _rows = _allrows
        _oslog(_rows)
        _ok = sum(1 for r in _rows if r['result'] == 'ok')
        print('fit_fix：定稿退级——压过头的候选 %d 张，实测退一级成功 %d 张、退不下来 %d 张、匹配不到 %d 张（逐张结果见 %s）' % (
            len(_rows), _ok, sum(1 for r in _rows if r['result'] == 'fail'), sum(1 for r in _rows if r['result'] == 'nomatch'), os.path.basename(_OSLOG)))
        for r in _rows:
            if r['result'] == 'fail': print('  · 第 %d 页「%s」%d→%d 级：%s' % (r['page'], r['head'], r['from'], r['to'], r['reason']))
except Exception as _e:
    print('fit_fix：定稿退级跳过（%s）' % _e)
# 2026-10-24（SD-235 排版来回跳治理 B，用户：「咨询一下顾问怎么解决排版来回跳的问题」）：收尾「B5 归零」——
#   成品里仍有留白页（check_gaps.room，与 K-B5 同一判据）时，撤掉下一页那张表本次运行的整表移页（keep_force 的 B: 行，只删不加），
#   重排一次；断表问题与页数不变多、留白页变少才收，否则退回基准。最多 2 轮（1024R6 第 459 页曾靠手工撤这一处分页）
try:
    import fitz as _fz5
    from check_gaps import room as _room5, FIT as _FIT5
    for _round in range(2):
        if not (last_pdf and os.path.exists(last_pdf) and _BASE): break
        _dd = _fz5.open(last_pdf); _sigs, _pgs = set(), []
        for _i in range(len(_dd) - 1):
            _rm = _room5(_dd, _i)
            if not (_rm and _rm[1] <= _FIT5 * _rm[0]): continue
            _tx = re.sub(r'[●•]', '', re.sub(r'\s+', '', _dd[_i + 1].get_text(clip=_rm[2])))
            _hit = match(re.split(r'注[:：]|公司差异[:：]', _tx)[0], tbls)
            if _hit: _sigs |= _hit; _pgs.append(_i + 1)
        if not _sigs: break
        _kf0 = open(KF, encoding='utf-8').read()
        _ls = [l for l in _kf0.split('\n') if not (l.startswith('B:') and l[2:] in _sigs)]
        _nd = len(_kf0.split('\n')) - len(_ls)
        if not _nd: print('fit_fix：B5 归零——留白页 %s 的表没有整表移页记录，不动' % _pgs); break
        _docx0 = shutil.copy2(OUT, OUT + '.b5')
        open(KF, 'w', encoding='utf-8').write('\n'.join(_ls))
        build(dump); _p5 = to_pdf(); _g5 = _gate(_p5)
        if _g5[0] <= _BASE[4][0] and _g5[2] <= _BASE[4][2] and _g5[1] < _BASE[4][1]:
            _gb = _BASE[4][1]; shutil.copy2(_p5, last_pdf); _BASE = _keep_base('B5 归零第 %d 轮后的成品' % (_round + 1))
            print('fit_fix：B5 归零——撤整表移页 %d 处（第 %s 页留白），留白页 %d → %d' % (_nd, '、'.join(map(str, _pgs)), _gb, _g5[1]))
        else:
            open(KF, 'w', encoding='utf-8').write(_kf0); shutil.copy2(_docx0, OUT)
            print('fit_fix：B5 归零——撤 %d 处后断表问题 / 留白页 / 页数为 %d / %d / %d（基准 %d / %d / %d），不收，保留留白页' % ((_nd,) + _g5 + _BASE[4]))
            os.remove(OUT + '.b5'); shutil.rmtree(os.path.dirname(_p5), ignore_errors=True); break
        os.remove(OUT + '.b5'); shutil.rmtree(os.path.dirname(_p5), ignore_errors=True)
except Exception as _e:
    print('fit_fix：B5 归零跳过（%s）' % _e)
# 2026-10-22 定稿后收窄（t13_post）：按成品实测把不折行、右侧仍空 ≥ 10pt 的列直接在 docx 里收到刚好够宽——列内行数不变、分页不动
try:
    import t13_post
    if not os.environ.get('FITFIX_POST'): raise StopIteration   # 2026-10-22 用户「按以前的方式」：停用定稿后处理
    for _r in range(2):
        _nch = t13_post.run(OUT, last_pdf, MD, tbls, KF) if last_pdf and os.path.exists(last_pdf) else 0
        if not _nch: break
        _p = to_pdf(); shutil.copy2(_p, last_pdf); shutil.rmtree(os.path.dirname(_p), ignore_errors=True)
        print('fit_fix：定稿后收窄 %d 列（不折行列收到刚好够宽，分页不动）' % _nch)
except StopIteration: pass
except Exception as _e:
    print('fit_fix：定稿后收窄跳过（%s）' % _e)
# 2026-10-22 定稿后加宽 / 借宽（T14 有空位却折行、T8 末行孤字）：改完重转 PDF 逐页比对页首，分页动了的那张撤回
try:
    if not os.environ.get('FITFIX_POST'): raise StopIteration
    if last_pdf and os.path.exists(last_pdf):
        _ng, _gp = t13_post.run_grow(OUT, last_pdf, MD, to_pdf)
        if _gp: shutil.copy2(_gp, last_pdf); shutil.rmtree(os.path.dirname(_gp), ignore_errors=True)
        if _ng: print('fit_fix：定稿后加宽 / 借宽 %d 处（逐页核对页首，分页不动）' % _ng)
except StopIteration: pass
except Exception as _e:
    print('fit_fix：定稿后加宽跳过（%s）' % _e)
# 加宽 / 借宽重写 docx 后再收一遍（收窄不动分页；1022R1 重建十九：加宽后仍有 5 处列空白）
try:
    if not os.environ.get('FITFIX_POST'): raise StopIteration
    _nch = t13_post.run(OUT, last_pdf, MD, tbls, KF) if last_pdf and os.path.exists(last_pdf) else 0
    if _nch:
        _p = to_pdf(); shutil.copy2(_p, last_pdf); shutil.rmtree(os.path.dirname(_p), ignore_errors=True)
        print('fit_fix：加宽后再收窄 %d 列' % _nch)
except StopIteration: pass
except Exception as _e:
    print('fit_fix：再收窄跳过（%s）' % _e)
# 2026-10-23（SD-232）：结尾不再 save()——内存里带着最后一遍没实测的决定；keep_force 保持与交出的成品同一遍（检查器按它读压缩级）
# 交出最后一遍的 PDF（与 docx 同一遍生成），sync.sh / quickref.py 见到标记就不再重转一次（2026-10-07，省一次整本转换）
if last_pdf and os.path.exists(last_pdf):
    os.replace(last_pdf, os.path.splitext(OUT)[0] + '.pdf')
    open(OUT + '.pdf-ok', 'w').write(str(int(os.path.getmtime(OUT))))
