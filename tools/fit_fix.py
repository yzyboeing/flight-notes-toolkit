#!/usr/bin/env python3
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
lv, block, wfix, wblock, splitok, pbreak = {}, set(), {}, set(), set(), set()
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
        _, k, sg = l.split(':', 2); wblock.add((sg, int(k)))
    elif l.startswith('C:'):
        _, cl, rest = l.split(':', 2); sg, fk = rest.rsplit('|', 1); condense[(sg, fk)] = int(cl)
    elif l.startswith('!'): block.add(l[1:])
    elif l.startswith('~'): lv[l[3:]] = int(l[1])
    else: lv[l] = 0
def save():
    with open(KF, 'w', encoding='utf-8') as f:
        for s, v in sorted(lv.items()): f.write((('~%d:' % v) if v else '') + s + '\n')
        for s in sorted(block): f.write('!' + s + '\n')
        for (sg, k), dd in sorted(wfix.items()): f.write('W:%d:%d:%s\n' % (k, dd, sg))
        for (sg, k) in sorted(wblock): f.write('WB:%d:%s\n' % (k, sg))
        for sg in sorted(splitok): f.write('S:' + sg + '\n')
        for h in sorted(pbreak): f.write('P:' + h + '\n')
        for (sg, fk), cl in sorted(condense.items()): f.write('C:%d:%s|%s\n' % (cl, sg, fk))

bg = lambda x: {x[k:k + 2] for k in range(len(x) - 1)}
def match(text, tbls):
    """返回最匹配的表签名集合（同文表分数相近的一并返回）；匹配不上返回空集。"""
    want = bg(text)
    if not want: return set()
    score = sorted(((len(want & t['bg']) / len(want), t['sig']) for t in tbls), reverse=True)
    if not score or score[0][0] < 0.75: return set()
    return {s for v, s in score if v >= max(0.75, score[0][0] - 0.1)}

dump = os.path.join(tempfile.mkdtemp(), 'dump.jsonl')
tried = {}
for n in range(1, 8):
    save()
    build(dump)
    tbls = [json.loads(l) for l in open(dump, encoding='utf-8') if l.strip()]
    for t in tbls: t['bg'] = bg(t['text'])
    sigs = {t['sig'] for t in tbls}
    changed = [s for s in list(lv) if s not in sigs]            # 表已改动：撤出
    for s in changed: del lv[s]
    block &= sigs
    for key in [k2 for k2 in wfix if k2[0] not in sigs]: del wfix[key]
    wblock = {k2 for k2 in wblock if k2[0] in sigs}
    pdf = to_pdf()
    res = json.loads(subprocess.run([sys.executable, os.path.join(T, 'check_splits.py'), pdf, '--json'],
                                    capture_output=True, text=True).stdout)
    # SD-85 实测加宽：末行孤字、短格折行 → 这张表这一列加宽，下一遍重排（原文 <br> 主动换行的不算）
    sys.path.insert(0, T); from layout_measure import measure
    wadd, wgive, now, cadd = 0, 0, set(), 0
    for x in measure(pdf):
        for sg in match(x['table'], tbls):           # 速查区与正文同文的表一并加宽
            tb = next(t for t in tbls if t['sig'] == sg)
            # 列号按格子中心的横向位置换算回源表列（PDF 识别合并单元格时会多切列，列号不可靠）
            Wt = tb.get('W') or []
            if not Wt: continue
            pos, acc, col = x['cx'] * sum(Wt) / max(x['tw'], 1), 0, len(Wt) - 1
            for q, wq in enumerate(Wt):
                acc += wq
                if pos <= acc: col = q; break
            key = (sg, col)
            if key in now or (x.get('prev') and x['prev'] + '|' in tb.get('br', '')): continue
            if key in wblock:   # 加宽无效（表已占满版面、邻列太窄）：只在末行孤字时收紧这一格的字距，最多两级
                fk = (x.get('first') or '')[:10]
                if fk and x.get('orphan') and condense.get((sg, fk), 0) < 2: condense[(sg, fk)] = condense.get((sg, fk), 0) + 1; cadd += 1
                now.add(key); continue   # 2026-10-03：与 check_layout T8 同口径——只有倒数第二行止于原文 <br> 才算主动换行（原来看第一行，带 <br> 的格一律跳过）
            now.add(key)
            new = wfix.get(key, 0) + int(x['extra_pt'] * 20) + 20
            if new > 1500 or tried.get(key, 0) >= 3: wblock.add(key); wfix.pop(key, None); wgive += 1; continue   # 加宽 3 次或超过 75pt 仍不行：放弃
            wfix[key] = new; tried[key] = tried.get(key, 0) + 1; wadd += 1
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
    shutil.rmtree(os.path.dirname(pdf), ignore_errors=True)
    add, up, gave, miss = 0, 0, 0, []
    for sp in res['splits']:
        cand = match(sp['text'], tbls)
        if not cand:
            if sp['fits']: miss.append('第 %d 页（%s…）' % (sp['page'], sp['text'][:16]))
            continue
        for s in cand - block - splitok:   # 允许按块分页的块索引表不再强制 / 压缩
            if s not in lv:
                if sp['fits']: lv[s] = 0; add += 1
                elif sp['over'] <= 1.2 and not sp['multi']:
                    lv[s] = 1 if sp['over'] <= 1.04 else 2 if sp['over'] <= 1.10 else 3; up += 1
            elif lv[s] < MAXLV:
                lv[s] = max(lv[s] + 1, 1); up += 1                 # 强制 / 压缩后仍断开：升一级
            else:
                del lv[s]; block.add(s); gave += 1                  # 8pt 仍放不下：放弃，照常分页
    print('fit_fix 第 %d 遍：%d 页，断表 %d 处（本可整页 %d）；新增整表 %d、压缩升级 %d、放弃 %d、撤出 %d；当前压缩 %d 张；列加宽 +%d、放弃 %d；标题孤页处理 %d%s' % (
        n, res['pages'], len(res['splits']), sum(s['fits'] for s in res['splits']), add, up, gave, len(changed),
        sum(1 for v in lv.values() if v), wadd, wgive, ladd, ('；匹配不到：' + '、'.join(miss)) if miss else ''))
    if cadd: print('fit_fix 第 %d 遍：末行孤字收紧字距 %d 格' % (n, cadd))
    if not (add or up or gave or changed or wadd or wgive or ladd or cadd): break
    if n == 7: print('fit_fix：7 遍仍未稳定，docx 保持本遍结果，交 check_layout 报告')
save()
