#!/usr/bin/env python3
"""release.py —— 一键出基线（规则重组方案 A4，2026-10-21；步骤与总纲第十节一一对应）。

用法：python3 release.py            # 只做前置检查并列出每一步要做什么（不改任何东西）
      python3 release.py --go       # 照单执行；任一步失败立即停下，后面的步骤不做
      python3 release.py --go --study   # 另做第 ⑧ 步学习系统（SD-220：只在用户指令更新学习系统时加）
前提：已按总纲第八节全量重建并收口（DOC_EDITION=<版本名> ./sync.sh --full），build/ 里是本次成品。

自动做：①版本名 ②前置检查（排版报告错误 0、规则自检、两库无未提交改动、成品版本号对得上、两本 verify、整本缩略图过目登记 SD-221）③推送两库（push_all.sh）
④推送成功后打标签并推标签 ⑤分发成品到桌面 / 用户改稿基线 / 交接包（旧版移 _to_delete）⑥对照表 ⑦源码包
⑧学习系统 check + snapshot（SD-220：默认跳过，用户指令时加 --study）⑨更新《基线信息》表格各行 ⑩重新生成《现状一页纸》⑪云盘复查（未确认的过几分钟再跑 drive_verify.py）。
不自动做（脚本最后提醒）：变动说明正文、《基线信息》「本基线新增」叙述、《00_先读我》本轮记录、OneNote（只在用户指令时）。"""
import datetime, glob, hashlib, json, os, re, shutil, subprocess, sys

H = os.path.expanduser('~')
G = os.path.join(H, 'flight-repos', 'gh-private')
PUB = os.path.join(H, 'flight-repos', 'pub')
T = os.path.join(PUB, 'tools')
WS = os.path.join(H, 'Desktop', '飞行理论笔记整理')
HB = os.path.join(WS, 'Muse交接包')
BASE_DIR = os.path.join(WS, '用户改稿', '基线')
TRASH = os.path.join(H, 'flight-repos', '_to_delete')
STUDY = os.path.join(H, 'flight-repos', 'study')
SNAP = os.path.join(H, 'Library', 'CloudStorage', 'GoogleDrive-yzyboeing@gmail.com', 'My Drive', '03_AI协作交接', 'B737笔记·Claude与噜噜交接', '学习系统', '笔记源快照')
STEMS = ('B737机型理论速查笔记',)   # SD-234（2026-10-24）：只维护这一本；《详解》《速查》两本只作溯源
GO = '--go' in sys.argv
TODAY = datetime.date.today().strftime('%Y%m%d')


def run(cmd, cwd=None, check=True):
    print('  $ ' + (cmd if isinstance(cmd, str) else ' '.join(cmd)))
    if not GO: return ''
    r = subprocess.run(cmd, cwd=cwd, shell=isinstance(cmd, str), capture_output=True, text=True)
    out = (r.stdout + r.stderr).strip()
    if out: print('    ' + out.replace('\n', '\n    ')[-1500:])
    if check and r.returncode: sys.exit('✗ 失败，已停下（后面的步骤没做）')
    return r.stdout


def move_old(folder, ed, exts):
    """folder 里 STEMS + 别的版本名 + exts 的成品移到 _to_delete/成品旧版_<今天>/<旧版本>/"""
    for p in sorted(glob.glob(os.path.join(folder, 'B737机型理论*'))):   # 2026-10-23：SD-225 改名后新旧书名都认（原来只认「知识笔记 / 知识速查」，新书名的旧版移不走）
        m = re.match(r'(B737机型理论(?:知识笔记|知识速查|速查笔记|详解|速查))(\d{4}R\d+)\.(\w+)$', os.path.basename(p))
        if not m or m.group(2) == ed or m.group(3) not in exts: continue
        dst = os.path.join(TRASH, '成品旧版_%s' % TODAY, m.group(2))
        print('  移旧版 %s → %s' % (os.path.relpath(p, H), os.path.relpath(dst, H)))
        if GO: os.makedirs(dst, exist_ok=True); shutil.move(p, os.path.join(dst, os.path.basename(p)))


def sha16(p):
    return hashlib.sha256(open(p, 'rb').read()).hexdigest()[:16]


def pages(p):
    import fitz
    return len(fitz.open(p))


def step(n, title):
    print('\n[%s] %s' % (n, title))


# ① 版本名
step(1, '版本名')
nx = subprocess.run([sys.executable, os.path.join(T, 'release_name.py'), '--next'], capture_output=True, text=True).stdout.split()
ED, TAG = nx[0], nx[1]
print('  本次：%s（标签 %s）' % (ED, TAG))

# ② 前置检查
step(2, '前置检查')
probs = []
rep = open(os.path.join(G, 'build', '排版检查报告.md'), encoding='utf-8').read()
m = re.search(r'错误 (\d+) 条', rep)
if not m or m.group(1) != '0': probs.append('排版检查报告错误不为 0（先收口）')
if '新 0 条' not in rep.split('\n')[2]: print('  ! 还有未裁定的新建议——确认都看过（出基线不拦）')
if subprocess.run([sys.executable, os.path.join(T, 'check_rules.py'), G], capture_output=True).returncode: probs.append('规则自检 check_rules.py 未通过')
for nm, p in (('私有库', G), ('工具库', PUB)):
    if subprocess.run(['git', 'status', '--porcelain'], cwd=p, capture_output=True, text=True).stdout.strip(): probs.append('%s有未提交改动' % nm)
ed_cfg = subprocess.run(['git', 'config', '-f', os.path.join(G, '成品设置.gitconfig'), '--get', 'notes.docEdition'], capture_output=True, text=True).stdout.strip()
if ed_cfg != ED: probs.append('成品版本号是 %s，不是 %s——用 DOC_EDITION=%s 重建' % (ed_cfg, ED, ED))
for s in STEMS:
    for e in ('docx', 'pdf'):
        if not os.path.exists(os.path.join(G, 'build', '%s.%s' % (s, e))): probs.append('缺 build/%s.%s' % (s, e))
if subprocess.run(['git', 'rev-parse', '-q', '--verify', 'refs/tags/' + TAG], cwd=G, capture_output=True).returncode == 0: probs.append('标签 %s 已存在' % TAG)
# SD-221（2026-10-21 用户：「应该是你自己审核的时候就发现并优化」）：两本都过 verify.py（空白页等），且整本缩略图已逐张过目并登记
for s in STEMS:
    pdf = os.path.join(G, 'build', s + '.pdf')
    if os.path.exists(pdf) and subprocess.run([sys.executable, os.path.join(T, 'verify.py'), pdf], capture_output=True).returncode:
        probs.append('%s 没过 verify.py（空白页 / 标签泄漏等）' % s)
try:
    _rec = json.load(open(os.path.join(G, 'build', '整本过目.json'), encoding='utf-8'))
    _cur = {s: hashlib.sha256(open(os.path.join(G, 'build', s + '.pdf'), 'rb').read()).hexdigest()[:16] for s in STEMS}
    if _rec.get('版本') != ED or _rec.get('成品') != _cur: probs.append('整本过目登记与当前成品不符——跑 contact_sheet.py 逐张看完再 --done')
except Exception:
    probs.append('没有整本过目登记——python3 %s/contact_sheet.py --repo %s，逐张看完、存疑页处理完后 --done' % (T, G))
for p in probs: print('  ✗ ' + p)
if probs: sys.exit('前置检查没过，未做任何改动')
print('  ✓ 全部通过')

# ③ 推送两库
step(3, '推送两库（含泄漏扫描）')
run([os.path.join(T, 'push_all.sh'), G])

# ④ 打标签（推送成功后）
step(4, '打标签')
run(['git', 'tag', '-a', TAG, '-m', '基线 %s' % ED], cwd=G)
run(['git', 'push', 'origin', TAG], cwd=G)

# ⑤ 分发成品
step(5, '分发成品')
for folder, exts in ((WS, ('docx', 'pdf')), (HB, ('docx', 'pdf')), (BASE_DIR, ('docx',))):
    move_old(folder, ED, exts)
    for s in STEMS:
        for e in exts:
            src, dst = os.path.join(G, 'build', '%s.%s' % (s, e)), os.path.join(folder, '%s%s.%s' % (s, ED, e))
            print('  复制 %s → %s' % (os.path.basename(src), os.path.relpath(dst, H)))
            if GO: shutil.copy2(src, dst)

# ⑥ 对照表
step(6, '两本标题与主题对照表')
for p in glob.glob(os.path.join(HB, '两本标题与主题对照表_*.md')):
    if not p.endswith('_%s.md' % ED):
        print('  移旧版 %s' % os.path.basename(p))
        if GO: os.makedirs(os.path.join(TRASH, '成品旧版_%s' % TODAY), exist_ok=True); shutil.move(p, os.path.join(TRASH, '成品旧版_%s' % TODAY, os.path.basename(p)))
run([sys.executable, os.path.join(T, 'title_map.py'), os.path.join(HB, '两本标题与主题对照表_%s.md' % ED)])

# ⑦ 源码包
step(7, '源码包')
run([os.path.join(T, 'lulu_bundle.sh'), ED])

# ⑧ 学习系统
step(8, '学习系统 check + snapshot')
if '--study' not in sys.argv:   # SD-220（2026-10-21 用户）：「学习系统以后都不要同步更新了除非我指令更新」
    print('  · 默认跳过（SD-220）；用户指令更新学习系统时加 --study')
elif os.path.isdir(STUDY):
    run([sys.executable, 'tools/study.py', 'check'], cwd=STUDY)
    run([sys.executable, 'tools/study.py', 'snapshot', SNAP], cwd=STUDY)
else:
    print('  ! 没有 %s，跳过' % STUDY)

# ⑨ 基线信息表格
step(9, '《基线信息》表格')
bi = os.path.join(HB, '基线信息.md')
info = {}
for s in STEMS:
    for e in ('docx', 'pdf'):
        info[(s, e)] = os.path.join(G, 'build', '%s.%s' % (s, e))
gh = subprocess.run(['git', 'rev-parse', '--short', 'HEAD'], cwd=G, capture_output=True, text=True).stdout.strip()
tk = subprocess.run(['git', 'rev-parse', '--short', 'HEAD'], cwd=PUB, capture_output=True, text=True).stdout.strip()
sd = max(int(x) for x in re.findall(r'SD-(\d+)', open(os.path.join(PUB, 'prompt', 'standing-decisions.md'), encoding='utf-8').read()))
rows = {
    '基线编号': '**%s**（git 标签 %s；封面版本号「%s　·　YZYBOEING」）' % (ED, TAG, ED),
    '笔记源提交': gh, '工具链提交': tk,
    '基线 Word': '`%s%s.docx`（sha256 前 16 位 %s）' % (STEMS[0], ED, sha16(info[(STEMS[0], 'docx')])),
    '生成日期': datetime.date.today().isoformat(),
}
secs = [f for f in glob.glob(os.path.join(G, 'notes_src', '[1-9] *', '*.md')) if re.match(r'\d+\.\d+ ', os.path.basename(f))]
nent = sum(len(re.findall(r'(?m)^### \d+\. ', open(f, encoding='utf-8').read())) for f in secs)
qsrc = open(os.path.join(G, '速查', '速查源.md'), encoding='utf-8').read()
qdesc = '%d 条，%d 段 %d 个主题' % (len(re.findall(r'(?m)^### ', qsrc)), len(re.findall(r'(?m)^%%PART%%', qsrc)), len(re.findall(r'(?m)^## ', qsrc)))
fdesc = '%d 章 %d 节 %d 条' % (len({os.path.dirname(f) for f in secs}), len(secs), nent)
try:
    rows['基线 PDF'] = '`%s%s.pdf`（%d 页；%s；sha256 前 16 位 %s）。**页码一律以这份为准**' % (STEMS[0], ED, pages(info[(STEMS[0], 'pdf')]), fdesc, sha16(info[(STEMS[0], 'pdf')]))
    rows['速查版'] = '`%s%s.pdf`（%d 页，%s；sha256 前 16 位 %s；Word 同名，%s）。独立成册' % (STEMS[1], ED, pages(info[(STEMS[1], 'pdf')]), qdesc, sha16(info[(STEMS[1], 'pdf')]), sha16(info[(STEMS[1], 'docx')]))
except Exception as ex:
    print('  ! 页数 / 校验值没取到：%s' % ex)
txt = open(bi, encoding='utf-8').read()
old_ed = re.search(r'基线编号[^|]*\|\s*\*\*(\d{4}R\d+)\*\*', txt)
for k, v in rows.items():
    pat = re.compile(r'(?m)^(\| %s[^|]*\| ).*?( \|)$' % re.escape(k))
    if pat.search(txt):
        txt = pat.sub(lambda mm: mm.group(1) + v + mm.group(2), txt, count=1); print('  %s → %s' % (k, v[:60]))
    else:
        print('  ! 《基线信息》里没有「%s」这一行，没改' % k)
txt = re.sub(r'(?m)^(\| 取代的旧基线 \| ).*?( \|)$', lambda mm: mm.group(1) + (old_ed.group(1) if old_ed else '—') + mm.group(2), txt, count=1)
txt = re.sub(r'更新到 \*\*SD-\d+\*\*', '更新到 **SD-%d**' % sd, txt)
if GO: open(bi, 'w', encoding='utf-8').write(txt)

# ⑩ 现状一页纸
step(10, '现状一页纸')
run([sys.executable, os.path.join(T, 'status_page.py')])

# ⑪ 云盘复查
step(11, '云盘复查')
targets = [os.path.join(HB, '%s%s.%s' % (s, ED, e)) for s in STEMS for e in ('docx', 'pdf')] + [bi, os.path.join(HB, '两本标题与主题对照表_%s.md' % ED), os.path.join(HB, '源码包')]
run([sys.executable, os.path.join(T, 'drive_verify.py')] + targets, check=False)

print('\n还要手工做：① 变动说明（噜噜收件箱_核对意见/Claude_%s变动说明_….md，上一版移走）；② 《基线信息》「与上一基线的差别」与「本基线新增」；'
      '③ 《00_先读我》本轮记录；④ 云盘未确认的过几分钟再跑 drive_verify.py；⑤ OneNote 只在用户指令时更新，写了就记《OneNote待同步记录》。' % ED)
if not GO: print('\n（只做了检查与列单；确认无误后加 --go 执行）')
