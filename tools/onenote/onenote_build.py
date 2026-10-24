#!/usr/bin/env python3
"""把基线 Word 写成 OneNote 阅读版（用户规则见同目录《OneNote写入流程.md》）。
用法：
  python3 onenote_auth.py                                   # 首次或令牌失效时登录
  python3 onenote_build.py --docx 基线.docx --notebook 日积月累 --group "B737 机型理论知识笔记（新版）"
  可选：--chapters 1-6（默认，第零章不写）  --only 1.6,3.2（只重写这些节，要求同分区其后的节一并重写以保页序）
        --figs-only（只补插图）  --audit-only（只核查）
步骤：docx → LibreOffice HTML → 按 h1 章 / h2 节拆页 → 转换（宋体-简、对齐、宽度、去块索引 / 来源）
      → 建分区组和分区 → 逐页建页并读回确认 → 按 notes_src 的 %%FIG%% 补插图 → 只读核查。"""
import argparse, glob, html, json, os, re, struct, subprocess, sys, time, uuid
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from onenote_api import req, content, list_pages, create_page, patch_with_image, patch_with_images, delete_page
from onenote_conv import convert, page, strip_index, FONT, TABW
NS = os.path.expanduser('~/flight-repos/gh-private/notes_src/')
WORK = os.path.expanduser('~/flight-repos/_work/onenote')
CHN = {'第一章': '第一章 系统描述', '第二章': '第二章 操作程序', '第三章': '第三章 运行手册', '第四章': '第四章 飞行技术提示', '第五章': '第五章 模拟机笔记', '附录': '附录 机型差异'}   # 2026-10-24 SD-234
# 2026-10-15 SD-185 第一章改名「系统描述」；接口不能改分区名——OneNote 里还叫旧名时沿用旧分区写入，页眉照新名（用户在 OneNote 里右键改名即可）
ALIAS = {'第一章 系统描述': '第一章 系统理论', '第二章 操作程序': '第二章 机组训练手册', '第四章 飞行技术提示': '第五章 技术提示', '第五章 模拟机笔记': '第四章 模拟机训练'}   # SD-234（2026-10-24）：沿用旧分区写入，用户在 OneNote 里右键改名
def sec_name(name, secs): return name if name in secs or ALIAS.get(name) not in secs else ALIAS[name]
NB_ID = '0-A7EF8B8B1AAC2778!s8966e7df85744afd957d97c2f66eb275'   # 用户自己的「飞行」笔记本（2026-10-05 用户：「以后更新就在这里更新」）
def locate(nb_id, group=''):
    """返回 ({分区名: id}, 新建分区用的父路径)。group 为空时分区直接在笔记本下。"""
    if group:
        sgs = req('GET', '/notebooks/%s/sectionGroups?$select=id,displayName' % nb_id)['value']
        sg = next((g for g in sgs if g['displayName'] == group), None) or req('POST', '/notebooks/%s/sectionGroups' % nb_id, {'displayName': group})
        parent = '/sectionGroups/%s' % sg['id']
    else:
        parent = '/notebooks/%s' % nb_id
    return {s['displayName']: s['id'] for s in req('GET', parent + '/sections?$select=id,displayName')['value']}, parent
NUM = {'1': '第一章', '2': '第二章', '3': '第三章', '4': '第四章', '5': '第五章', '6': '附录'}
def L(*a): print(*a, flush=True)
def txt(x): return re.sub(r'[\s​]+', ' ', html.unescape(re.sub(r'<[^>]+>', '', x))).strip()
def xref_plain(body):
    """2026-10-11 用户「做法一」：OneNote 不用链接——节 / 主题开头的「相关：↗ …」改成文字「相关：见 …」"""
    def f(m):
        p = m.group(0)
        if not (txt(p).startswith('相关：') and '↗' in p): return p
        return re.sub(r'</?a\b[^>]*>', '', p).replace('↗ ', '见 ').replace('↗', '见 ')
    return re.sub(r'<p\b[^>]*>.*?</p>', f, body, flags=re.S)
# ---------- 1. docx → html → 拆页 ----------
def split(docx):
    os.makedirs(WORK, exist_ok=True)
    subprocess.run(['soffice', '--headless', '--convert-to', 'html', '--outdir', WORK, docx], check=True, capture_output=True)
    s = open(os.path.join(WORK, os.path.splitext(os.path.basename(docx))[0] + '.html'), encoding='utf-8').read()
    head = s[:s.find('<body')]
    marks = [(m.start(), m.group(1), txt(m.group(2))) for m in re.finditer(r'<(h[12])\b[^>]*>(.*?)</\1>', s, re.S)]
    pages, chap = [], None
    for k, (pos, tag, t) in enumerate(marks):
        if tag == 'h1': chap = '附录' if t.startswith('附录') else t[:3]; continue   # SD-234：附录章（「附录　机型差异」）
        end = marks[k + 1][0] if k + 1 < len(marks) else s.find('</body>')
        body = s[s.find('>', pos) + 1:end]
        body = re.split(r'<p\b[^>]*>\s*<a name="TOPICS', body)[0]                   # SD-149「按主题查」排在第五章后，属索引，OneNote 不要
        body = re.sub(r'^.*?</h2>', '', body, count=1, flags=re.S)                  # 去掉与页标题重复的节标题
        body = re.sub(r'<p[^>]*page-break-before: always[^>]*>\s*(<br/>\s*)*</p>\s*<p[^>]*>\s*<font color="#c9d8e8">.*?</p>\s*$', '', body, flags=re.S)   # 章封面大号章号残留
        body = xref_plain(body)
        pages.append({'chap': chap, 'title': t, 'html': head + '<body>' + body + '</body></html>'})
    return pages
# ---------- 2. 插图 ----------
def plain(s): return re.sub(r'[\s　]+', '', html.unescape(re.sub(r'<[^>]+>', '', s)))
def figs_for(title):
    f = glob.glob(NS + '*/' + title + '.md')
    if not f: return []
    Ls = open(f[0], encoding='utf-8').read().split('\n'); out = []
    for i, l in enumerate(Ls):
        if not l.startswith('%%FIG'): continue
        if [x.strip() for x in l.split('|')][-1] == '速查': continue   # 2026-10-19：图行末分工标注（P03）「速查」只放速查，完整版 OneNote 不插
        fn, cap, wmm = [x.strip() for x in re.sub(r'^%%FIG(SIDE|RIGHT)?%%\s*', '', l).split('|')][:3]   # SD-207：第 4 段「同页」只给成品排版用
        # OneNote 版删了「详见」行，不能当锚点（1.14-2）
        prev = ([x for x in Ls[:i] if x.strip() and not re.match(r'\s*详见', x)] or [''])[-1].strip(); nxt = ([x for x in Ls[i + 1:] if x.strip()] or [''])[0].strip()   # 2026-10-19：图在文件末行（1.6）时没有下一行
        anchor = (re.sub(r'^#+\s*', '', nxt), 'before') if prev.startswith('<') else (re.sub(r'^#+\s*', '', prev), 'after')
        if prev.startswith('</table') and nxt.startswith('<'):   # 2026-10-16：图夹在两张表之间（1.13-1）——用上一张表的表头文字找到该表，插在表后
            k = max(j for j in range(i) if Ls[j].lstrip().startswith('<table'))
            hdr = next((x for x in Ls[k:i] if ('<th' in x or '<td' in x) and 'premise' not in x), '')   # 2026-10-19：前提通栏行在 OneNote 里移到表外（unmerge），用真正的表头找表
            anchor = (plain(re.sub(r'<[^>]+>', ' ', hdr))[:12], 'after-table')
        if prev.startswith('%%FIG'): anchor = ([x.strip() for x in re.sub(r'^%%FIG(SIDE|RIGHT)?%%\s*', '', prev).split('|')][1], 'after')   # 连排几张图：接在上一张图注后（2026-10-11 重要天气预告图三张）
        alt = (re.sub(r'^#+\s*', '', nxt), 'before') if not nxt.startswith(('<', '%%')) else None   # 2026-10-18：前一行（多为注）在 OneNote 里对不上时，用图后一行作后备锚点（1.15-1）
        out.append(dict(file=NS + fn, cap=cap, wmm=float(wmm or 120), anchor=anchor, alt=alt))
    return out
def find_anchor(h, text):
    text = re.sub(r'\[\[(\d+\.\d+)[^\]]*\]\]', r'\1', text); text = re.sub(r'(\d+\.\d+)(\s*·\s*\1)+', r'\1', text)
    text = re.sub(r'^%%SUB%%\s*', '', re.sub(r'\s*\{★+\}\s*$', '', text.strip()))   # 2026-10-19 R5：标题行末的分级标记 {★} / 节内小标题 %%SUB%% 是源码写法，页面上没有（R2 加星后第一次写 OneNote，图全部找不到锚点）
    want = plain(text)[:18]
    for m in re.finditer(r'<p\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</p>', h, re.S):
        if want and want in plain(m.group(2)): return m.group(1)
    # 2026-10-19 R5：表后的「注：」在 OneNote 里是淡蓝底单格表（注块），段落里找不到 → 找含这段文字的表，图插在整块前 / 后（锚点注总在块的边上：图把连续注隔开）
    for m in re.finditer(r'<table\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</table>', h, re.S):
        if want and want in plain(m.group(2)): return m.group(1)
def add_figs(pid, title):
    for fg in figs_for(title):
        h = content(pid, ids=True) or ''
        if 'alt="%s"' % html.escape(fg['cap']) in h: L('   图已有', fg['cap'][:16]); continue
        if fg['anchor'][1] == 'after-table':
            tid = next((m.group(1) for m in re.finditer(r'<table\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</table>', h, re.S) if fg['anchor'][0] and fg['anchor'][0] in plain(m.group(2))), None)
            pos = 'after'
        else:
            tid = find_anchor(h, fg['anchor'][0]); pos = fg['anchor'][1]
        if not tid and fg.get('alt'): tid = find_anchor(h, fg['alt'][0]); pos = fg['alt'][1]
        if not tid: L('   !! 图找不到插入位置', fg['cap'][:16]); continue
        data = open(fg['file'], 'rb').read(); pw, ph = struct.unpack('>II', data[16:24])
        w = min(TABW, round(fg['wmm'] / 270 * TABW)); name = 'fig' + uuid.uuid4().hex[:8]
        c = ('<img src="name:%s" width="%d" height="%d" alt="%s"/>' % (name, w, round(w * ph / pw), html.escape(fg['cap'])) +
             '<p style="margin-top:0;margin-bottom:0"><span style="font-family:%s;font-size:9pt;color:#595959">%s</span></p>' % (FONT, html.escape(fg['cap'], quote=False)))
        try:
            patch_with_image(pid, [{'target': tid, 'action': 'insert', 'position': pos, 'content': c}], name, data); L('   补图', fg['cap'][:16], w)
        except RuntimeError as e:
            L('   !! 补图出错（可能已插入，重跑会自动跳过）', fg['cap'][:16], str(e)[:40]); time.sleep(20)
        time.sleep(2)
def add_cell_imgs(pid, title, src=None):
    """1019R6（记录第 16 条 ND 符号表）：源文件表格格内 <img src="figures/…" w="mm">——成品里是格内小图；
    OneNote 转换时图丢失、格子留空。按表头文字找到该表，逐行在第一格的空段落前插入对应小图（一次 PATCH 带全部图）。"""
    quiet = src is not None   # 速查：整本源文件对每页都试一遍，本页没有该表不报
    if src is None:
        f = glob.glob(NS + '*/' + title + '.md')
        if not f: return
        src = open(f[0], encoding='utf-8').read()
    for tm in re.finditer(r'(?s)<table[^>]*>(.*?)</table>', src):
        rows = re.findall(r'(?s)<tr(?![^>]*class="(?:hdr|premise|note)")[^>]*>(.*?)</tr>', tm.group(1))
        imgs = [re.search(r'<img\s+src="([^"]+)"(?:\s+w="([\d.]+)")?', r) for r in rows]
        if not any(imgs): continue
        hdr = plain(re.sub(r'<[^>]+>', ' ', re.search(r'(?s)<tr class="hdr">(.*?)</tr>', tm.group(1)).group(1)))[:12]
        h = content(pid, ids=True) or ''
        tb = next((m for m in re.finditer(r'<table\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</table>', h, re.S) if hdr and hdr in plain(m.group(2))), None)
        if not tb:
            if quiet is False: L('   !! 格内图找不到表', title, hdr)
            continue
        if '<img' in tb.group(2): L('   格内图已有', hdr); continue
        trs = re.findall(r'(?s)<tr\b[^>]*>(.*?)</tr>', tb.group(2))[1:]   # 去表头
        cmds, files = [], []
        for tr, im in zip(trs, imgs):
            if not im: continue
            pm = re.search(r'<td\b[^>]*>\s*<p\b[^>]*\bid="([^"]+)"', tr)
            if not pm: continue
            data = open(NS + im.group(1), 'rb').read(); pw, ph = struct.unpack('>II', data[16:24])
            w = round(float(im.group(2) or 15) / 270 * TABW); name = 'cimg' + uuid.uuid4().hex[:8]
            cmds.append({'target': pm.group(1), 'action': 'insert', 'position': 'before', 'content': '<img src="name:%s" width="%d" height="%d"/>' % (name, w, round(w * ph / pw))})
            files.append((name, data))
        if cmds:
            try: patch_with_images(pid, cmds, files); L('   格内图', hdr, len(cmds))
            except RuntimeError as e: L('   !! 格内图出错', hdr, str(e)[:60])
            time.sleep(3)
# ---------- 3. 主流程 ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--notebook-id', default=NB_ID, help='默认：用户自己的「飞行」笔记本（另有同名共享只读本，必须按编号区分）')
    ap.add_argument('--group', default='', help='分区组名；留空表示分区直接放在笔记本下（2026-10-05 起）')
    ap.add_argument('--chapters', default='1-6')
    ap.add_argument('--prune', action='store_true', help='整章重写后删掉本次写入分区里、标题不在新基线中的旧页（节号 / 节名变了时用）')
    ap.add_argument('--only', default=''); ap.add_argument('--figs-only', action='store_true'); ap.add_argument('--audit-only', action='store_true')
    a = ap.parse_args()
    lo, hi = (a.chapters.split('-') + [a.chapters])[:2]; want_ch = [NUM[str(n)] for n in range(int(lo), int(hi) + 1)]
    pages = [p for p in split(a.docx) if p['chap'] in want_ch]; L('拆页', len(pages), '节')
    secs, parent = locate(a.notebook_id, a.group)
    # 2026-10-10 用户：「onenote 模拟机训练章节我正在人工做修正和优化，没有我的允许从现在开始先不要主动去更改」——
    # 锁定分区（~/.config/flight-onenote/locked_sections 每行一个分区名，或环境变量 ONENOTE_LOCKED 逗号分隔）：不写页、不补标题、不删重复、不核查
    lf = os.path.expanduser('~/.config/flight-onenote/locked_sections')
    locked = {x.strip() for x in (open(lf, encoding='utf-8').read().splitlines() if os.path.exists(lf) else []) + os.environ.get('ONENOTE_LOCKED', '').split(',') if x.strip()}
    if locked: L('锁定分区（跳过）：', '、'.join(sorted(locked)))
    secs = {k: v for k, v in secs.items() if k not in locked}
    pages = [p for p in pages if CHN[p['chap']] not in locked and sec_name(CHN[p['chap']], secs) not in locked]
    only = [x for x in a.only.split(',') if x]
    for p in pages:
        name = sec_name(CHN[p['chap']], secs)
        if name not in secs: secs[name] = req('POST', parent + '/sections', {'displayName': name})['id']; L('建分区', name)
        sid = secs[name]; existing = {}
        for q in list_pages(sid):   # 2026-10-16：同名页有打不开的残页时，补图选能读的那张
            if q['title'] in existing and a.figs_only and content(existing[q['title']]) is not None: continue
            existing[q['title']] = q['id']
        if a.audit_only: continue
        if only and not any(p['title'].startswith(o + ' ') for o in only):
            if a.figs_only and p['title'] in existing: add_figs(existing[p['title']], p['title'])
            continue
        if a.figs_only:
            if p['title'] in existing: add_figs(existing[p['title']], p['title']); add_cell_imgs(existing[p['title']], p['title'])
            continue
        for q in list_pages(sid):
            if q['title'] == p['title']: delete_page(q['id'])
        src, _ = strip_index(p['html'])
        pid = create_page(sid, p['title'], page(p['title'], convert(src), '%s ｜ %s' % (CHN[p['chap']], p['title'])))
        # 2026-10-19（1018R2 写 4.6 返回成功但页面没留下）：每节写完立即读回核查，读不到就当场重写一次
        if pid:
            time.sleep(8)
            if content(pid) is None:
                L('   !! 写后读不到，重写', p['title'])
                try: delete_page(pid)
                except Exception: pass
                time.sleep(6)
                pid = create_page(sid, p['title'], page(p['title'], convert(src), '%s ｜ %s' % (CHN[p['chap']], p['title'])))
        L('写', p['title'], 'OK' if pid else '!! 失败')
        if pid: add_figs(pid, p['title']); add_cell_imgs(pid, p['title'])
        time.sleep(6)                     # 2026-10-06：1.5 秒会触发 429 限流
    # 2026-10-09：新页有时标题为空（1009R4 写全书 27 页）——按页眉「章名 ｜ 节名」补标题；同名重复只留最新一页
    for name, sid in secs.items():
        for q in list_pages(sid):
            if q['title']: continue
            m = re.search(r'(?:第[一二三四五]章|附录) [^｜<]+ ｜ ([^<]+)', content(q['id']) or '')   # SD-234：附录页眉
            if not m: continue
            for k in range(6):
                try: req('PATCH', '/pages/%s/content' % q['id'], [{'target': 'title', 'action': 'replace', 'content': m.group(1).strip()}]); L('补标题', m.group(1).strip()); break
                except RuntimeError as e:
                    if not str(e).startswith('429'): raise
                    L('   限流，等 5 分钟'); time.sleep(300)
            time.sleep(5)
    time.sleep(20)
    for name, sid in secs.items():
        allp = req('GET', '/sections/%s/pages?$select=id,title,createdDateTime&$top=100' % sid)['value']
        for tt in {q['title'] for q in allp if q['title']}:
            same = sorted([q for q in allp if q['title'] == tt], key=lambda q: q['createdDateTime'], reverse=True)
            for q in same[1:]: delete_page(q['id']); L('删重复旧页', tt)
    if a.prune and not only and not a.audit_only and not a.figs_only:
        for name, sid in secs.items():
            want = {p['title'] for p in pages if sec_name(CHN[p['chap']], secs) == name}
            if not want: continue
            for q in list_pages(sid):
                if q['title'] not in want: delete_page(q['id']); L('删旧页（新基线无此节）', name, q['title'] or '（空标题）')
    # 只读核查
    bad = 0
    for name, sid in secs.items():
        ps = list_pages(sid); readable = [q for q in ps if content(q['id'])]
        titles = [q['title'] for q in readable]
        want = [p['title'] for p in pages if sec_name(CHN[p['chap']], secs) == name]
        miss = [t for t in want if t not in titles]; dup = sorted({t for t in titles if titles.count(t) > 1})
        bad += len(miss) + len(dup)
        L('核查', name, '可读', len(readable), '/', len(ps), '缺', miss, '重复', dup)
    L('完成' if not bad else '!! 有缺页或重复，见上')
if __name__ == '__main__': main()
