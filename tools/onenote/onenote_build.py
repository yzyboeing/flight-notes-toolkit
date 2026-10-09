#!/usr/bin/env python3
"""把基线 Word 写成 OneNote 阅读版（用户规则见同目录《OneNote写入流程.md》）。
用法：
  python3 onenote_auth.py                                   # 首次或令牌失效时登录
  python3 onenote_build.py --docx 基线.docx --notebook 日积月累 --group "B737 机型理论知识笔记（新版）"
  可选：--chapters 1-5（默认，第零章不写）  --only 1.6,3.2（只重写这些节，要求同分区其后的节一并重写以保页序）
        --figs-only（只补插图）  --audit-only（只核查）
步骤：docx → LibreOffice HTML → 按 h1 章 / h2 节拆页 → 转换（宋体-简、对齐、宽度、去块索引 / 来源）
      → 建分区组和分区 → 逐页建页并读回确认 → 按 notes_src 的 %%FIG%% 补插图 → 只读核查。"""
import argparse, glob, html, json, os, re, struct, subprocess, sys, time, uuid
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from onenote_api import req, content, list_pages, create_page, patch_with_image, delete_page
from onenote_conv import convert, page, strip_index, FONT, TABW
NS = os.path.expanduser('~/flight-repos/gh-private/notes_src/')
WORK = os.path.expanduser('~/flight-repos/_work/onenote')
CHN = {'第一章': '第一章 系统理论', '第二章': '第二章 机组训练手册', '第三章': '第三章 运行手册', '第四章': '第四章 模拟机训练', '第五章': '第五章 技术提示'}
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
NUM = {'1': '第一章', '2': '第二章', '3': '第三章', '4': '第四章', '5': '第五章'}
def L(*a): print(*a, flush=True)
def txt(x): return re.sub(r'[\s​]+', ' ', html.unescape(re.sub(r'<[^>]+>', '', x))).strip()
# ---------- 1. docx → html → 拆页 ----------
def split(docx):
    os.makedirs(WORK, exist_ok=True)
    subprocess.run(['soffice', '--headless', '--convert-to', 'html', '--outdir', WORK, docx], check=True, capture_output=True)
    s = open(os.path.join(WORK, os.path.splitext(os.path.basename(docx))[0] + '.html'), encoding='utf-8').read()
    head = s[:s.find('<body')]
    marks = [(m.start(), m.group(1), txt(m.group(2))) for m in re.finditer(r'<(h[12])\b[^>]*>(.*?)</\1>', s, re.S)]
    pages, chap = [], None
    for k, (pos, tag, t) in enumerate(marks):
        if tag == 'h1': chap = t[:3]; continue
        end = marks[k + 1][0] if k + 1 < len(marks) else s.find('</body>')
        body = s[s.find('>', pos) + 1:end]
        body = re.split(r'<p\b[^>]*>\s*<a name="TOPICS', body)[0]                   # SD-149「按主题查」排在第五章后，属索引，OneNote 不要
        body = re.sub(r'^.*?</h2>', '', body, count=1, flags=re.S)                  # 去掉与页标题重复的节标题
        body = re.sub(r'<p[^>]*page-break-before: always[^>]*>\s*(<br/>\s*)*</p>\s*<p[^>]*>\s*<font color="#c9d8e8">.*?</p>\s*$', '', body, flags=re.S)   # 章封面大号章号残留
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
        fn, cap, wmm = [x.strip() for x in re.sub(r'^%%FIG(SIDE)?%%\s*', '', l).split('|')]
        # OneNote 版删了「详见」行，不能当锚点（1.14-2）
        prev = [x for x in Ls[:i] if x.strip() and not re.match(r'\s*详见', x)][-1].strip(); nxt = [x for x in Ls[i + 1:] if x.strip()][0].strip()
        anchor = (re.sub(r'^#+\s*', '', nxt), 'before') if prev.startswith('<') else (re.sub(r'^#+\s*', '', prev), 'after')
        out.append(dict(file=NS + fn, cap=cap, wmm=float(wmm or 120), anchor=anchor))
    return out
def find_anchor(h, text):
    text = re.sub(r'\[\[(\d+\.\d+)[^\]]*\]\]', r'\1', text); text = re.sub(r'(\d+\.\d+)(\s*·\s*\1)+', r'\1', text)
    want = plain(text)[:18]
    for m in re.finditer(r'<p\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</p>', h, re.S):
        if want and want in plain(m.group(2)): return m.group(1)
def add_figs(pid, title):
    for fg in figs_for(title):
        h = content(pid, ids=True) or ''
        if 'alt="%s"' % html.escape(fg['cap']) in h: L('   图已有', fg['cap'][:16]); continue
        tid = find_anchor(h, fg['anchor'][0])
        if not tid: L('   !! 图找不到插入位置', fg['cap'][:16]); continue
        data = open(fg['file'], 'rb').read(); pw, ph = struct.unpack('>II', data[16:24])
        w = min(TABW, round(fg['wmm'] / 270 * TABW)); name = 'fig' + uuid.uuid4().hex[:8]
        c = ('<img src="name:%s" width="%d" height="%d" alt="%s"/>' % (name, w, round(w * ph / pw), html.escape(fg['cap'])) +
             '<p style="margin-top:0;margin-bottom:0"><span style="font-family:%s;font-size:9pt;color:#595959">%s</span></p>' % (FONT, html.escape(fg['cap'], quote=False)))
        try:
            patch_with_image(pid, [{'target': tid, 'action': 'insert', 'position': fg['anchor'][1], 'content': c}], name, data); L('   补图', fg['cap'][:16], w)
        except RuntimeError as e:
            L('   !! 补图出错（可能已插入，重跑会自动跳过）', fg['cap'][:16], str(e)[:40]); time.sleep(20)
        time.sleep(2)
# ---------- 3. 主流程 ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--notebook-id', default=NB_ID, help='默认：用户自己的「飞行」笔记本（另有同名共享只读本，必须按编号区分）')
    ap.add_argument('--group', default='', help='分区组名；留空表示分区直接放在笔记本下（2026-10-05 起）')
    ap.add_argument('--chapters', default='1-5')
    ap.add_argument('--only', default=''); ap.add_argument('--figs-only', action='store_true'); ap.add_argument('--audit-only', action='store_true')
    a = ap.parse_args()
    lo, hi = a.chapters.split('-'); want_ch = [NUM[str(n)] for n in range(int(lo), int(hi) + 1)]
    pages = [p for p in split(a.docx) if p['chap'] in want_ch]; L('拆页', len(pages), '节')
    secs, parent = locate(a.notebook_id, a.group)
    only = [x for x in a.only.split(',') if x]
    for p in pages:
        name = CHN[p['chap']]
        if name not in secs: secs[name] = req('POST', parent + '/sections', {'displayName': name})['id']; L('建分区', name)
        sid = secs[name]; existing = {q['title']: q['id'] for q in list_pages(sid)}
        if a.audit_only: continue
        if only and not any(p['title'].startswith(o + ' ') for o in only):
            if a.figs_only and p['title'] in existing: add_figs(existing[p['title']], p['title'])
            continue
        if a.figs_only:
            if p['title'] in existing: add_figs(existing[p['title']], p['title'])
            continue
        for q in list_pages(sid):
            if q['title'] == p['title']: delete_page(q['id'])
        src, _ = strip_index(p['html'])
        pid = create_page(sid, p['title'], page(p['title'], convert(src), '%s ｜ %s' % (name, p['title'])))
        L('写', p['title'], 'OK' if pid else '!! 失败')
        if pid: add_figs(pid, p['title'])
        time.sleep(6)                     # 2026-10-06：1.5 秒会触发 429 限流
    # 2026-10-09：新页有时标题为空（1009R4 写全书 27 页）——按页眉「章名 ｜ 节名」补标题；同名重复只留最新一页
    for name, sid in secs.items():
        for q in list_pages(sid):
            if q['title']: continue
            m = re.search(re.escape(name) + r' ｜ ([^<]+)', content(q['id']) or '')
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
    # 只读核查
    bad = 0
    for name, sid in secs.items():
        ps = list_pages(sid); readable = [q for q in ps if content(q['id'])]
        titles = [q['title'] for q in readable]
        want = [p['title'] for p in pages if CHN[p['chap']] == name]
        miss = [t for t in want if t not in titles]; dup = sorted({t for t in titles if titles.count(t) > 1})
        bad += len(miss) + len(dup)
        L('核查', name, '可读', len(readable), '/', len(ps), '缺', miss, '重复', dup)
    L('完成' if not bad else '!! 有缺页或重复，见上')
if __name__ == '__main__': main()
