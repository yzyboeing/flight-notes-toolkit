#!/usr/bin/env python3
"""把速查版 Word 写进 OneNote「基础知识速查区」分区（2026-10-08 用户：「分段分成约 10 页，加一个分区叫基础知识速查区」）。
按速查源的 %%PART%% 分段拆页（系统、一～八段、附录，共 10 页）；目录、按主题查不要；排版沿用完整版转换器（onenote_conv）。
用法（命令前加 ONENOTE_SRC_REF=baseline/<标签>，定宽 / 等宽标记从该基线的 速查/速查源.md 读）：
  python3 onenote_qr.py --docx 速查.docx --dry        # 只转换，结果写 ~/flight-repos/_work/onenote_qr/，报每页大小
  python3 onenote_qr.py --docx 速查.docx               # 写入（同名页先删后建）＋只读核查
  python3 onenote_qr.py --docx 速查.docx --only 系统   # 只重写这几页（接口不能调页序：其后的页要一起重写）
  python3 onenote_qr.py --docx 速查.docx --verify      # 读回逐页核对：全文、每格对齐与圆点、表宽"""
import argparse, html, os, re, subprocess, sys, time
os.environ['ONENOTE_SPEC'] = 'quickref'
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from onenote_api import req, content, list_pages, create_page, delete_page
from onenote_conv import convert, page, flat_lists
from onenote_build import NB_ID, locate, find_anchor
from onenote_api import patch_with_image
from onenote_conv import TABW, FONT
import struct, uuid
SECTION = '基础知识速查区'
REPO = os.path.expanduser('~/flight-repos/gh-private')
WORK = os.path.expanduser('~/flight-repos/_work/onenote_qr')
Z = r'[\s​⁠­ 　]+'
def L(*a): print(*a, flush=True)
def plain(h): return re.sub(Z, '', html.unescape(re.sub(r'<[^>]+>', '', h)))
def parts_from_src():
    ref = os.environ.get('ONENOTE_SRC_REF')
    s = (subprocess.run(['git', '-C', REPO, 'show', '%s:速查/速查源.md' % ref], capture_output=True, text=True, check=True).stdout if ref
         else open(os.path.join(REPO, '速查/速查源.md'), encoding='utf-8').read())
    out = []
    for l in s.split('\n'):
        m = re.match(r'%%PART%%\s*(.+)', l)
        if m: out.append([m.group(1).strip(), []]); continue
        m = re.match(r'## (.+)', l)
        if m and out: out[-1][1].append(m.group(1).strip())
    return out
def title_of(part, groups):
    return part + '　' + groups[0] if part == '附录' and len(groups) == 1 else part
def split(docx):
    os.makedirs(WORK, exist_ok=True)
    subprocess.run(['soffice', '--headless', '--convert-to', 'html', '--outdir', WORK, docx], check=True, capture_output=True)
    s = open(os.path.join(WORK, os.path.splitext(os.path.basename(docx))[0] + '.html'), encoding='utf-8').read()
    head = s[:s.find('<body')]
    end = s.find('<div title="footer"'); end = end if end > 0 else s.find('</body>')
    h2 = [(m.start(), plain(m.group(1))) for m in re.finditer(r'<h2\b[^>]*>(.*?)</h2>', s, re.S)]
    parts = parts_from_src()
    want = [plain(g) for _, gs in parts for g in gs]
    got = [t for _, t in h2]
    assert got == want, '速查 Word 的主题组与速查源对不上：%s' % [(a, b) for a, b in zip(got, want) if a != b][:3]
    pages, k = [], 0
    for part, gs in parts:
        a = h2[k][0]; k += len(gs); b = h2[k][0] if k < len(h2) else end
        body = re.sub(r'<p[^>]*>\s*(<br/>\s*)*</p>\s*$', '', s[a:b].rstrip())
        pages.append({'title': title_of(part, gs), 'html': head + '<body>' + body + '</body></html>'})
    return pages
def build(p): return page(p['title'], convert(p['html']), '%s ｜ %s' % (SECTION, p['title']))
# ---------- 读回核对（与 verify_pages.py 同口径）----------
def cells(h):
    h = flat_lists(h)          # 2026-10-09 圆点改 OneNote 自带列表：列表还原成段落再比
    out = []
    for tb in re.findall(r'<table\b.*?</table>', h, re.S):
        for ta, td in re.findall(r'<td\b([^>]*)>(.*?)</td>', tb, re.S):
            for a, p in (re.findall(r'<p\b([^>]*)>(.*?)</p>', td, re.S) or [(ta, td)]):
                t = plain(p)
                if not t: continue
                al = (re.search(r'text-align:\s*(\w+)', a) or [0, 'left'])[1]
                out.append((t[:12], al if al in ('center', 'right') else 'left', '•' in p or '●' in p))
    return out
def widths(h):
    return [sum(int(x) for x in re.findall(r'<td\b[^>]*?width:(\d+)', re.search(r'<tr\b.*?</tr>', tb, re.S).group(0)))
            for tb in re.findall(r'<table\b.*?</table>', h, re.S) if '<tr' in tb]
def verify(pages, sid):
    bad = 0
    allp = req('GET', '/sections/%s/pages?$select=id,title,createdDateTime&$top=100' % sid)['value']
    for p in pages:
        exp = convert(p['html'])
        got = sorted([q for q in allp if q['title'] == p['title']], key=lambda q: q['createdDateTime'], reverse=True)
        h = next((x for x in (content(q['id']) for q in got) if x), None)
        if not h: L('!! 读不到', p['title']); bad += 1; continue
        body = h[h.find('<body'):]
        e, g = plain(exp), plain(body).replace(plain('%s ｜ %s' % (SECTION, p['title'])), '', 1)
        probs = []
        if e not in g and g != e: probs.append('文字不同（期望 %d 字，读回 %d 字）' % (len(e), len(g)))
        ce, cg = cells(exp), cells(body)
        if ce != cg:
            probs.append('对齐/圆点不同 %s' % (next(((i, x, y) for i, (x, y) in enumerate(zip(ce, cg)) if x != y), ('长度', len(ce), len(cg))),))
        we, wg = widths(exp), widths(body)
        if len(we) != len(wg): probs.append('表数不同（%d / %d）' % (len(we), len(wg)))
        else:   # OneNote 每列可能少记 1px（取整），容差按列数放宽
            nc = [len(re.findall(r'<td\b', re.search(r'<tr\b.*?</tr>', tb, re.S).group(0))) for tb in re.findall(r'<table\b.*?</table>', exp, re.S) if '<tr' in tb]
            d = [(x, y) for x, y, n in zip(we, wg, nc) if abs(x - y) > max(3, n)]
            if d: probs.append('表宽不同 %s' % d[:3])
        L(('!! ' if probs else 'OK ') + p['title'], '；'.join(probs)); bad += bool(probs)
    L('核对完：问题页', bad)
    return bad
def qr_figs():
    """速查源里的插图行（2026-10-09 只有 DA / MDA 目视参考一张）→ [(条目标题, 文件, 图注, 宽 mm)]"""
    out, cur = [], None
    for l in open(os.path.join(REPO, '速查', '速查源.md'), encoding='utf-8').read().split('\n'):
        if l.startswith('### '): cur = l[4:].strip()
        elif l.startswith('%%FIG') and cur:
            fn, cap, wmm = [x.strip() for x in re.sub(r'^%%FIG(SIDE)?%%\s*', '', l).split('|')]
            out.append((cur, os.path.join(REPO, 'notes_src', fn), cap, float(wmm or 80)))
    return out
def add_qr_figs(pid):
    for title, fn, cap, wmm in qr_figs():
        h = content(pid, ids=True) or ''
        if plain(title) not in plain(h) or 'alt="%s"' % html.escape(cap) in h: continue
        # 2026-10-10：只认条目标题行（「编号. 标题」），不按前 18 字模糊匹配——表格里出现同名词时会插错位置（改进爬升图曾插进上一条表格的格子）
        want = plain(title)
        tid = next((m.group(1) for m in re.finditer(r'<p\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</p>', h, re.S)
                    if re.sub(r'^\d+[.．]', '', plain(m.group(2))) == want), None) or find_anchor(h, title)
        if not tid: L('   !! 图找不到插入位置', cap[:16]); continue
        data = open(fn, 'rb').read(); pw, ph = struct.unpack('>II', data[16:24])
        w = min(TABW, round(wmm / 270 * TABW * 1.6)); name = 'fig' + uuid.uuid4().hex[:8]
        c = ('<img src="name:%s" width="%d" height="%d" alt="%s"/>' % (name, w, round(w * ph / pw), html.escape(cap)) +
             '<p style="margin-top:0;margin-bottom:0"><span style="font-family:%s;font-size:9pt;color:#595959">%s</span></p>' % (FONT, html.escape(cap, quote=False)))
        patch_with_image(pid, [{'target': tid, 'action': 'insert', 'position': 'after', 'content': c}], name, data); L('   补图', cap[:16], w)
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True); ap.add_argument('--only', default='')
    ap.add_argument('--fresh', action='store_true', help='先删掉本分区全部页再按顺序重写（分段改名 / 重排后用）');ap.add_argument('--dry', action='store_true'); ap.add_argument('--verify', action='store_true'); ap.add_argument('--audit-only', action='store_true')
    a = ap.parse_args()
    pages = split(a.docx); L('拆页', len(pages), '页：', ' / '.join(p['title'] for p in pages))
    if a.dry:
        for i, p in enumerate(pages):
            out = build(p); f = os.path.join(WORK, '%02d_%s.html' % (i, re.sub(r'[/ 　]+', '_', p['title'])))
            open(f, 'w', encoding='utf-8').write(out)
            L('  %s：%d KB，表 %d 张，字 %d' % (p['title'], len(out.encode()) // 1024, out.count('<table border="1"'), len(plain(out))))
        return
    secs, parent = locate(NB_ID)
    if SECTION not in secs: secs[SECTION] = req('POST', parent + '/sections', {'displayName': SECTION})['id']; L('建分区', SECTION)
    sid = secs[SECTION]
    if a.verify: sys.exit(1 if verify(pages, sid) else 0)
    only = [x for x in a.only.split(',') if x]
    if a.fresh and not only and not a.audit_only:   # 2026-10-10：10 段重排后旧页名（03 起飞等）不会被同名删除覆盖，页序也乱
        for q in list_pages(sid): delete_page(q['id']); L('删旧页', q['title']); time.sleep(3)
    if not a.audit_only:
        for p in pages:
            if only and p['title'] not in only and not any(p['title'].startswith(o) for o in only): continue
            for q in list_pages(sid):
                if q['title'] == p['title']: delete_page(q['id'])
            pid = create_page(sid, p['title'], build(p))
            L('写', p['title'], 'OK' if pid else '!! 失败')
            if pid:
                time.sleep(8)
                try: add_qr_figs(pid)
                except RuntimeError as e: L('!! 补图失败（稍后 --only 该页重写）', p['title'], str(e)[:60])   # 2026-10-10：504 曾使整区写入中断
            time.sleep(10)
    for q in list_pages(sid):          # 新页有时标题为空：按页眉「基础知识速查区 ｜ 页名」补（2026-10-08 系统、七两页）
        if q['title']: continue
        m = re.search(SECTION + r' ｜ ([^<]+)', content(q['id']) or '')
        if m: req('PATCH', '/pages/%s/content' % q['id'], [{'target': 'title', 'action': 'replace', 'content': m.group(1).strip()}]); L('补标题', m.group(1).strip())
    if not a.audit_only: time.sleep(20)
    ps = list_pages(sid); readable = [q for q in ps if content(q['id'])]
    titles = [q['title'] for q in readable]
    miss = [p['title'] for p in pages if p['title'] not in titles]; dup = sorted({t for t in titles if titles.count(t) > 1})
    L('核查', SECTION, '可读', len(readable), '/', len(ps), '缺', miss, '重复', dup)
    L('完成' if not (miss or dup) else '!! 有缺页或重复，见上')
if __name__ == '__main__': main()
