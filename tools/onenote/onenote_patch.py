#!/usr/bin/env python3
"""按段落更新 OneNote 页（不整页重写，保留页序和已插入的图）。
用法：python3 onenote_patch.py --docx 新基线.docx --sections 1.5,3.11 [--dry]
做法：用新基线重新生成这几节的接口 HTML，和 OneNote 现有页面逐段（正文框里的顶层 <p>）对比：
  文字变了的段 → replace；新版没有的段 → 换成空段；表格、图片不动。
表格内容有变化时本工具不处理，需按《OneNote写入流程.md》重写该节及同分区其后各节。"""
import argparse, difflib, html, os, re, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from onenote_api import req, content, list_pages
from onenote_build import split, CHN, NB_ID, locate
from onenote_conv import convert, strip_index
def pt(s): return re.sub(r'[\s　 ]+', '', html.unescape(re.sub(r'<[^>]+>', '', s)))
def top_ps(h):
    """顶层段落（不在表格里）：返回 [(id 或 None, 原 html, 纯文本)]"""
    out, depth = [], 0
    for m in re.finditer(r'<table\b|</table>|<p\b[^>]*>.*?</p>', h, re.S):
        g = m.group(0)
        if g.startswith('<table'): depth += 1; continue
        if g == '</table>': depth -= 1; continue
        if depth == 0:
            i = re.search(r'\bid="([^"]+)"', g); out.append((i.group(1) if i else None, g, pt(g)))
    return out
def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--docx', required=True); ap.add_argument('--sections', required=True)
    ap.add_argument('--notebook-id', default=NB_ID); ap.add_argument('--group', default=''); ap.add_argument('--dry', action='store_true')
    a = ap.parse_args(); want = [x.strip() for x in a.sections.split(',') if x.strip()]
    pages = [p for p in split(a.docx) if p['chap'] in CHN and any(p['title'].startswith(w + ' ') for w in want)]
    secs, _ = locate(a.notebook_id, a.group)
    for p in pages:
        sid = secs[CHN[p['chap']]]
        cands = [q for q in list_pages(sid) if q['title'] == p['title'] and content(q['id'])]
        if len(cands) != 1: print('!!', p['title'], '可读页数', len(cands), '跳过'); continue
        pid = cands[0]['id']; cur = top_ps(content(pid, ids=True))
        cur = [c for c in cur if c[2] and not c[2].startswith(pt(CHN[p['chap']])) and not re.match(r'图\d+\.\d+-\d+', c[2])]   # 去掉页眉、空段、补图时加的图注
        src, _ = strip_index(p['html']); new = [n for n in top_ps(convert(src)) if n[2]]
        sm = difflib.SequenceMatcher(a=[c[2] for c in cur], b=[n[2] for n in new], autojunk=False)
        cmds = []
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == 'equal': continue
            if op == 'insert':
                if i1 > 0: tgt, pos, order = cur[i1 - 1][0], 'after', reversed(range(j1, j2))
                elif cur: tgt, pos, order = cur[0][0], 'before', range(j1, j2)
                else: print('  !! 新增段落找不到插入位置', p['title'], new[j1][2][:30]); continue
                for j in order:
                    cmds.append({'target': tgt, 'action': 'insert', 'position': pos, 'content': new[j][1]}); print('  增', p['title'], new[j][2][:30])
                continue
            for k in range(i1, i2):
                c = cur[k]
                if not c[0]: continue
                j = j1 + (k - i1)
                if op == 'replace' and j < j2:
                    cmds.append({'target': c[0], 'action': 'replace', 'content': new[j][1]}); print('  改', p['title'], c[2][:24], '→', new[j][2][:24])
                else:
                    cmds.append({'target': c[0], 'action': 'replace', 'content': '<p style="margin-top:0;margin-bottom:0"><span style="font-size:4pt">&#160;</span></p>'}); print('  删', p['title'], c[2][:30])
        if cmds and not a.dry:
            req('PATCH', '/pages/%s/content' % pid, cmds); time.sleep(2)
        print(p['title'], '段落改动', len(cmds), '(预演)' if a.dry else '')
if __name__ == '__main__': main()
