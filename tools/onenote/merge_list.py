#!/usr/bin/env python3
"""merge_list.py —— 由基线 docx 生成《OneNote合并单元格清单.md》（OneNote 接口写不了合并，由用户手动合并）。
用法：python3 merge_list.py --docx <基线docx> --out <清单.md> [--base 基线号]
按页 → 表（最近的小标题 + 表头）→ 「第 r 行第 c 列『文字』：向下跨 n 行 / 向右跨 n 列」列出；行号从表头算起（表头为第 1 行）。"""
import sys, re, html, argparse
from html.parser import HTMLParser
sys.path.insert(0, __import__('os').path.dirname(__file__))
from onenote_build import split

CHN = {1: '第一章 系统理论', 2: '第二章 机组训练手册', 3: '第三章 运行手册', 4: '第四章 模拟机训练', 5: '第五章 技术提示'}

class P(HTMLParser):
    def __init__(s):
        super().__init__(); s.tabs = []; s.head = '（节首）'; s.stack = []; s.htxt = None; s.cell = None
    def handle_starttag(s, t, a):
        a = dict(a)
        if re.fullmatch(r'h[1-6]', t): s.htxt = ''
        elif t == 'table': s.stack.append({'head': s.head, 'rows': []})
        elif t == 'tr' and s.stack: s.stack[-1]['rows'].append([])
        elif t in ('td', 'th') and s.stack:
            s.cell = {'cs': int(a.get('colspan') or 1), 'rs': int(a.get('rowspan') or 1), 't': ''}
        elif t == 'br' and s.cell is not None: s.cell['t'] += ' '
    def handle_endtag(s, t):
        if re.fullmatch(r'h[1-6]', t) and s.htxt is not None:
            x = re.sub(r'\s+', ' ', s.htxt).strip()
            if x and not re.match(r'^\d+\.\d+\s', x): s.head = x
            s.htxt = None
        elif t in ('td', 'th') and s.cell is not None and s.stack:
            s.stack[-1]['rows'][-1].append(s.cell); s.cell = None
        elif t == 'table' and s.stack: s.tabs.append(s.stack.pop())
    def handle_data(s, d):
        if s.htxt is not None: s.htxt += d
        if s.cell is not None: s.cell['t'] += d

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--docx', required=True); ap.add_argument('--out', required=True); ap.add_argument('--base', default='')
    a = ap.parse_args()
    out, n = [], 0
    for pg in split(a.docx):
        p = P(); p.feed(pg['html']); sec = []
        for tb in p.tabs:
            grid, items = {}, []
            for r, row in enumerate(tb['rows']):
                c = 0
                for cell in row:
                    while (r, c) in grid: c += 1
                    for dr in range(cell['rs']):
                        for dc in range(cell['cs']): grid[(r + dr, c + dc)] = 1
                    txt = re.sub(r'\s+', ' ', html.unescape(cell['t'])).strip()[:18]
                    if cell['rs'] > 1: items.append('- [ ] 第 %d 行第 %d 列「%s」：向下跨 %d 行' % (r + 1, c + 1, txt, cell['rs']))
                    if cell['cs'] > 1: items.append('- [ ] 第 %d 行第 %d 列「%s」：向右跨 %d 列' % (r + 1, c + 1, txt, cell['cs']))
                    c += cell['cs']
            if items and len(tb['rows']) > 0:
                hdr = ' | '.join(re.sub(r'\s+', ' ', html.unescape(x['t'])).strip()[:12] for x in tb['rows'][0])
                sec += ['', '**%s**（表头：%s）' % (tb['head'], hdr), ''] + items; n += len(items)
        if sec: out += ['', '## %s ｜ %s' % (CHN.get(int(pg['title'].split('.')[0]), pg['chap']), pg['title'])] + sec
    head = ['# OneNote 合并单元格清单（基线 %s）' % a.base, '',
            '> 用法：每条在 OneNote 里找到对应的表，选中「起始格」和它后面的空格（向下跨 N 行 = 选中该格及下面 N-1 格；向右跨 N 列 = 选中该格及右边 N-1 格），表格 → 合并 → 合并单元格。行号从表头算起（表头为第 1 行）。',
            '共 %d 处。' % n, '']
    open(a.out, 'w', encoding='utf-8').write('\n'.join(head + out) + '\n'); print('共', n, '处 →', a.out)
if __name__ == '__main__': main()
