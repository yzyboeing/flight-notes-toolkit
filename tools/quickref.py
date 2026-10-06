#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""quickref.py —— 生成《B737 机型理论知识速查》（SD-146，2026-10-06 用户；取代 SD-145 的自动摘句版，旧脚本留作 quickref_auto_sd145.py）

用户要求：
  · 「这个形式的速查笔记，不分章节，只分知识点」——参照 1005R2 时的《理论基础知识速查》：一条一个知识点，只有数据、限制、概念，没有多余解释；
  · 「目录可以参考总笔记的，按主题查、按飞行阶段查，或者按运行环境查，包括非正常与应急、限制与规章等」；
  · 正文顺序由 Claude 评估后定：正文按主题分组（只用主题小标题，不出章节号），五种查法都做成目录；
  · 随完整版更新，用户要时再给；噜噜一起核对。

源：gh-private/速查/速查源.md
  「## 主题」分组（主题名取自 按主题查索引.json）；「### 标题」一条知识点，不写编号（生成时全书连续编号）；
  每条第一行 <!-- 详见 [[x.y 节名|x.y A-n]] --> 绑定正文块：决定五维目录归类，正文改动时提示联动（check_quickref / quickref_sync_hint）。
产物：build/qr_single.md、build/qr_dims.json、build/B737机型理论知识速查.docx / .pdf（不进 git）。
用法：python3 quickref.py [--repo ~/flight-repos/gh-private]
"""
import sys, os, re, json, subprocess, shutil

T = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1]) if '--repo' in sys.argv else os.path.expanduser('~/flight-repos/gh-private')
SRC = os.path.join(REPO, '速查', '速查源.md')
IDX = os.path.join(REPO, '按主题查索引.json')
BUILD = os.path.join(REPO, 'build')
SINGLE_MD = os.path.join(BUILD, 'qr_single.md')
DIMS_JSON = os.path.join(BUILD, 'qr_dims.json')
OUT_DOCX = os.path.join(BUILD, 'B737机型理论知识速查.docx')

def parse_src():
    """→ [(主题, [ {t, body:[行], binds:[(节, 块)]} ])]"""
    groups, cur_g, cur_e = [], None, None
    for l in open(SRC, encoding='utf-8').read().split('\n'):
        if l.startswith('# ') or (l.startswith('<!--') and cur_e is None): continue
        m = re.match(r'^## (.+)$', l)
        if m:
            cur_g = (m.group(1).strip(), []); groups.append(cur_g); cur_e = None; continue
        m = re.match(r'^### (.+)$', l)
        if m and cur_g is not None:
            cur_e = {'t': m.group(1).strip(), 'body': [], 'binds': []}; cur_g[1].append(cur_e); continue
        if cur_e is not None:
            if l.startswith('<!-- 详见'):
                for a in re.findall(r'\[\[([^\]|]*)\|?([^\]]*)\]\]', l):
                    lab = re.sub(r'第\s*(\d+)\s*条', r'\1', (a[1] or a[0]).strip())
                    mm = re.match(r'(\d+\.\d+)(?:\s+([A-H]-\d+|\d+))?', lab)
                    if mm: cur_e['binds'].append((mm.group(1), mm.group(2)))
            cur_e['body'].append(l)
    return [g for g in groups if g[1]]

def main():
    groups = parse_src()
    # 1) 组装单册源：全书连续编号；「## 主题」→「### 主题」（单册目录列主题），「### 标题」→「#### N. 标题」
    out = ['# 机型理论知识速查', '', '## 第零章　速查', '', '%%COMPACT%%', '']   # 单册模式只认「第X章」触发目录页；章名本身不印（页脚用册名 DOC_HEADER）
    n = 0; ents = []
    for g, es in groups:
        out += ['### ' + g, '']
        for e in es:
            n += 1; e['n'] = n; e['g'] = g; ents.append(e)
            out += ['#### %d. %s' % (n, e['t'])] + [x for x in e['body'] if not x.startswith('<!--')] + ['', '']
    out += ['%%ENDCOMPACT%%', '']
    os.makedirs(BUILD, exist_ok=True)
    open(SINGLE_MD, 'w', encoding='utf-8').write('\n'.join(out))
    # 2) 五维目录：条目绑定的正文块落在哪个主题，就挂到哪个主题下（一条可挂多处）；正文分组主题本身也挂
    idx = json.load(open(IDX, encoding='utf-8'))
    clean = lambda t: re.sub(r'（单位：[^）]*）|【[^】]*】', '', t).strip()
    dims = []
    for d in idx:
        themes = []
        for tp in d['themes']:
            keys = {x.split('|')[0].strip() for x in tp['items']}
            hit = [e for e in ents if any(('%s %s' % (s, b)) in keys for s, b in e['binds'] if b) or e['g'] == tp['theme']]
            if hit: themes.append({'theme': tp['theme'], 'items': ['%d|%s' % (e['n'], clean(e['t'])) for e in sorted(hit, key=lambda x: x['n'])]})
        if themes: dims.append({'dim': d['dim'], 'themes': themes})
    json.dump(dims, open(DIMS_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    # 3) 排版：与完整版同一套生成器与 fit_fix（整表同页、列宽实测、孤字），单册模式
    env = dict(os.environ)
    for k in ('DOC_TOPICS', 'DOC_PREFACE', 'DOC_QRTOPICS', 'DOC_TOPICINDEX', 'NO_SEC_BREAK'): env.pop(k, None)
    env.setdefault('NODE_PATH', os.path.join(REPO, 'node_modules'))
    git_cfg = lambda k: subprocess.run(['git', '-C', REPO, 'config', '--get', k], capture_output=True, text=True).stdout.strip()
    for k, g in (('DOC_EDITION', 'notes.docEdition'), ('DOC_NOTICE', 'notes.docNotice')):   # 封面版本号、特别提示与全书同（sync 已导出；单独跑时从 git config 取）
        if not env.get(k): env[k] = git_cfg(g)
    env.update({'DOC_SINGLE': '1', 'DOC_QRDIMS': DIMS_JSON, 'DOC_SINGLE_TOC_TITLE': '目录', 'DOC_SUBTITLE': '数据 · 限制 · 概念', 'DOC_HEADER': 'B737机型理论知识速查'})
    r = subprocess.run([sys.executable, os.path.join(T, 'fit_fix.py'), SINGLE_MD, OUT_DOCX], env=env, cwd=REPO, capture_output=True, text=True)
    if r.returncode: sys.exit('速查版排版失败：' + (r.stderr or r.stdout)[-600:])
    sof = shutil.which('soffice') or '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    prof = os.path.join(os.environ.get('TMPDIR', '/tmp'), 'lo-sync-profile')
    subprocess.run([sof, '-env:UserInstallation=file://' + prof, '--headless', '--convert-to', 'pdf', '--outdir', BUILD, OUT_DOCX], capture_output=True)
    pdf = OUT_DOCX[:-5] + '.pdf'
    try:   # 与 sync.sh 全书同：打开即展开书签栏；书签去掉「（单位：…）」，正文标题保留
        import pymupdf
        d = pymupdf.open(pdf); d.set_pagemode('UseOutlines'); toc = d.get_toc(simple=False)
        for e in toc: e[1] = re.sub(r'\s*（单位：(?:[^（）]|（[^（）]*）)*）', '', e[1])
        d.set_toc(toc); np = len(d)
        d.save(pdf + '.tmp', garbage=3, deflate=True); d.close(); os.replace(pdf + '.tmp', pdf)
    except Exception: np = '?'
    print('速查版：%d 个主题、%d 条知识点，%s 页 → %s' % (len(groups), n, np, pdf))

if __name__ == '__main__': main()
