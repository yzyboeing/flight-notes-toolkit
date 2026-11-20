#!/usr/bin/env python3
"""生成工作区《现状一页纸.md》与《README.md》（2026-11-20 规则瘦身：两份都自动生成，不再手写）。

用法：python3 ~/flight-repos/pub/tools/status_page.py [--ws <工作区>] [--check]
- 只读各仓库、快照与上传日志，不改任何源文件；输出写到工作区根目录。
- 工作区要有两份输入：`README_模板.md`（固定说明，含 {{生成时间}} {{现状}} {{人工项}} 三个占位）、
  `状态_人工项.md`（脚本读不出来、由人决定的发布通道、否决项、下一步）。
- `--check`：只比对现有两份文件除生成时间外是否与重新生成的一致，不一致退出码 1。
工作区默认 ~/Projects/737机型笔记/工作资料/飞行理论笔记整理，可用 --ws 或环境变量 STATUS_WS 改。
"""
import argparse, datetime, glob, os, re, subprocess

HOME = os.path.expanduser('~')
FLIGHT = os.path.join(HOME, 'flight-repos')
G = os.path.join(FLIGHT, 'gh-private')
STUDY = os.path.join(FLIGHT, 'study')
PUB = os.path.join(FLIGHT, 'pub')
DEFAULT_WS = os.path.join(HOME, 'Projects', '737机型笔记', '工作资料', '飞行理论笔记整理')


def sh(*a, cwd=None):
    try:
        return subprocess.run(a, cwd=cwd, capture_output=True, text=True, timeout=120).stdout.strip()
    except Exception:
        return ''


def read(path):
    with open(path, encoding='utf-8') as f:
        return f.read()


def repo(path):
    head = sh('git', 'log', '-1', '--format=%h|%ad|%s', '--date=format:%m-%d %H:%M', cwd=path).split('|', 2)
    dirty = [l for l in sh('git', 'status', '--porcelain', cwd=path).splitlines() if l]
    ahead = sh('git', 'rev-list', '--count', '@{u}..HEAD', cwd=path) or '?'
    return (head + ['', '', ''])[:3], len(dirty), ahead


def scale():
    secs = sorted(f for f in glob.glob(os.path.join(G, 'notes_src', '[1-9] *', '*.md'))
                  if re.match(r'\d+\.\d+ ', os.path.basename(f)))
    chapters = sorted({os.path.basename(os.path.dirname(f)) for f in secs})
    entries = sum(len(re.findall(r'(?m)^### \d+\. ', read(f))) for f in secs)
    refs = set()
    for f in secs:
        t = read(f)
        refs.update(re.findall(r'%%FIG%%\s*figures/([^\s|]+)', t))
        refs.update(re.findall(r'<img[^>]*src="figures/([^"]+)"', t))
    fig_dir = os.path.join(G, 'notes_src', 'figures')
    dark = sum(1 for r in refs if os.path.exists(os.path.join(fig_dir, re.sub(r'\.png$', '.dark.png', r))))
    return chapters, len(secs), entries, len(refs), dark


def checks():
    out = sh('python3', os.path.join(PUB, 'tools', 'check_src.py'), '--src', 'notes_src', cwd=G)
    errs = len(re.findall(r'(?m)^\s*错误\s', out))
    warns = len(re.findall(r'(?m)^\s*警告\s', out))
    hints = re.findall(r'(?m)^\s*提示\s+K-(S\d+)[^\n]*?(\d+) 处', out)
    return errs, warns, '、'.join('K-%s %s' % h for h in hints if h[1] != '0') or '无'


def app_status():
    content = sh('git', 'log', '-1', '--format=%H', 'HEAD', '--', 'notes_src', cwd=G)
    matching = []
    for meta in glob.glob(os.path.join(FLIGHT, '_work', '学习系统快照_*', '基线.txt')):
        revision = re.search(r'(?m)^提交：gh-private ([0-9a-f]+)', read(meta))
        if content and revision and revision[1] == content:
            matching.append(meta)
    snapshot = os.path.basename(os.path.dirname(max(matching, key=os.path.getmtime))) if matching else '尚无与当前内容一致的快照'
    uploads = []
    for log in glob.glob(os.path.join(FLIGHT, '_work', 'archives', '737note_*_upload.log')):
        version = re.fullmatch(r'737note_(\d+(?:\.\d+)*)_(\d+)_upload\.log', os.path.basename(log))
        if not version:
            continue
        result = read(log)
        if 'EXPORT SUCCEEDED' in result or 'Upload succeeded' in result:
            uploads.append((int(version[2]), version[1], os.path.getmtime(log)))
    tf = '本机无成功上传记录'
    if uploads:
        number, version, mtime = max(uploads)
        tf = '%s (%d) 已上传（本机上传日志 %s；Apple 处理状态与加组情况见人工项）' % (
            version, number, datetime.datetime.fromtimestamp(mtime).strftime('%m-%d %H:%M'))
    pbx = os.path.join(STUDY, 'app', 'ios', 'App', 'App.xcodeproj', 'project.pbxproj')
    build = re.search(r'CURRENT_PROJECT_VERSION = (\d+);', read(pbx)) if os.path.exists(pbx) else None
    return content[:8] or '未知', snapshot, tf, build[1] if build else '未知'


def latest_sd():
    nums = []
    for p in (os.path.join(PUB, 'prompt', 'archive', '2026-11-20_规则瘦身前', 'standing-decisions.md'),
              os.path.join(G, '规则', '决策记录.md')):
        if os.path.exists(p):
            nums += [int(x) for x in re.findall(r'(?m)^## SD-(\d+)', read(p))]
    return max(nums) if nums else 0


def last_line(path, pattern):
    if not os.path.exists(path):
        return '（没有这个文件）'
    hits = [l for l in read(path).splitlines() if re.match(pattern, l)]
    return hits[-1] if hits else '（无）'


def status_block(ws, now):
    content, snapshot, tf, build = app_status()
    chapters, nsec, nent, nfig, ndark = scale()
    errs, warns, hints = checks()
    rules = os.path.join(G, '规则', '现行规则.md')
    rules_kb = os.path.getsize(rules) / 1024 if os.path.exists(rules) else 0
    L = ['## 现状（%s 自动生成）' % now, '',
         '| 项目 | 当前值 |', '|---|---|',
         '| 正文规模 | %d 章 %d 节 %d 条（%s） |' % (len(chapters), nsec, nent, '、'.join(c.split(' ', 1)[-1] for c in chapters)),
         '| 正文引用插图 | %d 张，其中 %d 张有原生深色版 `.dark.png` |' % (nfig, ndark),
         '| 当前内容提交 | gh-private `%s` |' % content,
         '| 与当前内容一致的快照 | `%s` |' % snapshot,
         '| study 构建号 | %s |' % build,
         '| 最近成功上传的 TestFlight | %s |' % tf,
         '| 源检查（check_src） | 错误 %d、警告 %d；提示：%s |' % (errs, warns, hints),
         '| 现行规则 | `gh-private/规则/现行规则.md` %.1f KB（上限 30 KB）；决策记录编到 SD-%d |' % (rules_kb, latest_sd()),
         '', '### 四库', '', '| 仓库 | 最新提交 | 未提交 | 未推送 |', '|---|---|---|---|']
    for name, path in (('gh-private（私有笔记库）', G), ('study（App）', STUDY), ('pub（公开工具库）', PUB), ('工作区', ws)):
        (h, d, s), dirty, ahead = repo(path)
        L.append('| %s | `%s` %s %s | %d | %s |' % (name, h, d, s[:60], dirty, ahead))
    L += ['', '### 最近一条交接与审核', '',
          '- 交接：' + last_line(os.path.join(ws, '00_先读我_AI交接.md'), r'- ')[:240],
          '- 审核：' + last_line(os.path.join(ws, '审核记录.md'), r'## ')[3:200], '']
    return '\n'.join(L)


def manual_block(ws):
    p = os.path.join(ws, '状态_人工项.md')
    if not os.path.exists(p):
        return '## 人工项\n\n（缺 `状态_人工项.md`）\n'
    body = read(p).split('\n', 1)[-1]
    body = re.sub(r'(?s)^\s*>.*?\n\n', '', body, count=1)   # 去掉说明引言
    return body.strip() + '\n'


def strip_time(text):
    return re.sub(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}', '<时间>', text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--ws', default=os.environ.get('STATUS_WS', DEFAULT_WS))
    ap.add_argument('--check', action='store_true')
    a = ap.parse_args()
    ws = a.ws
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    status, manual = status_block(ws, now), manual_block(ws)
    page = '\n'.join(['# 现状一页纸', '',
                      '> 自动生成（`pub/tools/status_page.py`，%s），不要手改。规则看 `gh-private/规则/现行规则.md`；人工决定的几项改 `状态_人工项.md`。' % now,
                      '', status, manual])
    tmpl_path = os.path.join(ws, 'README_模板.md')
    readme = read(tmpl_path).replace('{{生成时间}}', now).replace('{{现状}}', status).replace('{{人工项}}', manual) \
        if os.path.exists(tmpl_path) else None
    outputs = [('现状一页纸.md', page)] + ([('README.md', readme)] if readme else [])
    if a.check:
        bad = [n for n, t in outputs if not os.path.exists(os.path.join(ws, n))
               or strip_time(read(os.path.join(ws, n))) != strip_time(t)]
        print('已是最新' if not bad else '需要重跑 status_page.py：' + '、'.join(bad))
        raise SystemExit(1 if bad else 0)
    for name, text in outputs:
        with open(os.path.join(ws, name), 'w', encoding='utf-8') as f:
            f.write(text)
    print('已生成 ' + '、'.join(n for n, _ in outputs) + ('' if readme else '（缺 README_模板.md，README 未生成）'))


if __name__ == '__main__':
    main()
