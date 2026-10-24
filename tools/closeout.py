#!/usr/bin/env python3
"""closeout.py —— 一键收口（规则重组方案 A5，2026-10-21）：全量重建后 T13 / T14 收不拢的表，一条命令锁宽并清旧记录。

用法：python3 closeout.py --repo <gh-private> [--apply]
  1. t13_lock.py：报告里 T13（列没折行、右侧仍空）的表按实测收窄、表头写 w-NN + w-lock；
  2. fixw_lock.py：T14（有空位却折行）来回拉锯的表按实测加宽并锁住；
  3. 被锁的表在 build/keep_force_*.txt 里的旧 W / NW / BW / F / C 记录一并清掉（总纲第八节第 7 步原来要手工做）；
  4. 报告里「fit_fix 未加宽」的 T14（多为没定宽的表被旧收窄 / 加宽放弃记录挡住）：按页码找到那张表，清它的旧记录（含 WB / NWL / BWB）。
不加 --apply 只列出要改什么；加了才写源文件和 keep_force。改完照常 preview / sync --full 重建确认。"""
import argparse, glob, io, os, re, subprocess, sys

T = os.path.dirname(os.path.abspath(__file__))
N = lambda s: re.sub(r'[^\w]', '', re.sub(r'<[^>]+>', '', s.replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&')))


def sources(repo):
    return sorted(glob.glob(os.path.join(repo, 'notes_src', '[1-9]*', '[0-9]*.md'))) + [os.path.join(repo, '速查', '速查源.md')]


def tables(text):
    """[(表头, 表文字签名)]——签名与 fit_fix 记录同口径：去标签、去标点；去掉前提 / 注 / 警示行"""
    out = []
    for m in re.finditer(r'<table[^>]*>.*?</table>', text, re.S):
        hdr = re.search(r'<tr class="hdr">.*?</tr>', m.group(0), re.S)
        body = re.sub(r'<tr class="(premise|note|warn)">.*?</tr>', '', m.group(0), flags=re.S)
        out.append((hdr.group(0) if hdr else '', N(body)))
    return out


def purge_keep_force(repo, keys, apply):
    """清掉签名以 keys 中任一开头的 W / NW / BW / F / C 记录（整页「~N」「!」与 P: 不动）"""
    n = 0
    for kf in glob.glob(os.path.join(repo, 'build', 'keep_force_*.txt')):
        keep, drop = [], []
        for l in io.open(kf, encoding='utf-8').read().splitlines():
            sg = None
            if re.match(r'(W|NW|F):\d+:-?\d+:', l): sg = l.split(':', 3)[3]
            elif re.match(r'(WB|NWL):\d+:', l) and not re.match(r'(WB|NWL):\d+:-?\d+:', l): sg = l.split(':', 2)[2]   # 加宽放弃 / 收窄极限（SD-211 后本次运行有效，残留时一并清）
            elif l.startswith('BWB:'): sg = l.split(':', 2)[2]
            elif l.startswith('BW:'): sg = l.split(':', 4)[4]
            elif l.startswith('C:'): sg = l.split(':', 2)[2].rsplit('|', 1)[0]
            (drop if sg is not None and any(N(sg).startswith(k) or k.startswith(N(sg)[:30]) for k in keys) else keep).append(l)
        for l in drop: print('  清 %s：%s' % (os.path.basename(kf), l[:70]))
        n += len(drop)
        if apply and drop: io.open(kf, 'w', encoding='utf-8').write('\n'.join(keep) + '\n')
    return n


def t14_tables(repo):
    """报告里「fit_fix 未加宽」的 T14：取那一页的文字，keep_force 里签名前 12 字（去标点）出现在该页的表 → 签名前 30 字"""
    import pymupdf
    rp = os.path.join(repo, 'build', '排版检查报告.md')
    if not os.path.exists(rp): return set()
    hits = re.findall(r'^- \[T14\] (全书|单册) 第 (\d+) 页：表格右侧有空位.*?fit_fix 未加宽', io.open(rp, encoding='utf-8').read(), re.M)
    out = set()
    for bk, pg in hits:
        stem = 'B737机型理论速查笔记' if bk == '全书' else 'B737机型理论速查'
        pdf = os.path.join(repo, 'build', stem + '.pdf'); kf = os.path.join(repo, 'build', 'keep_force_%s.txt' % stem)
        if not (os.path.exists(pdf) and os.path.exists(kf)): continue
        txt = N(pymupdf.open(pdf)[int(pg) - 1].get_text())
        for l in io.open(kf, encoding='utf-8').read().splitlines():
            m = re.match(r'(?:(?:W|NW|F):\d+:-?\d+:|(?:WB|NWL|BWB):\d+:|BW:\d+:\d+:-?\d+:)(.+)', l)
            if m and len(N(m.group(1))) >= 12 and N(m.group(1))[:12] in txt: out.add(N(m.group(1))[:30])
    return out


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--repo', default='.'); ap.add_argument('--apply', action='store_true'); a = ap.parse_args()
    repo = os.path.abspath(a.repo)
    before = {f: io.open(f, encoding='utf-8').read() for f in sources(repo) if os.path.exists(f)}
    for tool in ('t13_lock.py', 'fixw_lock.py'):
        print('== %s' % tool)
        r = subprocess.run([sys.executable, os.path.join(T, tool), '--repo', repo] + (['--apply'] if a.apply else []), capture_output=True, text=True)
        print((r.stdout + r.stderr).rstrip())
        if r.returncode: sys.exit('%s 失败' % tool)
    keys = set()
    for f, old in before.items():
        new = io.open(f, encoding='utf-8').read()
        if new == old: continue
        for (h0, s0), (h1, s1) in zip(tables(old), tables(new)):
            if h0 != h1: keys.add(s1[:30])
    # 2026-10-21：T14「fit_fix 未加宽」多为没定宽的表被旧收窄 / 加宽放弃记录挡住（fixw_lock 只管定宽表，1021R1 / R2 两次手工清）：
    # 按报告页码找到那一页的表，清它的旧记录，让 fit_fix 下次重建重新加宽
    t14 = t14_tables(repo)
    if t14: print('== T14 未加宽的表（清旧记录，%d 张）' % len(t14)); keys |= t14
    print('== keep_force 旧记录（%d 张表被锁 / 清）' % len(keys))
    n = purge_keep_force(repo, keys, a.apply) if keys else 0
    print('收口：锁 %d 张表，清 %d 条旧记录%s' % (len(keys), n, '' if a.apply else '（未加 --apply，只列出）'))


if __name__ == '__main__':
    main()
