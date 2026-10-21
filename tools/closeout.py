#!/usr/bin/env python3
"""closeout.py —— 一键收口（规则重组方案 A5，2026-10-21）：全量重建后 T13 / T14 收不拢的表，一条命令锁宽并清旧记录。

用法：python3 closeout.py --repo <gh-private> [--apply]
  1. t13_lock.py：报告里 T13（列没折行、右侧仍空）的表按实测收窄、表头写 w-NN + w-lock；
  2. fixw_lock.py：T14（有空位却折行）来回拉锯的表按实测加宽并锁住；
  3. 被锁的表在 build/keep_force_*.txt 里的旧 W / NW / BW / F / C 记录一并清掉（总纲第八节第 7 步原来要手工做）。
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
            elif l.startswith('BW:'): sg = l.split(':', 4)[4]
            elif l.startswith('C:'): sg = l.split(':', 2)[2].rsplit('|', 1)[0]
            (drop if sg is not None and any(N(sg).startswith(k) or k.startswith(N(sg)[:30]) for k in keys) else keep).append(l)
        for l in drop: print('  清 %s：%s' % (os.path.basename(kf), l[:70]))
        n += len(drop)
        if apply and drop: io.open(kf, 'w', encoding='utf-8').write('\n'.join(keep) + '\n')
    return n


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
    print('== keep_force 旧记录（%d 张表被锁）' % len(keys))
    n = purge_keep_force(repo, keys, a.apply) if keys else 0
    print('收口：锁 %d 张表，清 %d 条旧记录%s' % (len(keys), n, '' if a.apply else '（未加 --apply，只列出）'))


if __name__ == '__main__':
    main()
