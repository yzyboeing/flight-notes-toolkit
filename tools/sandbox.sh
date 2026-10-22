#!/usr/bin/env bash
# sandbox.sh —— 一键沙盒（2026-10-21：SD-222 试改动时手搭沙盒试了 4 次才跑通——缺 node_modules、HOME 不对会写桌面交接包 / 云盘、
#               找不到用户目录里的 PyMuPDF）。在临时目录建两库的 worktree，构建完全不碰正式库、桌面和云盘。
#
# 用法：sandbox.sh new  <目录>            # 建沙盒：<目录>/gh-private、<目录>/pub（当前 HEAD，分离头指针）
#       sandbox.sh run  <目录> <命令…>    # 在沙盒 gh-private 里跑命令（隔离 HOME / TMPDIR），如：sandbox.sh run D ./sync.sh --full "试"
#       sandbox.sh drop <目录>            # 删 worktree（沙盒目录移到 ~/flight-repos/_to_delete）
# 改工具：直接改 <目录>/pub/tools/…；验证通过后 git -C <目录>/pub diff > x.diff，再到正式 pub 里 git apply。
set -euo pipefail
G="$HOME/flight-repos/gh-private"; P="$HOME/flight-repos/pub"
cmd="${1:-}"; D="${2:-}"
[ -n "$cmd" ] && [ -n "$D" ] || { sed -n '2,9p' "$0"; exit 1; }
case "$cmd" in
  new)
    [ -e "$D" ] && { echo "已存在：$D"; exit 1; }
    mkdir -p "$D/home/Library" "$D/tmp"
    git -C "$G" worktree add -q --detach "$D/gh-private" HEAD
    git -C "$P" worktree add -q --detach "$D/pub" HEAD
    ln -s "$G/node_modules" "$D/gh-private/node_modules"          # 依赖不在 git 里
    ln -s "$HOME/Library/Fonts" "$D/home/Library/Fonts"           # LibreOffice 字体（SD-107）
    [ -f "$HOME/.gitconfig" ] && cp "$HOME/.gitconfig" "$D/home/.gitconfig"
    # 引用了却没入库的图（K-S10 正常应为 0）：链过去免得构建失败
    (cd "$G" && python3 - "$D/gh-private" <<'PY'
import os, re, glob, subprocess, sys
dst = sys.argv[1]
trk = set(subprocess.run(['git', 'ls-files', '-z'], capture_output=True, text=True).stdout.split('\0'))
for f in glob.glob('notes_src/*/*.md') + ['速查/速查源.md']:
    for m in re.findall(r'figures/[^\s|)\]"\'<>]+\.(?:png|jpe?g|svg)', open(f, encoding='utf-8').read()):
        p = 'notes_src/' + m
        if p not in trk and os.path.exists(p) and not os.path.exists(os.path.join(dst, p)):
            os.makedirs(os.path.dirname(os.path.join(dst, p)), exist_ok=True); os.symlink(os.path.abspath(p), os.path.join(dst, p))
            print('  链接未入库的图：' + p)
PY
    )
    echo "沙盒已建：$D（gh-private $(git -C "$D/gh-private" rev-parse --short HEAD)，pub $(git -C "$D/pub" rev-parse --short HEAD)）"
    ;;
  run)
    shift 2
    cd "$D/gh-private"
    PYUB="$(python3 -m site --user-base)"
    HOME="$D/home" TMPDIR="$D/tmp" PYTHONUSERBASE="$PYUB" "$@"
    ;;
  drop)
    git -C "$G" worktree remove --force "$D/gh-private" 2>/dev/null || true
    git -C "$P" worktree remove --force "$D/pub" 2>/dev/null || true
    T="$HOME/flight-repos/_to_delete/sandbox_$(date +%Y%m%d_%H%M%S)"; mkdir -p "$(dirname "$T")"; mv "$D" "$T"
    echo "沙盒已拆，目录移到 $T"
    ;;
  *) sed -n '2,9p' "$0"; exit 1 ;;
esac
