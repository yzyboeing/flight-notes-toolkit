#!/usr/bin/env bash
# lulu_bundle.sh —— 出基线后，把私有库与工具链打包放进噜噜交接包「源码包」（SD-181 方案一，2026-10-14）
# 用法：lulu_bundle.sh <基线版本名，如 1013R4>
#   · 在 打基线标签、推送 之后运行（bundle 取当前 main 与全部标签）
#   · 旧 bundle 移到 ~/flight-repos/_to_delete/源码包旧版_<日期>/（不硬删）
#   · 字体、环境脚本、验收比对脚本不动
set -euo pipefail
ED="${1:?请给出基线版本名，如 1013R4}"
SELF="$(cd "$(dirname "$0")" && pwd)"; REPOS="$(cd "$SELF/../.." && pwd)"
PKG="$HOME/Desktop/飞行理论笔记整理/Muse交接包/源码包"
[ -d "$PKG/" ] || { echo "✗ 找不到 $PKG"; exit 1; }
OLD="$REPOS/_to_delete/源码包旧版_$(date +%Y%m%d)"; mkdir -p "$OLD"
find "$PKG/" -maxdepth 1 -name '*.bundle' -exec mv {} "$OLD/" \;
NB="$PKG/flight-theory-notes_${ED}.bundle"
TB="$PKG/flight-notes-toolkit_$(git -C "$REPOS/pub" rev-parse --short HEAD).bundle"
git -C "$REPOS/gh-private" bundle create "$NB" --all >/dev/null 2>&1
git -C "$REPOS/pub"        bundle create "$TB" --all >/dev/null 2>&1
git -C "$REPOS/gh-private" bundle verify "$NB" >/dev/null 2>&1 && git -C "$REPOS/pub" bundle verify "$TB" >/dev/null 2>&1
echo "✓ $(basename "$NB")（$(du -h "$NB" | cut -f1)）  ✓ $(basename "$TB")（$(du -h "$TB" | cut -f1)）"
echo "  私有库 main：$(git -C "$REPOS/gh-private" log --oneline -1 | cut -c1-60)"
echo "  旧 bundle → $OLD"
