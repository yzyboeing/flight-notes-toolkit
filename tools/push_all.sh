#!/usr/bin/env bash
# push_all.sh —— 一次把两个仓库都推上去
#
#     私有  flight-theory-notes    笔记正文，直接推
#     公开  flight-notes-toolkit   工具链，**先过泄漏扫描才推**
#
# 用法（在笔记库里任意位置，或显式给路径）：
#     push_all.sh [笔记库路径]
#
# 公开仓库的规则（rules.md「内容不外发」）：
#   · 推之前跑一次泄漏扫描：关键词表 + 本机绝对路径
#   · 关键词表**不存在时停下报告，不得当作通过**——它就是用来防这件事的
#   · 有命中就停下，不推
set -uo pipefail
RED=$'\033[31m'; GRN=$'\033[32m'; YEL=$'\033[33m'; DIM=$'\033[2m'; RST=$'\033[0m'
ok(){ printf '%s✓%s %s\n' "$GRN" "$RST" "$*"; }
no(){ printf '%s✗%s %s\n' "$RED" "$RST" "$*"; }
info(){ printf '%s·%s %s\n' "$DIM" "$RST" "$*"; }
warn(){ printf '%s!%s %s\n' "$YEL" "$RST" "$*"; }

SELF_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="${1:-$(git rev-parse --show-toplevel 2>/dev/null)}"
[ -d "${ROOT:-}/notes_src" ] || { no "找不到笔记库（含 notes_src）"; exit 1; }
ROOT="$(cd "$ROOT" && pwd)"
PUB="$(cd "$SELF_DIR/.." && pwd)"
RC=0

push_one() {   # $1=仓库路径  $2=名字
  local r="$1" n="$2" ahead
  ahead="$(git -C "$r" rev-list --count @{u}..HEAD 2>/dev/null || echo 0)"
  if [ "$ahead" -eq 0 ]; then info "$n 无待推送提交"; return 0; fi
  if git -C "$r" push -q origin HEAD; then
    ok "$n 已推送 $ahead 条 → $(git -C "$r" remote get-url origin)"
  else
    no "$n 推送失败（检查网络，或跑 gh auth status）"; RC=1
  fi
}

# ---------- 1. 私有：笔记库 ----------
push_one "$ROOT" "笔记库（私有）"

# ---------- 2. 公开：工具链，先扫描 ----------
AHEAD_T="$(git -C "$PUB" rev-list --count @{u}..HEAD 2>/dev/null || echo 0)"
if [ "$AHEAD_T" -eq 0 ]; then
  info "工具链 无待推送提交"
else
  KW=""
  for c in "$HOME/.leakscan-keywords" "$ROOT/.leakscan-keywords"; do
    [ -s "$c" ] && { KW="$c"; break; }
  done
  if [ -z "$KW" ]; then
    no "找不到泄漏扫描关键词表（~/.leakscan-keywords 或 $ROOT/.leakscan-keywords）"
    warn "按规则：表不存在时停下报告，不得当作通过。工具链**未推送**。"
    exit 1
  fi
  info "泄漏扫描（关键词表 $KW，$(wc -l < "$KW" | tr -d ' ') 条）"
  HITS="$(cd "$PUB" && grep -rnIf "$KW" --exclude-dir=.git . 2>/dev/null | grep -v '^\./CHANGELOG\.md' || true)"
  # 模式拼出来而不是写死，否则本脚本自己会被自己扫中
  UPAT="/$(printf 'Users')/|/$(printf 'home')/[a-z]|/$(printf 'sessions')/"
  PATHS="$(cd "$PUB" && grep -rnIE "$UPAT" --exclude-dir=.git . 2>/dev/null || true)"
  if [ -n "$HITS" ] || [ -n "$PATHS" ]; then
    no "扫描有命中，工具链**未推送**（只列文件:行，不显示内容）"
    { [ -n "$HITS" ] && printf '%s\n' "$HITS"; [ -n "$PATHS" ] && printf '%s\n' "$PATHS"; } \
      | cut -d: -f1-2 | sort -u | sed 's/^/    /'
    exit 1
  fi
  ok "泄漏扫描无命中"
  push_one "$PUB" "工具链（公开）"
fi
exit $RC
