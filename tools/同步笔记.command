#!/bin/bash
# 同步笔记.command —— 双击运行的同步入口（替代原来的 同步笔记.app）
#
# 为什么不再用 .app：AppleScript 编译之后读不出里面执行了什么命令，
# 想改参数得开脚本编辑器重新编译。.command 就是一个普通 shell 脚本，
# 双击即在终端运行，随时可以用文本编辑器打开看、直接改。
#
# 安装：把本文件（或它的替身）放到桌面 / 访达侧边栏，首次需要
#       chmod +x '同步笔记.command'
#
# 它比 sync.sh 多做的事：
#   · 先告诉你改了哪些文件，再问你要不要构建
#   · 不写死 --full，默认走快的那条路
#   · 顺带推工具链仓库（sync.sh 只管笔记库）
#   · 出错时窗口不关，让你看得到报错

set -uo pipefail
# 不强制设 LC_ALL：终端本来就是 UTF-8，硬设反而会在某些机器上打出 setlocale 警告

B=$'\033[1m'; DIM=$'\033[2m'; GRN=$'\033[32m'; RED=$'\033[31m'; YEL=$'\033[33m'; RST=$'\033[0m'
line() { printf '%s────────────────────────────────────────────────%s\n' "$DIM" "$RST"; }
die()  { printf '\n%s✗ %s%s\n' "$RED" "$*" "$RST"; hold; exit 1; }
hold() { printf '\n%s按回车键关闭…%s' "$DIM" "$RST"; read -r _ || true; }

# ---------- 定位两个仓库 ----------
SELF="$0"; while [ -L "$SELF" ]; do SELF="$(readlink "$SELF")"; done
HERE="$(cd "$(dirname "$SELF")" && pwd)"
for base in "$HERE" "$HERE/.." "$HERE/../.." "$HERE/../../.." "$HOME/flight-repos"; do
  [ -d "$base/gh-private/notes_src" ] && { ROOT="$(cd "$base/gh-private" && pwd)"; break; }
  [ -d "$base/notes_src" ]            && { ROOT="$(cd "$base" && pwd)"; break; }
done
[ -n "${ROOT:-}" ] || die "找不到笔记库（含 notes_src 的那个目录）"
TOOLKIT="$(cd "$ROOT/../pub/tools" 2>/dev/null && pwd)" || true
[ -f "${TOOLKIT:-}/sync.sh" ] || die "找不到工具链 pub/tools/sync.sh"
PUB="$(cd "$TOOLKIT/.." && pwd)"

clear 2>/dev/null || true
printf '%s\n' "$B  737 理论知识笔记 · 同步$RST"
line
printf '  笔记库　%s\n  工具链　%s\n' "$ROOT" "$PUB"
line

# ---------- 先看看有什么要做 ----------
cd "$ROOT" || die "进不去 $ROOT"
CHANGED="$(git status --porcelain -- notes_src 2>/dev/null | wc -l | tr -d ' ')"
AHEAD_N="$(git rev-list --count @{u}..HEAD 2>/dev/null || echo 0)"
AHEAD_T="$(git -C "$PUB" rev-list --count @{u}..HEAD 2>/dev/null || echo 0)"

if [ "$CHANGED" -gt 0 ]; then
  printf '  %s笔记有 %s 处改动：%s\n' "$YEL" "$CHANGED" "$RST"
  git -c core.quotepath=false status --porcelain -- notes_src | head -12 | sed 's/^/      /'
  [ "$CHANGED" -gt 12 ] && printf '      %s…还有 %s 处%s\n' "$DIM" "$((CHANGED-12))" "$RST"
else
  printf '  %s笔记没有改动%s\n' "$DIM" "$RST"
fi
[ "$AHEAD_N" -gt 0 ] && printf '  %s笔记库有 %s 条提交未推送%s\n' "$YEL" "$AHEAD_N" "$RST"
[ "$AHEAD_T" -gt 0 ] && printf '  %s工具链有 %s 条提交未推送%s\n' "$YEL" "$AHEAD_T" "$RST"
command -v soffice >/dev/null 2>&1 || [ -x /Applications/LibreOffice.app/Contents/MacOS/soffice ] \
  || printf '  %s未装 LibreOffice —— 排版校验会被跳过%s\n' "$YEL" "$RST"
line

# ---------- 选做什么 ----------
printf '  %s1%s  快速同步　　重建改动过的分册 → 提交 → 推两个仓库（日常用这个）\n' "$B" "$RST"
printf '  %s2%s  全书重建　　同上，另加 465 页合订本（慢，几分钟）\n' "$B" "$RST"
printf '  %s3%s  只校验　　　不构建、不提交，只看有没有问题\n' "$B" "$RST"
printf '  %s4%s  只推送　　　不构建，把两个仓库已有的提交推上去\n' "$B" "$RST"
printf '  %s0%s  退出\n' "$B" "$RST"
printf '\n选择 [1]: '
read -r PICK || PICK=1
PICK="${PICK:-1}"

case "$PICK" in
  0) printf '\n已取消\n'; hold; exit 0 ;;
  3) printf '\n'; bash "$TOOLKIT/sync.sh" --check; RC=$?
     [ $RC -eq 0 ] && printf '\n%s✓ 校验通过%s\n' "$GRN" "$RST" || printf '\n%s✗ 校验未通过%s\n' "$RED" "$RST"
     ;;
  4) printf '\n'; bash "$TOOLKIT/push_all.sh" "$ROOT"; RC=$? ;;
  1|2)
     if [ "$CHANGED" -eq 0 ] && [ "$PICK" = 1 ]; then
       printf '\n%s没有改动可提交，将只推送已有提交%s\n' "$DIM" "$RST"
     fi
     printf '\n改了什么？（直接回车用「更新笔记」）: '
     read -r MSG || MSG=""
     MSG="${MSG:-更新笔记}"
     printf '\n'
     if [ "$PICK" = 2 ]; then bash "$TOOLKIT/sync.sh" --full --no-push "$MSG"
     else                       bash "$TOOLKIT/sync.sh"        --no-push "$MSG"; fi
     RC=$?
     if [ "$RC" -eq 0 ]; then
       line; printf '  %s推送两个仓库%s\n' "$B" "$RST"
       bash "$TOOLKIT/push_all.sh" "$ROOT" || RC=$?
     fi
     ;;
  *) die "无效选择：$PICK" ;;
esac

line
if [ "${RC:-1}" -eq 0 ]; then printf '%s✓ 完成%s\n' "$GRN" "$RST"
else printf '%s✗ 有步骤失败，上面的输出里有原因%s\n' "$RED" "$RST"; fi
hold
exit "${RC:-1}"
