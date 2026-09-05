#!/usr/bin/env bash
# sync.sh —— 笔记库一键同步：校验 → 重建 docx → 提交 → 推送
#
# 在 vault（＝ git 仓库）里任意位置执行：
#     ./sync.sh "改了什么"          只重建有改动的模块
#     ./sync.sh --full "改了什么"   重建全书（页数随字体环境变化，PingFang SC 下约 370 页）
#     ./sync.sh --check             只跑校验，不构建不提交
#     ./sync.sh --no-push "..."     构建并提交，但不推送
#     ./sync.sh --no-build "..."    只提交推送，不重建 docx
#
# 任一校验失败即中止，不会把有问题的内容推上去。
# 工具链默认取本脚本所在目录，可用环境变量 TOOLKIT 覆盖。

set -uo pipefail

RED=$'\033[31m'; GRN=$'\033[32m'; YEL=$'\033[33m'; DIM=$'\033[2m'; RST=$'\033[0m'
die()  { printf '%s✗ %s%s\n' "$RED" "$*" "$RST" >&2; exit 1; }
ok()   { printf '%s✓%s %s\n' "$GRN" "$RST" "$*"; }
info() { printf '%s·%s %s\n' "$DIM" "$RST" "$*"; }
warn() { printf '%s!%s %s\n' "$YEL" "$RST" "$*"; }

# ---------- 参数 ----------
FULL=0; PUSH=1; CHECK_ONLY=0; BUILD=1; MSG=""
BUILT=""                        # 换行分隔的成品清单（bash 3.2 下比数组稳）
while [ $# -gt 0 ]; do
  case "$1" in
    --full)     FULL=1 ;;
    --check)    CHECK_ONLY=1 ;;
    --no-push)  PUSH=0 ;;
    --no-build) BUILD=0 ;;
    -m)         shift; MSG="${1:-}" ;;
    -h|--help)  sed -n '2,16p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    -*)         die "未知参数：$1" ;;
    *)          MSG="$1" ;;
  esac
  shift
done

# ---------- 定位 ----------
SELF_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(git rev-parse --show-toplevel 2>/dev/null)" || die "当前目录不在任何 git 仓库里"
cd "$ROOT" || die "无法进入 $ROOT"
[ -d notes_src ] || die "仓库根目录没有 notes_src/，这不是笔记库"

TOOLKIT="${TOOLKIT:-$SELF_DIR}"
[ -f "$TOOLKIT/assemble.py" ] || TOOLKIT="$ROOT/../pub/tools"
[ -f "$TOOLKIT/assemble.py" ] || die "找不到工具链（assemble.py），可设 TOOLKIT 环境变量指定"
TOOLKIT="$(cd "$TOOLKIT" && pwd)"
info "仓库 $ROOT"
info "工具链 $TOOLKIT"

mod_name() {
  case "$1" in
    0) echo "基础知识速查区" ;;   1) echo "第一章_系统理论" ;;
    2) echo "第二章_机组训练手册" ;; 3) echo "第三章_运行规范" ;;
    4) echo "第四章_模拟机训练" ;;  5) echo "第五章_技术提示" ;;
    *) echo "mod$1" ;;
  esac
}

# ---------- 1. 源码级校验 ----------
echo; info "[1/5] 源码级校验"
python3 "$TOOLKIT/check_src.py" --src notes_src || die "源码校验未通过，已中止（什么都没提交）"

if [ "$CHECK_ONLY" = 1 ]; then echo; ok "仅校验模式，结束"; exit 0; fi

DIRTY=1
[ -z "$(git status --porcelain)" ] && DIRTY=0
if [ "$DIRTY" = 0 ] && [ "$FULL" = 0 ]; then
  echo; warn "工作区没有任何改动"
  if [ "$PUSH" = 1 ]; then
    info "仍尝试推送本地已有提交"
    git push -q origin HEAD && ok "推送完成" || warn "无可推送内容"
  fi
  exit 0
fi
[ "$DIRTY" = 0 ] && info "工作区无改动，但 --full 要求重建"

# ---------- 2. 找出改了哪些模块 ----------
# core.quotepath=false：否则中文路径会被转义成 \346\250\241，匹配不到模块号
CHANGED="$(
  { git -c core.quotepath=false diff --name-only HEAD -- notes_src 2>/dev/null
    git -c core.quotepath=false ls-files --others --exclude-standard -- notes_src 2>/dev/null
  } | sed -n 's#^notes_src/\([0-9]\)[^/]*/.*#\1#p' | sort -u | tr '\n' ' '
)"
CHANGED="$(echo "$CHANGED" | sed 's/ *$//')"

echo
if [ "$FULL" = 1 ]; then
  info "[2/5] 全书重建模式"
elif [ -z "$CHANGED" ]; then
  info "[2/5] 改动不涉及 notes_src 正文，跳过构建"
  BUILD=0
else
  info "[2/5] 改动模块：$CHANGED"
fi

# ---------- 3-4. 组装 + 渲染 + 排版校验 ----------
if [ "$BUILD" = 1 ]; then
  echo; info "[3/5] 组装"
  python3 "$TOOLKIT/assemble.py" --src notes_src --out build || die "组装失败"

  NODE_PATH="$ROOT/node_modules"; export NODE_PATH
  [ -d "$NODE_PATH/docx" ] || die "缺 docx 模块，先在仓库根目录跑：npm install docx"
  # 封面署名与 git 提交身份是两回事：git 身份可能用账号名/化名，
  # 而成品封面要署真名。优先取仓库本地配置 notes.docAuthor：
  #     git config notes.docAuthor "by　张三"
  # 没设才退回 user.name。
  DOC_AUTHOR="${DOC_AUTHOR:-$(git config --get notes.docAuthor || git config user.name)}"
  export DOC_AUTHOR

  echo; info "[4/5] 渲染 docx"
  if [ "$FULL" = 1 ]; then
    out="build/737理论知识笔记_全书.docx"
    node "$TOOLKIT/build_docx.js" build/full.md "$out" || die "全书渲染失败"
    BUILT="$out"
  else
    for n in $CHANGED; do
      if [ ! -f "build/mod$n.md" ]; then warn "build/mod$n.md 不存在，跳过"; continue; fi
      out="build/$(mod_name "$n").docx"
      node "$TOOLKIT/build_docx.js" "build/mod$n.md" "$out" || die "模块 $n 渲染失败"
      BUILT="$BUILT${BUILT:+$'\n'}$out"
    done
  fi

  SOF="$(command -v soffice 2>/dev/null || true)"
  [ -n "$SOF" ] || SOF="/Applications/LibreOffice.app/Contents/MacOS/soffice"
  if [ -x "$SOF" ]; then
    echo "$BUILT" | while IFS= read -r f; do
      [ -n "$f" ] || continue
      "$SOF" --headless --convert-to pdf --outdir build "$f" >/dev/null 2>&1
      pdf="build/$(basename "${f%.docx}").pdf"
      if [ ! -f "$pdf" ]; then warn "未生成 $pdf，跳过排版校验"; continue; fi
      python3 "$TOOLKIT/verify.py" "$pdf" || exit 1
    done || die "排版校验未通过，已中止（未提交）"
    ok "排版校验通过"
  else
    warn "未装 LibreOffice，跳过排版校验（空白页 / 表格跨页查不了）"
    warn "补装：brew install --cask libreoffice"
  fi
else
  echo; info "[3-4/5] 跳过构建"
fi

# ---------- 5. 提交推送 ----------
echo; info "[5/5] 提交"
if [ "$DIRTY" = 0 ]; then
  warn "工作区无改动，跳过提交（成品已重建）"
  if [ -n "$BUILT" ]; then
    echo; ok "Word 成品："
    echo "$BUILT" | while IFS= read -r f; do [ -n "$f" ] && printf '    %s/%s\n' "$ROOT" "$f"; done
  fi
  exit 0
fi
if [ -z "$MSG" ]; then
  n="$(git status --porcelain -- notes_src | wc -l | tr -d ' ')"
  MSG="fix: 更新 $n 个笔记文件"
  warn "未给提交说明，自动用：$MSG"
fi
case "$MSG" in
  add:*|fix:*|refactor:*|docs:*|style:*|rule:*|chore:*) ;;
  *) MSG="fix: $MSG" ;;
esac

git add -A
git commit -q -m "$MSG" || die "提交失败"
ok "已提交 $(git log --oneline -1)"

if [ "$PUSH" = 1 ]; then
  git push -q origin HEAD || die "推送失败（检查网络，或跑 gh auth status）"
  ok "已推送到 $(git remote get-url origin)"
else
  info "--no-push，未推送"
fi

if [ -n "$BUILT" ]; then
  echo; ok "Word 成品："
  echo "$BUILT" | while IFS= read -r f; do [ -n "$f" ] && printf '    %s/%s\n' "$ROOT" "$f"; done
fi
