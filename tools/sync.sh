#!/usr/bin/env bash
# sync.sh —— 笔记库一键同步：校验 → 重建 docx → 本地提交（默认不推送）
#
# 在 vault（＝ git 仓库）里任意位置执行：
#     ./sync.sh "改了什么"          只重建有改动的模块
#     ./sync.sh --full "改了什么"   重建全书（页数随字体环境变化，PingFang SC 下约 370 页）
#     ./sync.sh --check             只跑校验，不构建不提交
#     ./sync.sh --push "..."        构建、提交并推送（唯一会推送的方式；默认只本地提交）
#     ./sync.sh --no-push "..."     与默认相同，保留以兼容旧用法
#     ./sync.sh --no-build "..."    只提交推送，不重建 docx
#
# 任一校验失败即中止。SD-50：脚本默认绝不推送，只有显式 --push 才推。
# 工具链默认取本脚本所在目录，可用环境变量 TOOLKIT 覆盖。

set -uo pipefail

RED=$'\033[31m'; GRN=$'\033[32m'; YEL=$'\033[33m'; DIM=$'\033[2m'; RST=$'\033[0m'
die()  { printf '%s✗ %s%s\n' "$RED" "$*" "$RST" >&2; exit 1; }
ok()   { printf '%s✓%s %s\n' "$GRN" "$RST" "$*"; }
info() { printf '%s·%s %s\n' "$DIM" "$RST" "$*"; }
warn() { printf '%s!%s %s\n' "$YEL" "$RST" "$*"; }

# ---------- 参数 ----------
FULL=0; PUSH=0; EXPLICIT_PUSH=0; CHECK_ONLY=0; BUILD=1; MSG=""
BUILT=""                        # 换行分隔的成品清单（bash 3.2 下比数组稳）
while [ $# -gt 0 ]; do
  case "$1" in
    --full)     FULL=1 ;;
    --check)    CHECK_ONLY=1 ;;
    --no-push)  PUSH=0; EXPLICIT_PUSH=0 ;;
    --push)     PUSH=1; EXPLICIT_PUSH=1 ;;
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

# 推送只由显式 --push 打开（SD-50 脚本级防呆；notes.noPush 配置不再需要，保留兼容）
[ "$EXPLICIT_PUSH" -eq 1 ] || PUSH=0

TOOLKIT="${TOOLKIT:-$SELF_DIR}"
[ -f "$TOOLKIT/assemble.py" ] || TOOLKIT="$ROOT/../pub/tools"
[ -f "$TOOLKIT/assemble.py" ] || die "找不到工具链（assemble.py），可设 TOOLKIT 环境变量指定"
TOOLKIT="$(cd "$TOOLKIT" && pwd)"
info "仓库 $ROOT"
info "工具链 $TOOLKIT"

mod_name() {
  case "$1" in
    0) echo "基础知识速查区" ;;   1) echo "第一章_系统理论" ;;
    2) echo "第二章_机组训练手册" ;; 3) echo "第三章_运行手册" ;;
    4) echo "第四章_模拟机训练" ;;  5) echo "第五章_技术提示" ;;
    *) echo "mod$1" ;;
  esac
}

# ---------- 1. 源码级校验 ----------
echo; info "[1/5] 源码级校验"
python3 "$TOOLKIT/check_src.py" --src notes_src || die "源码校验未通过，已中止（什么都没提交）"
if [ -f "$TOOLKIT/check_blocks.py" ]; then
  python3 "$TOOLKIT/check_blocks.py" --src notes_src || die "块形态校验未通过，已中止（什么都没提交）"
fi

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

# ---------- 1b. 速查区改动的连带检查（SD-92，2026-09-30 用户：速查区改了，后续章节对应内容要一并改） ----------
# SD-139（2026-10-05 用户）：第零章速查区已删除，速查连带检查停用
# if [ -f "$TOOLKIT/quickref_sync_hint.py" ]; then python3 "$TOOLKIT/quickref_sync_hint.py" || true; fi

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
  # 封面版次与声明（可选）：git config notes.docEdition "20260930"；git config notes.docNotice "…"
  # SD-142（2026-10-05 用户）：封面版本号与文件名同为版本名「月日 + R + 当天第几版」（如 1005R7）——
  # notes_src 相对最新 baseline/* 标签有改动（含未提交）时，印下一版的版本名（出基线时按它打标签）；没有改动时印最新基线的版本名
  if [ "$FULL" = 1 ] && [ -z "${DOC_EDITION:-}" ]; then
    LASTBL="$(git tag --list 'baseline/*' --sort=-creatordate | head -1)"
    if [ -z "$LASTBL" ] || ! git diff --quiet "$LASTBL" -- notes_src 2>/dev/null; then
      VER="$(python3 "$TOOLKIT/release_name.py" --next | awk '{print $1}')"
    else
      VER="$(python3 "$TOOLKIT/release_name.py" "$LASTBL")"
    fi
    if [ -n "$VER" ] && [ "$(git config --get notes.docEdition)" != "$VER" ]; then
      git config notes.docEdition "$VER"; ok "版本号更新为 $VER（SD-142）"
    fi
  fi
  DOC_EDITION="${DOC_EDITION:-$(git config --get notes.docEdition)}"; export DOC_EDITION
  # SD-118 速查「按主题查」索引（主题表按条目标题引用）
  DOC_QRTOPICS="${DOC_QRTOPICS:-$ROOT/速查主题索引.json}"; export DOC_QRTOPICS
  # SD-139 目录后「按主题查」索引（主题 → 正文块「节号 块号」）
  DOC_TOPICS="${DOC_TOPICS:-$ROOT/按主题查索引.json}"; export DOC_TOPICS
  DOC_NOTICE="${DOC_NOTICE:-$(git config --get notes.docNotice)}"; export DOC_NOTICE
  DOC_PREFACE="${DOC_PREFACE:-$ROOT/前言.md}"; export DOC_PREFACE   # SD-51 前言页（文件不存在则不出前言）
  DOC_PREFACE_SIGNATURE="${DOC_PREFACE_SIGNATURE:-$(git config --get notes.docPrefaceSignature)}"; export DOC_PREFACE_SIGNATURE
  # SD-67（2026-09-29 用户）：主题线索引页撤出成品。源文件 主题线索引.md 保留在库里，
  # 需要时把下一行取消注释即可恢复。
  # DOC_TOPICINDEX="${DOC_TOPICINDEX:-$ROOT/主题线索引.md}"; export DOC_TOPICINDEX
  # 印刷方式（可选）：git config notes.docBW 1（黑白：限制值改黑色加粗 + 下划线）；git config notes.docDuplex 1（双面：镜像页边距、奇偶页页眉页脚、每章从右页开始）
  DOC_BW="${DOC_BW:-$(git config --get notes.docBW)}"; export DOC_BW
  DOC_DUPLEX="${DOC_DUPLEX:-$(git config --get notes.docDuplex)}"; export DOC_DUPLEX

  echo; info "[4/5] 渲染 docx"
  if [ "$FULL" = 1 ]; then
    # 2026-09-30 用户：成品文件名「B737机型理论知识笔记」
    out="build/B737机型理论知识笔记.docx"
    DOC_HEADER="${DOC_HEADER:-B737-NG / B737-8 机型理论知识笔记}"; export DOC_HEADER
    python3 "$TOOLKIT/make_book.py" || die "拼合订本失败"
    # SD-79（2026-09-30 用户）：一页放得下的表绝不拆开——fit_fix.py 两遍排版（出 PDF 实测断表 → 强制整表同页 → 重出）
    python3 "$TOOLKIT/fit_fix.py" build/book.md "$out" || die "全书渲染失败（横版）"
    # SD-71（2026-09-29 用户）：只出横版。竖版的表格在窄版面里文字堆叠严重、不利阅读，不再产出、也不再为竖版优化排版。
    # 需要临时出竖版时手动运行：DOC_PORTRAIT=1 node "$TOOLKIT/build_docx.js" build/book.md build/737 机型理论知识笔记_竖版.docx
    # SD-139（2026-10-05 用户）：「以后的速查笔记也就不用再维护了，以后只维护这一本笔记」——速查单册停出（原生成步骤见 git 历史）
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
      pdf="build/$(basename "${f%.docx}").pdf"
      # 转换一律输出到干净的临时目录再拷回来。两个原因：
      #   · 直接往已存在的 PDF 上覆写，在某些文件系统（网络盘 / 受限挂载）会以
      #     Io/Abort 失败；先写临时目录再 cp 覆盖最稳。
      #   · 临时目录里有没有产物，是「这次到底转成功没有」的唯一可信判据，
      #     不会被上一轮留下的旧 PDF 冒充。
      # -env:UserInstallation 给 headless 一个独立 profile：LibreOffice 图形界面
      # 开着时默认 profile 被占用，转换会失败且不报错；独立 profile 可以并存。
      LOPROF="${TMPDIR:-/tmp}/lo-sync-profile"
      LOTMP="$(mktemp -d)"; LOLOG="$(mktemp)"
      # SD-107：LibreOffice 不读 ~/Library/Fonts，把思源字体复制进独立 profile 的 user/fonts
      mkdir -p "$LOPROF/user/fonts"
      for ff in "$HOME"/Library/Fonts/SourceHan*.otf; do [ -f "$ff" ] && { cmp -s "$ff" "$LOPROF/user/fonts/$(basename "$ff")" || cp "$ff" "$LOPROF/user/fonts/"; }; done
      "$SOF" -env:UserInstallation="file://$LOPROF" \
             --headless --convert-to pdf --outdir "$LOTMP" "$f" >"$LOLOG" 2>&1
      fresh="$LOTMP/$(basename "${f%.docx}").pdf"
      if [ ! -f "$fresh" ]; then
        printf '%s✗ %s%s\n' "$RED" "LibreOffice 没能生成 $(basename "$pdf") —— 排版校验未做，不算通过" "$RST" >&2
        printf '%s  LibreOffice 输出：%s\n' "$DIM" "$RST" >&2
        tail -12 "$LOLOG" | sed 's/^/    /' >&2
        printf '%s  排查：确认已装（brew install --cask libreoffice）；仍失败可删掉 %s 重试%s\n' \
               "$DIM" "$LOPROF" "$RST" >&2
        rm -rf "$LOTMP" "$LOLOG"; exit 1
      fi
      cp "$fresh" "$pdf" || { printf '%s✗ 写不进 %s%s\n' "$RED" "$pdf" "$RST" >&2; rm -rf "$LOTMP" "$LOLOG"; exit 1; }
      # 打开 PDF 时直接展开书签栏（PageMode=UseOutlines）；没有 PyMuPDF 时跳过
      python3 - "$pdf" <<'PYEOF' 2>/dev/null || true
import sys, os
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf
import re
p = sys.argv[1]; d = pymupdf.open(p); d.set_pagemode('UseOutlines')
# 2026-10-01 用户：「去掉所有目录中的单位符号例如 m/kt」——书签里去掉标题末尾的「（单位：…）」括注（可嵌套一层括号），正文标题保留
toc = d.get_toc(simple=False); ch = 0
for e in toc:
    t = re.sub(r'\s*（单位：(?:[^（）]|（[^（）]*）)*）', '', e[1])
    if t != e[1]: e[1] = t; ch += 1
if ch: d.set_toc(toc)
d.save(p + '.tmp', garbage=3, deflate=True); d.close(); os.replace(p + '.tmp', p)
PYEOF
      rm -rf "$LOTMP" "$LOLOG"
      python3 "$TOOLKIT/verify.py" "$pdf" || exit 1
    done || die "排版校验未通过，已中止（未提交）"
    ok "排版校验通过"
    # SD-79 总检查器：规则见 pub/prompt/layout-checklist.md。报告写 build/排版检查报告.md；
    # 有「错误」只警告不中止（成品仍要交给用户看），但收工前必须改到 0。
    if [ "$FULL" = 1 ] && [ -f "$TOOLKIT/check_layout.py" ]; then
      echo; info "[4b] 排版与规则总检查（check_layout.py）"
      if python3 "$TOOLKIT/check_layout.py" --repo "$ROOT" --no-src --out build/排版检查报告.md >/dev/null 2>&1; then
        ok "总检查：$(sed -n '3p' build/排版检查报告.md)"
      else
        warn "总检查：$(sed -n '3p' build/排版检查报告.md) —— 错误必须改到 0，详见 build/排版检查报告.md"
      fi
    fi
    # SD-114 联动检查：改动有没有按逻辑同步到全笔记（引用完整性是错误；旧写法残留 / 同值异色 / 联动面是提醒）
    if [ "$FULL" = 1 ] && [ -f "$TOOLKIT/check_linkage.py" ]; then
      echo; info "[4c] 联动检查（check_linkage.py，对比上一个 baseline/* 标签）"
      if LK="$(python3 "$TOOLKIT/check_linkage.py" --repo "$ROOT" --out build/联动检查报告.md 2>&1)"; then
        ok "$LK"
      else
        warn "$LK —— 断链必须改到 0，详见 build/联动检查报告.md"
      fi
    fi
    # SD-111：每次全量重建同时更新《笔记结构索引》（给噜噜判断知识点放哪里），随新基线一起交给噜噜
    if [ "$FULL" = 1 ] && [ -f "$TOOLKIT/notes_index.py" ]; then
      python3 "$TOOLKIT/notes_index.py" notes_src "build/B737机型理论知识笔记.pdf" build/笔记结构索引.md "$(git -C "$ROOT" rev-parse --short HEAD 2>/dev/null)" >/dev/null 2>&1 \
        && ok "笔记结构索引已更新：build/笔记结构索引.md" || warn "笔记结构索引生成失败（不影响成品）"
    fi
    # 私有库自带全部规则与交接资料（2026-10-02 用户：库交给任何新 AI 都能接手）：每次全量重建同步一次副本
    if [ "$FULL" = 1 ]; then
      mkdir -p "$ROOT/规则" "$ROOT/交接" "$ROOT/协作/噜噜"
      cp "$TOOLKIT/../prompt/portable-instruction.md" "$ROOT/规则/规则总纲.md" 2>/dev/null
      cp "$TOOLKIT/../prompt/layout-checklist.md"     "$ROOT/规则/排版检查清单.md" 2>/dev/null
      cp "$TOOLKIT/../prompt/content-core.md"         "$ROOT/规则/内容校对核心规则.md" 2>/dev/null
      # 给噜噜的规则与台账（2026-10-05 单源）：正本在 pub/prompt（排版规则总表、内容校对核心规则）与私有库根目录（内容裁定台账.json），
      # 先拷进桌面交接包；下面的噜噜镜像再从交接包整体拷回（含 *.json），不要再直接写 协作/噜噜/（那里每次会被清空重建）
      HB="$HOME/Desktop/飞行理论笔记整理/Muse交接包"
      if [ -d "$HB" ]; then
        cp "$TOOLKIT/../prompt/layout-checklist.md" "$HB/排版规则总表.md" 2>/dev/null
        cp "$TOOLKIT/../prompt/content-core.md"     "$HB/内容校对核心规则.md" 2>/dev/null
        [ -f "$ROOT/内容裁定台账.json" ] && cp "$ROOT/内容裁定台账.json" "$HB/" 2>/dev/null
      fi
      cp "$TOOLKIT/../prompt/standing-decisions.md"   "$ROOT/规则/长期决策.md" 2>/dev/null
      [ -f build/笔记结构索引.md ] && cp build/笔记结构索引.md "$ROOT/笔记结构索引.md"
      HAND="$HOME/Desktop/飞行理论笔记整理"
      if [ -d "$HAND" ]; then
        cp "$HAND/00_先读我_AI交接.md" "$ROOT/交接/" 2>/dev/null
        cp "$HAND"/AI交接/0*.md "$ROOT/交接/" 2>/dev/null
        rm -rf "$ROOT/协作/噜噜"; mkdir -p "$ROOT/协作/噜噜"
        cp "$HAND"/Muse交接包/*.md "$HAND"/Muse交接包/*.json "$HAND"/噜噜_*.md "$ROOT/协作/噜噜/" 2>/dev/null
        [ -f "$HAND/噜噜收件箱/README.md" ] && cp "$HAND/噜噜收件箱/README.md" "$ROOT/协作/噜噜/噜噜收件箱_README.md"
      fi
      ok "规则 / 结构索引 / 交接快照 / 噜噜指令 已同步进私有库"
    fi
  else
    warn "未装 LibreOffice，跳过排版校验（空白页 / 表格跨页查不了）"
    warn "补装：brew install --cask libreoffice"
  fi
else
  echo; info "[3-4/5] 跳过构建"
fi

# ---------- 5. 提交推送 ----------
echo; info "[5/5] 提交"
[ -n "$(git status --porcelain)" ] && DIRTY=1   # 全量重建同步的规则副本 / 索引也要提交
if [ "$DIRTY" = 0 ]; then
  warn "工作区无改动，跳过提交（成品已重建）"
  # 没东西可提交，不代表没东西可推——之前的提交可能还堆在本地
  AHEAD="$(git rev-list --count @{u}..HEAD 2>/dev/null || echo 0)"
  if [ "$PUSH" = 1 ] && [ "${AHEAD:-0}" -gt 0 ]; then
    info "本地还有 $AHEAD 条未推送的提交，推送中"
    git push -q origin HEAD || die "推送失败（检查网络，或跑 gh auth status）"
    ok "已推送到 $(git remote get-url origin)"
  elif [ "${AHEAD:-0}" -gt 0 ]; then
    info "--no-push，$AHEAD 条提交仍未推送"
  fi
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
