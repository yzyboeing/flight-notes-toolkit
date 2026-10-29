#!/usr/bin/env bash
# qr_build.sh —— 只构建速查（SD-227，2026-10-22 用户：「内容修改两本源文件同改，日常只构建速查，速查稳定后再全量重建完整版」）
# 用法（在 gh-private 里）：DOC_EDITION=<版本名> ../pub/tools/qr_build.sh "改了什么"
# 步骤：源码校验 → 留存上一版 → 速查排版（fit_fix 收尾含定稿后收窄 / 加宽）→ 速查检查（报告只留「单册」）→ 变化页对比 → 完整版落后轮数 +1 → 本地提交
# 两本同表比对、完整版检查、整本过目留到全量（sync.sh --full）；完整版落后满 5 轮在这里和现状页提醒全量收口。
set -uo pipefail
TOOLKIT="$(cd "$(dirname "$0")" && pwd)"; ROOT="$(pwd)"
MSG="${1:-速查构建}"
export BUILD_QR=1   # SD-242：本脚本只出旧速查单册（溯源对照），检查链要带上速查源
c() { printf '\033[%sm%s\033[0m\n' "$1" "$2"; }
[ -d notes_src ] && [ -f 速查/速查源.md ] || { c 31 "请在 gh-private 里运行"; exit 1; }
pgrep -f 'sync.sh|qr_build.sh' | grep -v $$ >/dev/null && { c 31 "已有构建在跑（sync / qr_build），先等它结束"; exit 1; }

c 2 "· [1/6] 源码校验"
python3 "$TOOLKIT/check_src.py" --src notes_src > build/源码校验_速查.txt 2>&1
if grep -q '错误' build/源码校验_速查.txt; then grep '错误' build/源码校验_速查.txt | head -20; c 31 "✗ 源码校验未通过，已中止（什么都没提交）"; exit 1; fi

c 2 "· [2/6] 留存上一版速查（变化页对比用）"
mkdir -p build/_上一版; [ -f build/B737机型理论速查.pdf ] && cp -p build/B737机型理论速查.pdf build/_上一版/

c 2 "· [3/6] 速查排版（fit_fix；日志 build/quickref_fitfix.log）"
python3 "$TOOLKIT/quickref.py" --repo "$ROOT" || { c 31 "✗ 速查排版失败"; exit 1; }

c 2 "· [4/6] 速查检查（check_layout，只看单册）"
python3 "$TOOLKIT/check_layout.py" >/dev/null 2>&1
python3 - <<'PY'
import re
t = open('build/排版检查报告.md', encoding='utf-8').read()
keep = [l for l in t.splitlines() if not l.startswith('- [') or '单册' in l or '速查' in l]
open('build/排版检查报告_速查.md', 'w', encoding='utf-8').write('\n'.join(keep) + '\n')
e = t.split('## 错误')[1].split('\n## ')[0] if '## 错误' in t else ''
n = sum(1 for l in e.splitlines() if l.startswith('- [') and ('单册' in l or '速查' in l))
print('  速查错误 %d 条（必须改到 0）→ build/排版检查报告_速查.md' % n)
PY

c 2 "· [5/6] 变化页对比"
[ -f build/_上一版/B737机型理论速查.pdf ] && python3 "$TOOLKIT/page_diff.py" build/_上一版/B737机型理论速查.pdf build/B737机型理论速查.pdf build/变化页对比 --tag 速查 | sed 's/^/  /'

c 2 "· [6/6] 完整版落后轮数与提交"
N=$(cat build/完整版落后轮数.txt 2>/dev/null || echo 0); N=$((N + 1)); echo "$N" > build/完整版落后轮数.txt
if [ "$N" -ge 5 ]; then c 33 "! 完整版已落后 $N 轮速查构建——提醒用户全量收口（sync.sh --full）"; else c 2 "  完整版落后 $N 轮（满 5 轮提醒全量）"; fi
git add -A notes_src 速查 && git add -f build/keep_force_B737机型理论速查.txt build/完整版落后轮数.txt 2>/dev/null
git commit -qm "速查构建：$MSG（完整版落后 $N 轮）" && c 32 "✓ 已提交 $(git rev-parse --short HEAD)"
