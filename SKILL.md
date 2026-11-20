---
name: flight-theory-notes
description: "维护中文《B737-NG / B737-8 机型理论速查笔记》（一份真源 notes_src，5 章 69 节）及其成品：iPad App（唯一主产物）、按需的 Word / PDF。用于：加知识点、判定章节与位置、做两型对比表、校对数值与两型差异、按手册库补充重点知识点、审核外部 AI（ChatGPT / Codex）的交付、App 门禁与发版、表格排版问题。当用户说'加一个知识点'、'这段该入哪一节'、'做成对比表'、'这个表排版不对'、'检查两型差异'、'校对第 N 章'、'审核 Codex 的交付'、'给 GPT 一份指令'、'发 TestFlight'时使用。不用于英文航空翻译、培训 PPT、通用代码开发；不替代现行有效的公司手册。"
---

# 飞行理论笔记（flight-theory-notes）

你在帮一位现役 B737 机长维护私人学习笔记《B737-NG / B737-8 机型理论速查笔记》，适用 737-NG 与 737-8。唯一内容真源是私有库 `~/flight-repos/gh-private/notes_src/`；iPad App「737 机型笔记」（仓库 `~/flight-repos/study`）是唯一主产物；Word / PDF 链路冻结，用户点名才出；OneNote 已停更（SD-266），禁止 Graph 写入。

## 一、先读什么（2026-11-20 规则瘦身后）

| 顺序 | 文件 | 回答什么 |
|---|---|---|
| 1 | 私有库 `gh-private/规则/现行规则.md` | **唯一现行规则**：红线、工作方式、结构、表格、插图、写法、流程、规则维护（≤ 30 KB） |
| 2 | 工作区 `README.md`、`现状一页纸.md`（自动生成）、`00_先读我_AI交接.md` | 现状、规模、构建号、最近改了什么 |
| 3 | `~/flight-repos/study/更新手册.md`、`DESIGN.md` | 构建、门禁、发版命令；App 界面规范（改界面时才读） |
| 4 | 本仓库 `prompt/checks.md` | 检查器规则号与 W06 清单（工具数据，改检查器时才读） |

`prompt/archive/2026-11-20_规则瘦身前/` 冻结了 SD-1～266 长期决策、旧规则总纲、排版总表、内容核心、经验与教训和旧单文件版，**只查来历，不加载**。现行规则含章节标题与运行限制值，按本仓库红线放在私有库，不进公开库。

**优先级**：用户在对话里的最新原话 ＞ 私有库《现行规则》＞ 其他所有文件。

## 二、本仓库是什么

公开的方法论与工具链：`tools/` 下的源检查器（`check_src.py` 等）、规则自检（`check_rules.py`）、出图与门禁辅助、`push_all.sh`（过包检查与私有词泄漏扫描）。笔记正文、私有库章节标题、具体运行限制值、公司名不进本仓库。

## 三、红线（摘要，全文见现行规则第〇节）

1. 数值逐字；两型不默认相同，依据只来自本机手册库，不上网、不凭记忆；手册查不到的是培训资料，不删不标。
2. 不硬删文件（移到 `_to_delete/`）；只按路径提交；不 `commit -a` / stash / reset / 强推；不读不印密钥。
3. git 推送、永久删除或覆盖用户原始资料、写 MEM、TestFlight 上传与加组、网页上线，须用户本人逐批批准。
4. 没跑就说没跑；全程中文。

## 四、维护本技能

- 规则只改私有库 `规则/现行规则.md`；新的用户决定写进同目录 `决策记录.md`（SD-267 起，每条带「落实：」行）。
- 改检查器时在 `prompt/checks.md` 登记 K- 号。
- 改完依次跑：`python3 tools/export_prompt.py`、`python3 tools/check_rules.py`、`python3 tools/check_package.py --private-keywords ~/.leakscan-keywords`、`python3 -m unittest discover -s tests`。
- 废掉一个口径时，把它的说法登记进 `check_rules.py` 的 STALE 表。公开推送用 `push_all.sh`。
