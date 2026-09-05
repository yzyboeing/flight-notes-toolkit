# 归档：flight-theory-notes skill v1.8

本分支保存 `~/.codex/skills/flight-theory-notes` 在 2026-09-05 被改成薄路由**之前**的原始文件。该 skill 从未有过版本管理，这是它唯一的版本化副本。

## 为什么归档而不是删除

其中的**规则**已于 2026-09-05 逐条并入 `main` 分支（见 `CHANGELOG.md` 的
「合并 Codex 侧并行分支的独有规则」与「并入 Codex 侧 skill 的最后四条独有规则」两节，
共 18 项，每项都经实证核实）。但**文件本身**从未进过任何仓库：

- `scripts/build_docx.js` 与 `main` 分支的 `tools/build_docx.js` 差 993 行，
  是同一功能的**另一次独立实现**，可能藏有更好的做法；
- `references/` 下 7 个文件、`dist/`、`tools/export_prompt.js`、`VERSION`
  在别处不存在同名文件。

## 不要从这里取规则

**规则真源是 `main` 分支的 `prompt/SKILL.md` 与 `prompt/standing-decisions.md`。**
本分支内容已停止维护，其中若干条已被实证否定并记入
`standing-decisions.md` 的「历史决策」（5 章重构方案、条目级 `[NG]` 标记、
块编号并入库后作废、结尾语、`EAEFF3` 共享项色等）。仅作追溯用途。
