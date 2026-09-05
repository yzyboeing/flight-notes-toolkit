# flight-theory-notes

把中文航空原始材料整理成结构化复习材料，产出带目录的 `.docx`；也用于已有笔记库的重组、去重、编号治理和增量维护。

## 一份源，两种形态

```text
SKILL.md + references/     唯一的规则源，只改这里
        │
        ├─ Claude / Codex：整个目录作为 skill 安装，按需加载 references
        └─ ChatGPT / Gemini：node tools/export_prompt.js → dist/ 单文件 prompt
```

不要手工修改 `dist/` 下的 prompt。它由脚本生成，下次导出会被覆盖。

## 目录

```text
SKILL.md
VERSION
references/
  fidelity.md
  decisions.md
  structure.md
  formats.md
  output.md
  source_material.md
  library.md
scripts/build_docx.js
tools/export_prompt.js
tools/iteration_prompt_variant.md
```

## 用法

安装为 Skill：把整个目录复制到 skills 目录。

导出单文件 Prompt：

```bash
node tools/export_prompt.js
```

生成笔记成品：

```bash
node scripts/build_docx.js src/01_systems/01.06_electrical.md build/01.06_electrical.docx
```

首次在 Word 中打开需更新域，目录才会显示条目；LibreOffice 或预览器里目录为空属正常。

## 改规则

1. 修改 `SKILL.md` 或 `references/` 对应文件。
2. 更新 `VERSION`。
3. 运行 `node tools/export_prompt.js`。
4. 校验后再决定是否提交。

## 内容边界

本目录只保存工作方法和脚本，不保存笔记正文、受控手册内容、运行限制值、机号、内部编号、组织身份信息或凭据。笔记正文使用本地 git 管理，不进入公开托管仓库。
