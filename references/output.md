# Word 成品与 Markdown 源文件

## Word 成品规格

- 成品是可编辑、带自动目录的 `.docx`，全篇 A4 横版、窄边距。
- 首页使用 Word TOC field，收录一级和二级标题并可点击跳转。首次用 Word 打开可能需要更新域；LibreOffice 或预览器中目录为空不能单独证明目录损坏。
- 标题使用 Word 内置 Heading 1 / 2 / 3。
- 模式 B：Heading 1 为 `0 速查区`、`1 系统理论`、`2 运行规范`、`3 模拟机与训练`、`4 技术提示`；Heading 2 为带编号节名；Heading 3 为章内骨架小节。
- 模式 A：Heading 1 为带代号块名；Heading 2 为块内条目。
- 页脚居中显示“第 X 页　共 Y 页”，灰色小号字。除独立文档的块名页眉外，不增加其他页眉页脚内容。
- 表头跨页重复，表格行尽量不跨页；含合并单元格的大表必要时从新页开始。
- 生成后必须检查知识点数量和顺序、关键原文覆盖、机型差异、安全关键数据、表格列语义和 OOXML 完整性。
- 表格必须做“原始单元格 → 成品单元格”逐格对照：原始单元格有文字时，对应成品单元格不得为空，且文字不得落入相邻错格。要区分原稿本就留空、合并单元格的被覆盖位置、以及解析遗漏；前两者不得误报成缺失。
- 单独的 `-`、`—` 等占位符是原始单元格内容，不得按项目符号删掉而生成空白格。全文文字覆盖率不能代替逐格校验。
- 必须渲染并逐页查看 PNG，检查目录、字体、乱码、裁切、重叠、空白、分页、表格尺寸、合并单元格和允许色彩。结构检查不能替代视觉检查。
- 以出版物级别做最终版式复核：标题层级、字号、行距、段前段后、表头、表格宽度、列宽、页眉页脚、分页和留白全篇一致；不靠新增说明文字或装饰性内容实现“美化”。

## Markdown 是源，DOCX 是产物

```text
FlightNotes/
├── src/
│   ├── 00_quickref.md
│   ├── 01_systems/01.06_electrical.md
│   ├── 02_ops/
│   └── 03_sim/
├── inbox.md
├── build/
└── scripts/build_docx.js
```

文件名使用英文或拼音；中文标题写在 front matter 中。纯文本源用于本地 git、diff、增量更新和检索。

## HTML 表格约定

Markdown 源中的成品表格使用内嵌 HTML，不使用不支持合并单元格和样式语义的原生 Markdown 表格。

```html
<table class="ftn">
  <tr class="hdr"><th>类别</th><th>项目</th><th>内容</th></tr>
  <tr class="premise"><td colspan="3">存在之一即<em>不适用</em>：</td></tr>
  <tr><td rowspan="2">天气</td><td>道面</td><td>跑道状况代码 <em>4 及以下</em></td></tr>
</table>
```

语义约定：

- `class="hdr"`：表头行。
- `class="premise"`：共性前提通栏行。
- `class="note"`：末尾补充通栏行。
- `<em>`：红色加粗。
- `<strong>`：黑色加粗。
- `rowspan` / `colspan`：合并单元格。

构建脚本按 SD-4 的内容职责和最长有意义信息单元估算初始列宽，再依据渲染调整，不按累计字数比例分配。

## 章节 front matter

```yaml
---
id: "01.06"
title_cn: "电气"
aircraft: ["NG", "738"]
source: ["FCOM", "FCTM"]
related: ["03.C.06"]
updated: "2026-09-04"
---
```

无法生成文件的平台直接输出完整内嵌 HTML，供用户保存为 `.md` 或粘贴到 Word / OneNote。

## 构建与验证

使用当前 skill 根目录中的命令：

```bash
node scripts/build_docx.js src/01_systems/01.06_electrical.md build/01.06_electrical.docx
```

脚本必须校验输入、创建缺失的输出目录、失败时返回非零状态，并生成 A4 横版、自动目录、内置标题样式、重复表头、自适应表格尺寸和页码字段。任何成品交付前仍须按 `documents` skill 渲染并逐页检查。
