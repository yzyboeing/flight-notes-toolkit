# flight-notes-toolkit

把零散的飞行理论笔记整理成**可查、可维护、可版本管理**的结构化手册的一套方法与工具链。

面向 Boeing 737-NG / 737-8 的中文学习笔记（系统知识、非正常程序、模拟机参考、技术提示、CRM、规章），但整理规则与渲染器与机型无关，改一套归属判据就能用在别的机型或别的技术文档上。

> 本仓库**只包含方法论与工具**。笔记正文属于个人/公司资料，放在单独的私有仓库，不在这里。

---

## 这套东西解决什么问题

手写笔记、课件截图、通告 PDF 攒了几十页之后，典型状态是：同一个数值在三个地方写法不同、找一条限制要翻五页、想加一条新内容不知道该放哪。

这套工具的思路是把笔记拆成两层：

- **内容层** —— 一节一个 markdown 文件，用极少的标记表达「这是阈值」「这是共性前提」「这是补充说明」，人可读、可 diff、可版本管理
- **呈现层** —— 一个渲染器把这些文件编译成排版规整的 A4 横版 Word 手册：自动列宽、自动对齐、表格不跨页、超链接目录

改内容只动 markdown，排版规则改一次全库生效。

---

## 仓库结构

```
prompt/SKILL.md       整理规则（真源）——格式约定、排版规则、归属判定、校验流程
tools/assemble.py     从源文件组装章节 md 与全书 md
tools/build_docx.js   md + 内嵌 HTML 表格 → A4 横版 docx
tools/docx2md.py      既有 docx → 源格式（无损保留加粗/红字/合并单元格）
tools/verify.py       交付前 PDF 校验
examples/             源格式示例
```

---

## 源格式

markdown 骨架 + 内嵌 HTML 表格，只有六个标记：

| 写法 | 渲染 | 用途 |
|---|---|---|
| `<em>…</em>` | 红色加粗 | 阈值、限制、禁止性措辞、判据分岔点 |
| `<strong>…</strong>` | 黑色加粗 | 关键术语、动作、结论 |
| `<tr class="hdr">` | 灰底表头，跨页重复 | 表头行 |
| `<tr class="premise">` | 浅灰底 + 左侧竖条 | 共性前提通栏行 |
| `<tr class="note">` | 白底通栏 | 末尾补充行 |
| `colspan` / `rowspan` | 合并单元格 | 主判据合并列 |

完整示例见 [`examples/sample_section.md`](examples/sample_section.md)。

---

## 渲染器做了什么

| 能力 | 说明 |
|---|---|
| 自动列宽 | 短列（项目、序号、参数名、数值列）按内容实际长度定宽，余量按比例全给说明类长列 |
| 自动对齐 | 按列判定：短列与数值列整列居中，说明类长列左对齐，不混排 |
| 表格不跨页 | 能放进一页的一律整表保持；超页表格逐级缩至 8.5/8/7.5/7pt，**只在真能塞下时才缩** |
| rowspan 正确性 | 用占位网格推算每个单元格的真实起始列，列宽计算与渲染共用同一套 |
| 嵌套标记 | `<strong>…<em>X</em></strong>` 递归解析，红色优先 |
| 目录 | Word `TOC \h \o "1-2"` 域，超链接，章 + 全部小节 |
| 封面 / 页码 | 粗线 + 小初号标题 + 副标题 + 作者行；页码右下角小五号，封面不显示 |

---

## 用法

```bash
npm install docx
python3 tools/assemble.py                # 源文件 → build/mod0..5.md + build/full.md

DOC_AUTHOR="by　某某" NODE_PATH=./node_modules \
  node tools/build_docx.js build/full.md build/manual.docx

soffice --headless --convert-to pdf --outdir build build/manual.docx
python3 tools/verify.py build/manual.pdf   # 空白页 / 标签泄漏 / 项目符号 / front matter
```

生成的 docx 打开后按 **Ctrl+A → F9** 刷新目录（TOC 是 Word 域，首次打开为空）。

`tools/assemble.py` 里的 `CH` 字典定义了模块划分与文件顺序，换用到别的项目改这里。

---

## 内容原则

这套规则最核心的部分不是排版，是这几条：

- **不丢失任何原始信息、不改变技术含义、不基于臆测添加内容**
- **数值逐字保留**——不换算、不四舍五入、不「修正」看起来异常的值
- **安全关键数据只标不改**（决断高、高度、速度、限制、频率、跑道数据）：发现可疑之处照录原文并标出疑点，由使用者按现行手册裁定。被「自动纠正」错的数据比原始错误更危险，因为它看起来像被核实过
- **技术冲突不自行裁决**，列出双方原文与出处，请使用者判定
- **原始材料不完整时不猜写**，标 `【缺失，待原件补充】`

细则见 [`prompt/SKILL.md`](prompt/SKILL.md)。

---

## 依赖

- Node.js + [`docx`](https://www.npmjs.com/package/docx) ≥ 9
- Python 3 + `python-docx`（仅 docx2md 需要）、`pymupdf`（仅 verify 需要）
- LibreOffice（仅校验环节需要，用于转 PDF）

---

## 免责

本仓库提供的是文档整理方法与工具。用它整理出的任何飞行技术资料**不得作为运行依据**，一切以所属运营人现行有效的手册、通告与规章为准。
