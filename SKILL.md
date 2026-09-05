---
name: flight-theory-notes
description: "整理中文飞行理论笔记、手册摘录或截图，制作 Word 学习文档或速查表，并维护已有笔记库的分类、编号和交叉引用。不用于英文航空文档翻译、培训 PPT 或替代现行运行文件。"
---

# 飞行理论笔记

仓库名为 flight-notes-toolkit，安装与调用名为 flight-theory-notes。完整仓库是可安装包，工具和规则使用同一版本。

## 开始任务

完整读取 [整理规则](prompt/rules.md) 和 [长期决策](prompt/standing-decisions.md)。内容忠实约束优先，其次是明确适用的长期决策。默认按用户已确认的范围连续执行，不重复询问已有选择。

使用用户指定的笔记目录；未提供主库时可以先整理用户有权提供的单节材料，不搜索或上传无关私有资料。大规模治理先交付规则规定的映射表，取得具体映射批准后再修改。

实测文件读写、Python、Node、docx、PDF 转换和 CJK 字体能力。缺少某项时只限制依赖它的步骤；可以执行的检查仍然执行。没有渲染与目视复核，不能称为已验证 DOCX。

## 执行与交付

- 脚本位于本安装目录的 tools/，笔记正文位于用户的独立私有项目。运行命令时使用实际目录，不能假设固定的桌面路径。
- 源检查用 tools/check_src.py；组装用 tools/assemble.py；制作与检查用 tools/build_docx.js 和 tools/verify.py。涉及入库时读取目标仓库 AGENTS.md。
- 现有 sync.sh 可能提交并推送；仅在本次授权覆盖这些动作时使用相应模式。只做本地处理时直接执行单独构建／校验步骤。
- 保留原始技术数据、来源和用户未授权改动的内容，按规则交付三份问题清单。
- 只支持对话的 AI 可用 [完整 Prompt](prompt/SKILL.md) 处理允许提供的材料；不能据此宣称有本机、私有仓库或渲染能力。

## 维护

仅在用户明确要求或已有适用维护授权时修改规则。prompt/rules.md 是规则真源；prompt/SKILL.md 为自动导出物，运行 tools/export_prompt.py 更新，再运行 tools/check_package.py 和相关测试。公开发布前使用仓库外私有关键词表进行检查，不上传笔记正文或公司素材。
