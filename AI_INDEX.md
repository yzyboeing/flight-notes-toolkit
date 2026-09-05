# AI 统一入口

这里导航七个仓库：六个可安装 Skill，加一个私有内容项目。入口不授予私有仓库访问、资料上传、执行脚本或发布权限。先确认任务与目标，再按对应 SKILL.md 的路由读取资料；不要加载所有 Skill。

| 项目 | 类型 / 权限 | 执行入口 | 版本包 |
|---|---|---|---|
| [flight-doc-translate](https://github.com/yzyboeing/flight-doc-translate) | Skill / 公开 | 根 SKILL.md | flight-doc-translate |
| [flight-notes-toolkit](https://github.com/yzyboeing/flight-notes-toolkit) | Skill / 公开 | 根 SKILL.md | flight-theory-notes |
| [flight-theory-notes](https://github.com/yzyboeing/flight-theory-notes) | 笔记内容 / 私有 | 项目 AGENTS.md；配合上一行工具 | 不分发内容包 |
| [feixun-ppt](https://github.com/yzyboeing/feixun-ppt) | Skill / 私有 | 根 SKILL.md；母版单独本地配置 | feixun-ppt |
| [jimmyloon](https://github.com/yzyboeing/jimmyloon) | Skill / 私有 | 根 SKILL.md；脱敏配置另行提供 | jimmyloon |
| [jimmysurge](https://github.com/yzyboeing/jimmysurge) | Skill / 私有 | 根 SKILL.md；脱敏配置另行提供 | jimmysurge |
| [vbs](https://github.com/yzyboeing/vbs) | Skill / 公开 | skills/vbs/SKILL.md | vbs |

## 开始使用

1. 到目标仓库的 Releases 找 `skill-vX.Y.Z`，不是旧配置版本标签。下载同版本 `.skill.zip` 和 `.sha256`；仅接受文本的环境选择 `.rules.txt`。
2. 完整 ZIP 含规则、引用、脚本、环境探测与 SOURCE.json；保持顶层 Skill 目录名。不可只复制 SKILL.md。独立 rules.txt 从同一源码生成，不是另一份维护真源。
3. 在目标 AI 支持的 Skill 目录安装。Codex 和 Claude Code 的目录与操作不同；没有原生 Skill 功能时可以显式要求读取目录里的 SKILL.md，但不保证自动发现或工具执行。
4. 用任务解释器运行 `python3 doctor.py --stage rules`，再探测所需阶段。按 next_checks 执行更具体的检查。发现依赖不等于制作、转写、设备验证或视觉 QA 成功。
5. 明确当前任务允许读哪些源文件、改哪些输出；遵守原 Skill 审核关口。规则更新不自动授权提交、推送或上传材料。

完整安装与发布方法见 [DISTRIBUTION.md](DISTRIBUTION.md)，实际兼容性证据见 [COMPATIBILITY.md](COMPATIBILITY.md)。

## 给 AI 的通用开场

> 按 AI_INDEX.md 选择与本任务对应的 Skill。报告实际读取的入口、SOURCE.json 中的版本与提交、所需引用是否齐全，以及本环境能完成和不能完成的步骤。私有文件只读取我明确放入范围且允许在本平台处理的内容。不把链接可见、依赖发现或文本分析当作实际执行成功。

讨论与执行分开时，交接：源文件清单、Skill 版本、已确认术语、排版要求、允许修改的范围、未决问题、已批准的具体版本、后续审核关口和验收标准。不要把讨论建议冒充批准。

## 保密与更新

私有仓库授权不等于允许把内容发送给任意云端模型。尤其 feixun-ppt 按其仓库规则禁止上传第三方 AI 平台；母版不随包分发。配置包不包含活动配置或连接凭据。公开目录只列导航信息，不包含私有规则正文。

七个仓库不合并；GitHub 源码为维护依据，安装副本来自确定提交。平台适配只处理安装和工具差异，不复制、改写审核与技术规则。
