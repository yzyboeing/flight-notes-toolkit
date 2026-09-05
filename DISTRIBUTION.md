# 版本化分发与环境检查

## 使用者

每个 Skill 的 `skill-package.json` 声明安装名、源码根目录、允许分发的文件、规则文本范围和依赖阶段。版本号是 **Skill 包版本**，不替代配置版本或笔记版本。标签为 `skill-vX.Y.Z`，六个包可独立升级；私有内容库不打包。

每次发布包含：

- `<skill>-X.Y.Z.skill.zip`：单一、正确命名的 Skill 目录；规则及脚本齐全，不捆绑 Python、Node 或模型。
- `<skill>-X.Y.Z.rules.txt`：按原文件边界汇总的规则及引用，供文本环境使用；不得绕过原始加载路由或长度限制。
- `<skill>-X.Y.Z.sha256`：ZIP 与文本的校验值。应从可信仓库的同版本 Release 获取；同一不可信来源的哈希不能证明作者身份。
- ZIP 内 `SOURCE.json`：源提交、构建器提交、逐文件 SHA-256。`doctor.py` 是统一的只读探测入口。

使用本仓库的安装器（Python 3.10+，仅标准库）：

```bash
python3 tools/distribute_skills.py install <下载的包.skill.zip> \
  --sha256 <可信sha256文件中该ZIP的64位校验值> \
  --destination <目标AI的skills父目录>
```

它校验外部哈希、逐文件哈希、路径穿越、重复成员和符号链接，拒绝覆盖已有安装。更新前把**指定旧 Skill 目录**备份到发现目录之外，再安装；不删除原始项目或其他 Skill。它不联网、不登录、不自动安装依赖。

Codex 的当前官方通用用户目录为 `~/.agents/skills`；已有安装应先检查实际发现位置，避免在不同目录重复安装同名包。Claude Code 使用 `~/.claude/skills`。其他宿主按其官方文档配置。参考：[OpenAI](https://learn.chatgpt.com/docs/build-skills)、[Claude Code](https://code.claude.com/docs/en/skills)、[Agent Skills 格式](https://agentskills.io/specification)。网页版/云任务不因此获得本机目录。

安装后在包目录运行：

```bash
python3 doctor.py --stage rules
python3 doctor.py
```

退出码：0 表示所选阶段的已声明依赖被发现；1 表示有依赖缺失；2 表示清单/阶段/路径错误。使用真正执行任务的 Python 解释器。各阶段独立，缺少渲染不阻断纯文本规则阅读。报告始终保留 `functional_test: not_run` 和 `visual_qa: not_run`；按原 Skill 实际执行后，另行记录结果，不把探测值改成成品验收。

文档包使用 `npm ci --ignore-scripts` 安装锁定依赖，或配置兼容的 NODE_PATH。中文字体必须用真实渲染图验证。VBS 继续用包内阶段检测器验证 JZSub/ASR；飞训包继续用资源检查器核对本地配置。配置必须由用户另行提供脱敏输入，不能从设备提取凭据。

## 维护者

1. 在源码仓库修改规则/工具；同步包清单中的文件允许列表和版本。先跑原有完整检查、隐私检查及回归测试。
2. 提交修改，记录源提交。构建器只读取 Git 提交中的允许文件；不包含未提交内容、缓存、媒体、公司母版或配置。检查打包清单，允许列表不是内容保密扫描的替代品。
3. 在已提交的本工具库运行以下命令，输出目录放仓库外：

```bash
python3 tools/distribute_skills.py build --repo <目标源码仓库> \
  --ref <完整提交号> --out <新的输出目录>
```

构建器与 doctor 来自工具库 HEAD 的确定提交。新输出不覆盖旧文件；同一源提交和构建器提交产生一致 ZIP。不要在构建器存在未提交修改时发布。

4. 解压验证、运行 doctor、执行有代表性的实际案例；分别记录格式、依赖、脚本、模型行为、视觉/设备结果。未测就是未测。
5. `git fetch` 核对远端，正常推送；为该提交创建 `skill-vX.Y.Z` Release，上传这三个文件。标签和已发布资产不覆盖；修正用新补丁版本。JimmyLoon/Surge 的 Skill Release 不取代配置 Latest。
6. 核对 Release 对应提交、资产哈希，再同步本地已选平台安装。修改兼容性报告，注明样例、平台/模型、源码版本和未覆盖范围。

本工具不自动提交、推送、删除旧版本、操作客户端或创建远程执行服务。
