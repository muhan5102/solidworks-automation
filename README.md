# SolidWorks 自动化实践

本项目长期积累工程图识读、参数化建模、装配和出图的方法，以及每次实践的真实质量与用时。
日常执行使用已安装的 `solidworks-automation` Skill/MCP；这里保存项目脚本与案例，不是工具安装目录。

开始工作先读 [项目要求](AGENTS.md) 和 [当前训练进度](practice/训练进度.md)，再进入本轮对应案例。
记录保留首次失败、人工提示、纠错及熟题重放的边界，不能用后来的成功覆盖首次成绩。

| 入口 | 放什么 |
|---|---|
| [practice/](practice/README.md) | 按来源、案例组织的维护脚本和实践记录；`shared/` 放多案例共用函数 |
| [docs/](docs/README.md) | 指南、模板、训练复盘、工具验证及规则变更 |
| [tools/](tools/README.md) | 小型仓库维护工具；检查工具只读，不连接CAD |
| `outputs/`（本机未纳管） | 原始运行证据、原生模型、干净交付副本、交接和本机维护备份 |

目录职责、命名和产物流转见 [仓库结构与产物管理](docs/guides/仓库结构与产物管理.md)。
原始图集、试卷、参考模型和已提交课程成果保留原位，不复制进本仓库。

## 练习方式与成果

- **王静图集**用于单件练习，重点是三视图识读、按图纸定义草图、选择特征路线和提高首次建模正确率。
- **成图真题**用于综合练习，按题面安排零件建模、指定零件图、装配和装配图，分别记录限时完成项、续作和实际核验结果。

案例示例见 [王静端盖 LJT04.37](practice/wangjing/end-cover-37/实践记录.md) 和 [第8届手动换向阀三维装配](practice/year08/valve-assembly/实践记录.md)。
各项交付的完成状态、用时和验证范围统一见 [训练进度](practice/训练进度.md)。

## 运行与协作

从仓库根目录运行只读检查：

```powershell
python -X utf8 -m unittest discover -s tests -v
python -X utf8 tools/check_repository.py
```

首次接手看[本机环境与检查](docs/guides/本机环境与接手检查.md)，具体案例按[运行入口](practice/运行入口.md)核对阶段。
GitHub保存维护代码与记录，模型和运行证据的独立备份、恢复结果另留本机。

项目改动默认通过工作分支和PR审核合并。[Repository checks](.github/workflows/repository-check.yml)
在PR、推送到`main`和手动触发时自动执行仓库检查；本地与CI使用同一检查脚本。
具体授权节奏与检查范围见[仓库协作与检查](docs/guides/仓库结构与产物管理.md#仓库协作与检查)。
参与维护和提交PR见[CONTRIBUTING.md](CONTRIBUTING.md)，PR填写问题、改动和实际验证结果。

执行建模脚本前，按案例核对参数、软件会话与全新的输出目录。旧案例存在阶段续作和现场修复入口，
不能把整个目录的Python文件依次执行。`outputs/` 中的现场脚本是历史证据，维护入口以 `practice/` 为准。

本机兼容经验入口为 `$CODEX_HOME/local-reference/solidworks-2024/compatibility-notes.md`，
它保持普通参考资料身份，不随本仓库分发。来源和验证边界见 [NOTICE.md](NOTICE.md)。
未经当前对话明确授权不提交或推送。
