# SolidWorks 自动化工作区

本工作区整理自旧 Codex 会话的 `outputs` 目录，保留鱼鳍软爪脚本、预览和一级减速器旧方案。
它用于识图建模实践、流程改进与历史归档，不是当前自动化工具的安装目录。
日常 SolidWorks 操作使用已安装的 `solidworks-automation` Skill/MCP。

## 识图建模实践

执行本项目任务遵守 [建模与整套题基本要求](AGENTS.md)，实际练习与测试使用
[实践记录模板](docs/建模实践记录模板.md)，通用要求的实质修改记入
[更新记录](docs/建模要求更新记录.md)。后续实践可持续维护这些文本；新想法先留案例证据，
验证后再合并规则，正式测试期间保持使用的规则版本不变。
早期真题和王静图集用于比较方法、改进质量与速度；近年省赛、国赛留作整套测试，
具体留出届次尚待确定。已接触题须如实记录，不用重复熟练后的结果覆盖首次测试。

已开始第8届连接件首轮练习，见 [实践记录](practice/year08/connector/实践记录.md)、
[阶段建模脚本](practice/year08/connector/build_connector.py) 和
[核验脚本](practice/year08/connector/verify_connector.py)。当前还未进行整卷测试。
随后完成较简单的[手柄精简流程试跑](practice/year08/handle/实践记录.md)，用于验证已有回转方法复用；其中旧10分钟预算已由后续规则替代。
旧鱼鳍软爪与减速器不作为新实践的基线。
常用件优先核对已有的[参数化模板与调用验证](docs/常用件模板调用.md)，减少齿轮、弹簧等重复建模。

截至2026-11-16，已开展连接件、阀杆、滑阀及多种盘类件练习，当前规则版本0.12。
按[当前进度与成果入口](practice/训练进度.md)定位脚本和案例，近期重点见
[盘类件流程与复核复盘](docs/2026-11-16-盘类件流程与复核复盘.md)。
[11月14日总结](docs/2026-11-14-训练总结与已验证方法.md)和
[MCP连接修复记录](docs/2026-11-15-MCP连接修复.md)保留各自验证范围；
执行秒数、同题脚本复用和新题完整用时分别记录，不能相互替代。

## 目录与入口

| 目录 | 用途与状态 | 主要入口 |
|---|---|---|
| `docs/` | 识图建模实践模板与规则更新 | [实践记录模板](docs/建模实践记录模板.md)、[更新记录](docs/建模要求更新记录.md) |
| `practice/` | 新实践的脚本与复盘 | [第8届连接件](practice/year08/connector/实践记录.md) |
| `gripper/` | 鱼鳍软爪范例：参数、几何预检、建模与慢速演示 | [建模脚本](gripper/build_umi_fin_gripper.py)、[慢速演示](gripper/build_umi_fin_gripper_showcase.py)、[轮廓预览脚本](gripper/preview_profile.py) |
| `previews/` | 历史预览，不能作为当前模型的验收证据 | [旧软爪轮廓](previews/profile_preview.png)、[旧齿轮轮廓](previews/gear_profile_preview.png) |
| `prompt/` | 2026-08 的提示词归档，含旧环境限制和路径 | [历史提示词](prompt/SolidWorks自动化Prompt.md) |
| `tools/` | 旧计划任务反馈工具与可选依赖安装脚本 | [计划任务反馈](tools/auto_run_and_report.ps1)、[依赖安装](tools/install_deps.bat) |
| `archive_reducer/` | 未获认可的一级减速器旧方案，仅供追溯 | [旧设计说明](archive_reducer/reducer_design_report.md)、[旧参数](archive_reducer/reducer_design.json)、[旧几何脚本](archive_reducer/reducer_geometry.py) |

原生练习产物与临时证据放在本机 `outputs/`，通过仓库 `.gitignore` 保持未纳管；
历史归档不含原生模型，已提交的课程报告和课程 CAD 成果不收纳到这里。
`gripper/profile_preview.png` 是范例配套预览；它与 `previews/profile_preview.png` 的内容不同，两者均保留。

## 使用鱼鳍软爪范例

在仓库根目录运行：

```powershell
# 只运行几何预检，不连接 SolidWorks
python .\gripper\build_umi_fin_gripper.py --check-only

# 重生成 gripper/profile_preview.png，不连接 SolidWorks
python .\gripper\preview_profile.py
```

需要实际建模时，先检查脚本中的尺寸和保存文件名，再运行
[run_gripper.bat](gripper/run_gripper.bat)；慢速演示使用
[run_showcase.bat](gripper/run_showcase.bat)。这两个入口均执行同目录脚本。

- 范例目标环境：SolidWorks 2024 中文版；使用已安装依赖的 Python 环境。历史批处理保留原环境入口，换机器时需先调整。
- 几何预检和二维预览仍会导入 `numpy`、`pywin32`，预览另需 Pillow；实际建模还需 SolidWorks 和有效模板。
- `build_log.txt` 写在 `gripper/`；自动运行器的 `run_status.txt`、`run_console.log` 也写在该目录，均已忽略。
- 原生零件沿用脚本的桌面保存位置及失败回退方式，未迁移到仓库。运行前确认不会覆盖需要保留的同名模型。
- 范例的旧 API 写法和坐标草图只代表当时实现，不作为当前 MCP 能力、参数可编辑性或工程合格的证明。

## 旧工具与归档

`tools/auto_run_and_report.ps1` 依赖预先注册的计划任务。2026-11-14 检查时，
默认任务 `CodexSW_Gripper_Auto` 仍指向旧会话目录，与当前仓库入口不同。
该工具及任务配置原样保留；不要据此认定它会执行本仓库脚本。
`tools/install_deps.bat` 会安装 Python 包，仅在确有缺失时使用。

`archive_reducer/` 保持原有文件与相互引用，仍包含旧会话输出路径。
其脚本、参数、说明及预览共同作为历史记录，未重新建模或验收，不应直接运行作为当前交付。

## 本机经验记录

通用兼容经验继续保存在仓库外的既有本机资料目录：

- SolidWorks 2024 本机经验导航：`$CODEX_HOME/local-reference/solidworks-2024/compatibility-notes.md`。
- Z轴与输送带建模复盘：上述目录下的`references/2026-11-14-z-axis-conveyor-retrospective.md`。

`$CODEX_HOME`表示本机Codex资料根目录；这些文件不属于仓库，复制仓库到其他机器后，需另行提供对应资料。
它们保持普通参考记录身份，不复制到课程项目，也不直接写入新版 Skill。

## 声明与引用

设计依据、第三方标准引用、AI 协作说明与免责声明见 [NOTICE.md](NOTICE.md)。
本仓库内容**未经工程验证**，不得直接用于生产或任何需要承担安全责任的场合。
