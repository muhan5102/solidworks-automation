# SolidWorks 自动化工作区

本工作区整理自 Codex 会话工作区的 `outputs` 目录，用于驱动本机 SolidWorks 2024
（中文版）参数化建模。

## 运行环境

- Python：`D:\Python\python.exe`（已安装 pywin32、numpy、PIL，无需迁移依赖库）
- SolidWorks：2024 中文版，模板
  `C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates\gb_part.prtdot`

## 目录结构

| 目录 | 内容 |
|---|---|
| `gripper/` | UMI 鱼鳍软爪：成功版脚本 `build_umi_fin_gripper.py`、慢速演示版 `build_umi_fin_gripper_showcase.py`、对应 bat、轮廓预览脚本 `preview_profile.py` |
| `prompt/` | `SolidWorks自动化Prompt.md`：SolidWorks 自动化提示词/规范文档 |
| `tools/` | `auto_run_and_report.ps1`（bat 自动执行并自反馈）、`install_deps.bat` |
| `archive_reducer/` | 一级减速器旧方案（用户评价不合格，仅归档参考，不要作为成果交付） |
| `previews/` | 历史轮廓预览图 |

## 使用方法

- 鱼鳍爪建模：双击 `gripper\run_gripper.bat`（用 D:\Python 运行，连接已打开的 SolidWorks）
- 轮廓预览（不启动 SolidWorks）：`D:\Python\python.exe gripper\preview_profile.py`

## 已知要点（务必遵守）

1. SolidWorks 必须以**可见交互窗口**运行；由脚本 `Dispatch` 拉起的无窗口实例画草图正常，
   但 `FeatureExtrusion2` 会静默返回 None，无法拉伸。
2. 拉伸成功路径：画完草图 → 退出草图 → 选中草图特征 →
   `FeatureExtrusion2(True, False, False, 0, 0, mm2m(厚), 0, ...)`。
3. 草图外轮廓必须无自交、无多余线段、端点闭合，否则拉伸失败。
4. API 底层单位为米，必须 `mm2m()`；复杂草图用 `AddToDB=True / DisplayWhenAdded=False`；
   中文路径保存失败时用 ASCII 回退。
5. 前视基准面选择用 `FeatureByPositionReverse` 遍历（`SelectByID2` 本机不可用）。
6. 装配体自动化未支持：须"插入零件 + 配合"一步步完成，禁止用绝对坐标定位代替配合。
7. 分步建模原则：先画外轮廓 → 用户确认 → 再加孔/细节 → 再拉伸。
