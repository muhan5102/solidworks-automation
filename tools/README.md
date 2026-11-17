# 仓库维护工具

在仓库根目录执行：

```powershell
python -X utf8 tools/check_repository.py
```

检查当前维护文档的相对链接、Python语法、公共导入符号及公共包路径。
不执行任何案例脚本、不连接SolidWorks，也不据此判定模型几何正确。
本机迁移清单、旧新路径对照和搬迁哈希留在 `outputs/maintenance/`，不作为长期代码纳管。

[Repository checks工作流](../.github/workflows/repository-check.yml)在PR、推送到`main`及手动触发时调用同一命令，
并检查相应提交范围的Git差异格式。运行环境为GitHub托管Linux与Python 3.12，只需Python标准库。
协作采用工作分支与PR，流程及验收边界见[仓库协作与检查](../docs/guides/仓库结构与产物管理.md#仓库协作与检查)。
