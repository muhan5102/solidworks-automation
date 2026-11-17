# 仓库维护工具

在仓库根目录执行：

```powershell
python -X utf8 -m unittest discover -s tests -v
python -X utf8 tools/check_repository.py
```

标准库回归测试用临时样本验证语法、公共导入、包路径和断链的拒绝及修复后恢复，
并确认检查不执行案例脚本；与实际仓库检查一同纳入CI，无CAD或第三方Python依赖。
差异检查由[check_changed_lines.py](check_changed_lines.py)执行，使用真实Git样本验证完整PR/推送范围；
强推后若旧基准未被检出，先按SHA获取，无法获取时明确失败，不降级为仅检查最后一条提交。

检查当前维护文档的相对链接、Python语法、公共导入符号及公共包路径。
不执行任何案例脚本、不连接SolidWorks，也不据此判定模型几何正确。
本机迁移清单、旧新路径对照和搬迁哈希留在 `outputs/maintenance/`，不作为长期代码纳管。

[Repository checks工作流](../.github/workflows/repository-check.yml)在PR、推送到`main`及手动触发时调用同一命令，
并检查相应提交范围的Git差异格式。运行环境为GitHub托管Linux与Python 3.12，只需Python标准库。
基准获取使用Git和托管环境已有的GitHub CLI，凭证仅使用当前作业的只读令牌，不写入Git配置。
协作采用工作分支与PR，流程及验收边界见[仓库协作与检查](../docs/guides/仓库结构与产物管理.md#仓库协作与检查)。
