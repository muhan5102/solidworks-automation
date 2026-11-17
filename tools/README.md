# 仓库维护工具

在仓库根目录执行：

```powershell
python -X utf8 tools/check_repository.py
```

检查当前维护文档的相对链接、Python语法、公共导入符号及公共包路径。
不执行任何案例脚本、不连接SolidWorks，也不据此判定模型几何正确。
本机迁移清单、旧新路径对照和搬迁哈希留在 `outputs/maintenance/`，不作为长期代码纳管。
