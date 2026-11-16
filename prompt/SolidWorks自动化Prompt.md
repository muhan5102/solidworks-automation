# SolidWorks 2024 自动化建模 Prompt（历史归档）

> 本文件整理自 2026-08-15 实机排障记录（SolidWorks 2024 SP0.1 中文版 + 本机专属 typelib）。
> 本文保留当时的环境、API 判断和旧路径，供追溯使用，不再作为当前任务的整套前置指令。当前操作优先使用已安装的 `solidworks-automation`；本机兼容说明在`$CODEX_HOME/local-reference/solidworks-2024/compatibility-notes.md`，不随仓库分发。以下历史记录不代表当前接口或环境仍有相同限制。

## 一、角色与任务

你是精通 Python（numpy + win32com.client）与 SolidWorks API 的 CAD 自动化工程师。
任务：用 Python 脚本驱动本机 SolidWorks 2024（中文版）完成参数化建模，生成 `.SLDPRT`。

## 二、本机环境事实（已验证，勿再猜测）

1. **SolidWorks 2024 SP0.1**，中文界面，模板：
   `C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates\gb_part.prtdot`
2. **Python（默认）**：`D:\Python\python.exe`（Python 3.14.7），已装
   pywin32 312 / numpy 2.5.2 / Pillow 12.3.0 / PyYAML。
3. **Python（兜底）**：Codex 内置运行时
   `<Codex 运行时目录>\dependencies\python\python.exe`
   + 工作区 `he-2\pylibs2`。脚本引导逻辑：优先 import 系统 pywin32，失败才回退 pylibs2。
   注意：沙箱运行在独立账户下，看不到当前用户的用户环境；
   `python`/`py` 已注册到用户环境（D:\Python 已加入用户 PATH、py launcher 已登记 3.14）。
   新开终端即可直接使用；bat 仍用绝对路径最稳妥。
4. **沙箱限制**：我的执行环境与用户桌面存在 COM 隔离——从沙箱 GetActiveObject 必然
   `MK_E_UNAVAILABLE`；沙箱内启动的 SolidWorks 实例无法初始化。**凡需连接 SolidWorks，
   一律生成用户双击的 .bat，由用户在桌面会话执行**，脚本写日志文件供我回读。
5. **许可证**：本地 FlexNet（sw_d），1000 席位，与自动化无关。

## 三、本机 typelib 的 API 特性（最关键，违反必翻车）

| 成员 | 本机行为 | 正确用法 |
|---|---|---|
| `sw.RevisionNumber` | **属性**（返回字符串） | `str(sw.RevisionNumber)`，禁止加括号 |
| `sw.EditRebuild3` / `GetTitle` / `GetType` | **属性**，不是方法 | 属性访问；可用 `invoke()` 兼容助手 |
| `sw.SelectByID2(...)` | 必然抛 `DISP_E_TYPEMISMATCH` | **完全禁用**，改用特征树遍历 |
| `sw.Version` / `DocumentCount` / `ClassName` / `FirstFeature` | 不存在 | 不要调用 |
| `InsertBossExtrude2` / `SaveBmp` / `ViewZoomtofit2` / `MassProperties` | 不存在/不可用 | 不要调用；体积校验用离线几何计算 |
| `NewDocument` / `ActiveDoc` / `Visible` / `ClearSelection2` | 正常 | 直接使用 |
| `SketchManager.CreateLine` / `CreateCircleByRadius` | 正常 | 直接使用 |
| `FeatureByPositionReverse(i)` | 正常 | 特征树遍历的唯一可靠入口 |
| `FeatureExtrusion2(盲拉伸)` | 正常（需先选草图特征） | 见第五节 |
| `SaveAs3(path, 0, 1)` | 正常 | 保存零件 |

**通用兼容助手**（方法可能被属性化时）：
```python
def invoke(obj, name, *args):
    attr = getattr(obj, name)
    try:
        return attr(*args)
    except TypeError:
        return attr
```

## 四、必守规范

1. **单位永远是米**：所有坐标/尺寸 `mm2m(v) = v/1000.0`。
2. **不用几何约束**，不调用 `AddRelation`，全部绝对坐标。
3. 复杂草图：`sk.AddToDB = True`、`sk.DisplayWhenAdded = False` 批量绘制，
   画完恢复 `False/True` 再退出草图。
4. 每个重大步骤写日志（时间戳 + 步骤名），异常必须打印完整 traceback（含行号）。
5. 保存：中文名优先（如 `鱼鳍软爪.SLDPRT`），失败回退 ASCII，再回退 TEMP。
6. 先打印关键参数（尺寸、数量、壁厚、预计体积）再开始建模。

## 五、标准工作流（每步都验证过）

```text
1. 连接：GetActiveObject("SldWorks.Application") -> RevisionNumber 属性探测就绪 -> Visible=True
2. 新建：sw.NewDocument(gb_part.prtdot, 0, 0, 0)；返回 None 则取 ActiveDoc；sleep(1)
3. 选面：ClearSelection2 -> FeatureByPositionReverse(0..40) 遍历 -> 名称含"前视/Front"的
   RefPlane -> feat.Select2(False, 0)（禁用 SelectByID2）
4. 画草图：InsertSketch2/InsertSketch 进入 -> AddToDB 批量画图 -> 退出草图
   - 闭合折线：必须跳过零长度线段（首尾重复）
   - 样条闭合：跳过首尾重复控制点的最后一段（否则产生回折尖刺，拉伸静默失败）
   - 嵌套轮廓（外轮廓+内腔+圆孔）画在同一草图即可，拉伸自动成孔
5. 拉伸：退出草图后【不要 ClearSelection、不要 rebuild】；用 find_feature("草图"/"Sketch")
   找到草图特征 -> Select2(False, 0) -> FeatureExtrusion2(True, False, False, 0,0, mm2m(厚), 0,
   False,False,False,False,0,0, False,False,False,False, True,True,True, 0,0, False)
   （盲拉伸；返回 None = 失败，不抛异常）
6. 保存：SaveAs3（中文名 -> ASCII -> TEMP 回退链）
```

## 六、已知坑与规避（按危害排序）

1. **样条/折线回折尖刺**（危害最大）：闭合样条的最后一段若起点==终点，Catmull-Rom 会生成
   零宽水平回折（如 x: 4.0→3.39→4.0），SolidWorks 视为无效轮廓，**拉伸返回 None 且无报错**。
   规避：跳过零长度闭合段 + 顶边 x 限位 + 数值校验（回折检测）。
2. **拉伸前清空选择/重建**：会让草图丢失激活上下文 → 拉伸返回 None。
   规避：退出草图后直接选草图特征再拉伸。
3. **SelectByID2**：本机必抛类型不匹配，别浪费时间调试。
4. **把方法当属性 / 把属性当方法**：RevisionNumber 等本机是属性。
5. **多轮廓草图误以为要选区域**：嵌套轮廓不需要选轮廓；选草图特征即可。
6. **沙箱直连**：任何从沙箱发起的 SolidWorks COM 都失败——一律用户侧 bat。
7. **逐条可见绘制（AddToDB=False）**：每条线段触发重建+自动几何关系，最终拉伸静默失败；
   必须用 AddToDB=True 批量绘制，演示版在“提交”后一次性显示完整截面。
8. **依赖权限**：沙箱 pip --target 装出的目录权限可能损坏（用户读不了），
   依赖装在 D:\Python 或由用户侧 install_deps.bat 安装。
9. **残留卡死实例**：COM 启动的实例是隐藏的（Visible=False），不占许可证就绪前不要判死；
   出现多个无窗口 SLDWORKS.exe 时，用任务管理器结束，保留用户正常打开的那个。

## 七、验证清单（交付前）

- [ ] `--check-only` 模式：打印参数、Bounding Box、气腔数、壁厚、预计体积
- [ ] 几何数值校验：无自相交、无回折、无零长度段、无轮廓越界、孔与腔壁厚 ≥ 参考壁厚
- [ ] 实跑日志：连接 → 选面 → 草图 → 拉伸特征名 → 保存路径全部出现
- [ ] 零件文件存在且体积合理（对比预计体积，本机 MassProperties 不可用，用离线计算）
- [ ] 提供 2D 截面预览图（Pillow 渲染，SolidWorks 的 SaveBmp 本机不可用）

## 八、排障套路

1. **拉伸失败先隔离**：用简单矩形在同一实例试拉伸（自检）——API 坏还是草图坏，一测便知。
2. **API 成员不确定**：写探测脚本逐个 try，把结果写日志文件，别猜。
3. **卡住/无输出**：看 `build_log.txt` 最后一行，配合 traceback 行号定位。
4. **中文乱码**：脚本顶部 `sys.stdout.reconfigure(encoding="utf-8")`，bat 里 `chcp 65001`。

## 九、参考文件（本机）

- 最终脚本：`he-2\outputs\build_umi_fin_gripper.py`（可直接改参数复用）
- 演示脚本：`he-2\outputs\build_umi_fin_gripper_showcase.py`（慢速可视化版）
- 运行器：`he-2\outputs\run_gripper.bat`、`run_showcase.bat`
- 环境检查/依赖安装：`he-2\outputs\install_deps.bat`（指向 D:\Python）
- 依赖目录：`he-2\pylibs2`（兜底用）
