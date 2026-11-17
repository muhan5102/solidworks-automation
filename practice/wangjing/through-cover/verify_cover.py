"""LJT04.33集中重开终检；预期值按PDF94/印刷88原图核对。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path.home() / ".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_review import collect_model_summary, save_preview
from sw_sketch_quality import inspect_model_sketches

out = Path(sys.argv[1]).resolve()
target = out / "透盖_LJT04.33.SLDPRT"
started = datetime.now().astimezone().isoformat()
timer = time.perf_counter()
s = win32com.client.Dispatch("SldWorks.Application.32")
assert g(s, "GetProcessID") == json.loads((out / "session.json").read_text(encoding="utf-8"))["process_id"]
m = g(s, "ActiveDoc")
assert Path(g(m, "GetPathName")).resolve() == target
e = win32com.client.VARIANT(16387, 0)
w = win32com.client.VARIANT(16387, 0)
assert m.Save3(1, e, w) and e.value == 0
s.CloseDoc(g(m, "GetTitle"))
assert g(s, "GetFirstDocument") is None
m = s.OpenDoc6(str(target), 1, 1, "", e, w)
assert m is not None and e.value == 0
assert Path(g(m, "GetPathName")).resolve() == target
assert m.ForceRebuild3(False)
health = collect_model_summary(m)
quality = inspect_model_sketches(m)
faces = []
for body in g(m, "GetBodies2", 0, False):
    for face in g(body, "GetFaces"):
        surface = g(face, "GetSurface")
        row = {"box": [v * 1000 for v in g(face, "GetBox")]}
        for kind in ["Cylinder", "Cone"]:
            if g(surface, "Is" + kind):
                row[kind] = list(g(surface, kind + "Params"))
        faces.append(row)

def near(a, b):
    return abs(a - b) < 1e-5

def cylinders(radius, x, z, y0, y1):
    return [f for f in faces if "Cylinder" in f
            and near(f["Cylinder"][6] * 1000, radius)
            and near(f["Cylinder"][0] * 1000, x)
            and near(f["Cylinder"][2] * 1000, z)
            and near(f["box"][1], y0) and near(f["box"][4], y1)]

manual = [sk for sk in quality["sketches"] if sk["name"] != "沉孔向导轮廓"]
generated = [sk for sk in quality["sketches"] if sk["name"] == "沉孔向导轮廓"]
d = g(m.FeatureByName("安装沉孔"), "GetDefinition")
checks = {
    "单实体无特征错误警告": len(g(m, "GetBodies2", 0, False)) == 1 and not health["faulty_features"] and not health["warning_features"],
    "两张手工草图完全定义无Fix无孤立": len(manual) == 2 and all(sk["status"] == "pass" for sk in manual),
    "孔向导原生内部截面单列核验": len(generated) == 1 and generated[0]["constrained_status"] == 3 and generated[0]["fix_count"] == 1 and not generated[0]["dangling"],
    "法兰D120厚13": len(cylinders(60, 0, 0, 0, 13)) == 1,
    "凸台D80长5总长18": len(cylinders(40, 0, 0, 13, 18)) == 1,
    "背孔D68深5": len(cylinders(34, 0, 0, 13, 18)) == 1,
    "D35前孔及环槽口两端": len(cylinders(17.5, 0, 0, 0, 3.75)) == 1 and len(cylinders(17.5, 0, 0, 9.25, 13)) == 1,
    "槽底D48宽4居中于法兰": len(cylinders(24, 0, 0, 4.5, 8.5)) == 1,
    "仅两个梯形槽侧锥面无新增倒角": len([f for f in faces if "Cone" in f]) == 2,
    "孔向导D9贯穿D16深6且无额外锥沉孔": near(g(d, "ThruHoleDiameter"), .009)
        and near(g(d, "CounterBoreDiameter"), .016) and near(g(d, "CounterBoreDepth"), .006)
        and g(d, "EndCondition") == 1 and not g(d, "NearSideCounterSink") and not g(d, "FarSideCounterSink"),
    "一个原生六孔阵列": sum(f["type"] == "CirPattern" for f in health["features"]) == 1
        and g(g(m.FeatureByName("六组安装孔阵列"), "GetDefinition"), "TotalInstances") == 6,
}
for i in range(6):
    angle = math.pi / 2 + i * math.pi / 3
    x, z = 49 * math.cos(angle), 49 * math.sin(angle)
    checks[f"第{i+1}孔D98分布圆D16入口0至6及D9余段6至13"] = (
        len(cylinders(8, x, z, 0, 6)) == 1 and len(cylinders(4.5, x, z, 6, 13)) == 1)

# 持久化草图回读原始宽度驱动，槽口宽5.5并非误作矩形环槽。
f = m.FeatureByName("轴向截面")
dd = g(f, "GetFirstDisplayDimension")
dims = []
while dd is not None:
    dim = g(dd, "GetDimension2", 0)
    dims.append({"name": g(dim, "FullName"), "value": g(dim, "SystemValue"), "driven": g(dim, "DrivenState")})
    dd = g(f, "GetNextDisplayDimension", dd)
checks["槽根4与槽口5_5为原生驱动"] = all(
    any(near(v["value"], value) and v["driven"] == 2 for v in dims) for value in [.004, .0055])
save_preview(m, str(out / "透盖_等轴测.bmp"), "isometric")
save_preview(m, str(out / "透盖_侧视.bmp"), "front")
m.ShowNamedView2("*Bottom", 6)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert m.SaveBMP(str(out / "透盖_沉孔入口.bmp"), 1600, 1000)
m.ShowNamedView2("*Isometric", 7)
m.GraphicsRedraw2()
report = {
    "started": started, "finished": datetime.now().astimezone().isoformat(),
    "seconds": time.perf_counter() - timer, "passed": all(checks.values()),
    "checks": checks, "faces": faces, "dimensions": dims, "quality": quality, "health": health,
    "limits": "待目视对原图；未制作工程图、装配，STEP未重导入。原生孔向导截面的内部Fix单列核验。",
}
(out / "verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"passed": report["passed"], "failed": [k for k, v in checks.items() if not v],
                  "seconds": report["seconds"]}, ensure_ascii=False))
assert report["passed"]
