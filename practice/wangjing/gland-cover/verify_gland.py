"""LJT04.34集中重开检查：相切双耳、入口角度与实际材料范围。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_review import collect_model_summary, save_preview
from sw_sketch_quality import inspect_model_sketches

out = Path(sys.argv[1]).resolve()
target = out/"填料压盖_LJT04.34.SLDPRT"
started = datetime.now().astimezone().isoformat()
timer = time.perf_counter()
s = win32com.client.Dispatch("SldWorks.Application.32")
m = g(s, "ActiveDoc")
assert g(s, "GetProcessID") == 12356 and Path(g(m, "GetPathName")).resolve() == target
e = win32com.client.VARIANT(16387, 0)
w = win32com.client.VARIANT(16387, 0)
assert m.Save3(1, e, w) and e.value == 0
s.CloseDoc(g(m, "GetTitle"))
assert g(s, "GetFirstDocument") is None
m = s.OpenDoc6(str(target), 1, 1, "", e, w)
assert m is not None and e.value == 0 and m.ForceRebuild3(False)
quality, health = inspect_model_sketches(m), collect_model_summary(m)
faces = []
for body in g(m, "GetBodies2", 0, False):
    for face in g(body, "GetFaces"):
        surface = g(face, "GetSurface")
        row = {"box": [v*1000 for v in g(face, "GetBox")]}
        for kind in ["Cylinder", "Cone", "Plane", "Torus"]:
            if g(surface, "Is"+kind):
                row[kind] = list(g(surface, kind+"Params"))
        faces.append(row)
def near(a, b):
    return abs(a-b) < 1e-5
def cyl(r, x, z, y0, y1):
    return [f for f in faces if "Cylinder" in f and near(f["Cylinder"][6]*1000, r)
        and near(f["Cylinder"][0]*1000, x) and near(f["Cylinder"][2]*1000, z)
        and near(f["box"][1], y0) and near(f["box"][4], y1)]
def dims(name):
    feature = m.FeatureByName(name)
    d = g(feature, "GetFirstDisplayDimension")
    result = []
    while d is not None:
        dim = g(d, "GetDimension2", 0)
        result.append({"value": g(dim, "SystemValue"), "driven": g(dim, "DrivenState")})
        d = g(feature, "GetNextDisplayDimension", d)
    return result
profile, outline = dims("轴套截面"), dims("双耳相切轮廓")
cones = [f for f in faces if "Cone" in f]
checks = {
    "单实体无特征错误警告": len(g(m, "GetBodies2", 0, False)) == 1 and not health["faulty_features"] and not health["warning_features"],
    "两张手工草图完全定义无Fix无孤立": len(quality["sketches"]) == 2 and quality["status"] == "pass",
    "轴套D34长25": len(cyl(17, 0, 0, 0, 25)) == 1,
    "D22孔从入口2至总长40": len(cyl(11, 0, 0, 2, 40)) == 1,
    "入口只一个锥面轴向2且相对轴60度": len(cones) == 1 and near(cones[0]["box"][1], 0)
        and near(cones[0]["box"][4], 2) and near(abs(cones[0]["Cone"][7]), math.pi/3),
    "原图2与端面30度为驱动": all(any(near(d["value"], v) and d["driven"] == 2 for d in profile) for v in [.002, math.pi/6]),
    "双耳中心距70与R12为驱动": all(any(near(d["value"], v) and d["driven"] == 2 for d in outline) for v in [.070, .012]),
    "两个耳部R12位置及15厚度": len(cyl(12, 0, 35, 25, 40)) == 1 and len(cyl(12, 0, -35, 25, 40)) == 1,
    "两个D12安装孔位于耳心贯穿15": len(cyl(6, 0, 35, 25, 40)) == 1 and len(cyl(6, 0, -35, 25, 40)) == 1,
    "中央外形D44两侧弧面": len(cyl(22, 0, 0, 25, 40)) == 2,
    "仅轴套回转与双耳拉伸两个成形特征": sum(f["type"] in ["Revolution", "Boss", "Extrusion", "ICE"] for f in health["features"]) == 2
        and near(g(g(m.FeatureByName("双耳法兰"), "GetDefinition"), "GetDepth", True), .015),
    "无自行增加圆角曲面": not any("Torus" in f for f in faces),
}
save_preview(m, str(out/"填料压盖_等轴测.bmp"), "isometric")
m.ShowNamedView2("*Top", 5)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"填料压盖_双耳端视.bmp"), 1600, 1000)
m.ShowNamedView2("*Front", 1)
m.ViewDisplayWireframe()
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"填料压盖_线框侧视.bmp"), 1600, 1000)
m.ViewDisplayShaded()
m.ShowNamedView2("*Isometric", 7)
m.GraphicsRedraw2()
report = {"started": started, "finished": datetime.now().astimezone().isoformat(),
          "seconds": time.perf_counter()-timer, "passed": all(checks.values()), "checks": checks,
          "quality": quality, "health": health, "faces": faces,
          "profile_dimensions": profile, "outline_dimensions": outline,
          "limits": "待目视对照；未制作工程图/装配，STEP未重导入；本轮未做改参实验。"}
(out/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"passed": report["passed"], "failed": [k for k, v in checks.items() if not v],
                  "seconds": report["seconds"]}, ensure_ascii=False))
assert report["passed"]
