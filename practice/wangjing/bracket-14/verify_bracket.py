"""LJT05.14：重开后一次集中核验，原图参数与实际孔面交叉检查。"""
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
target = out/"支架_LJT05.14.SLDPRT"
s = win32com.client.Dispatch("SldWorks.Application.32")
m = g(s, "ActiveDoc")
assert g(s, "GetProcessID") == 12356 and Path(g(m, "GetPathName")).resolve() == target
started = datetime.now().astimezone().isoformat()
timer = time.perf_counter()
e = win32com.client.VARIANT(16387, 0)
w = win32com.client.VARIANT(16387, 0)
assert m.Save3(1, e, w) and e.value == 0
s.CloseDoc(g(m, "GetTitle"))
assert g(s, "GetFirstDocument") is None
m = s.OpenDoc6(str(target), 1, 1, "", e, w)
assert m is not None and e.value == 0 and m.ForceRebuild3(False)
quality = inspect_model_sketches(m)
health = collect_model_summary(m)
surfaces = []
for body in g(m, "GetBodies2", 0, False):
    for face in g(body, "GetFaces"):
        surface = g(face, "GetSurface")
        row = {"box": [v*1000 for v in g(face, "GetBox")]}
        for kind in ["Cylinder", "Torus", "Cone", "Plane"]:
            if g(surface, "Is"+kind):
                row[kind] = list(g(surface, kind+"Params"))
        surfaces.append(row)
n = (.5, math.sqrt(3)/2, 0)
u = (math.sqrt(3)/2, -.5, 0)
def near(a, b):
    return abs(a-b) < 1e-5
def axial(radius, x0, x1, y=0, z=0):
    return [f for f in surfaces if "Cylinder" in f
        and near(f["Cylinder"][6]*1000, radius)
        and near(abs(f["Cylinder"][3]), 1)
        and near(f["Cylinder"][1]*1000, y) and near(f["Cylinder"][2]*1000, z)
        and near(f["box"][0], x0) and near(f["box"][3], x1)]
def feature_dims(name):
    feature = m.FeatureByName(name)
    dd = g(feature, "GetFirstDisplayDimension")
    result = []
    while dd is not None:
        d = g(dd, "GetDimension2", 0)
        result.append({"name": g(d, "FullName"), "value": g(d, "SystemValue"), "driven": g(d, "DrivenState")})
        dd = g(feature, "GetNextDisplayDimension", dd)
    return result
base_dims = feature_dims("底板截面")
checks = {
    "一个实体无特征错误警告": len(g(m, "GetBodies2", 0, False)) == 1 and not health["faulty_features"] and not health["warning_features"],
    "七张草图完全定义无Fix无孤立": len(quality["sketches"]) == 7 and quality["status"] == "pass",
    "底板45长7厚60度为驱动": all(any(near(d["value"], value) and d["driven"] == 2 for d in base_dims)
        for value in [.045, .007, math.pi/3]),
    "底板中心轴向位置15与高度52": any(near(d["value"], .052) for d in base_dims)
        and any(near(d["value"], .015) for d in base_dims),
    "底板宽30中面拉伸": near(g(g(m.FeatureByName("倾斜底板"), "GetDefinition"), "GetDepth", True), .030),
    "纵筋8宽横筋25宽": near(g(g(m.FeatureByName("纵筋"), "GetDefinition"), "GetDepth", True), .008)
        and near(g(g(m.FeatureByName("横筋"), "GetDefinition"), "GetDepth", True), .025),
    "左右D10孔实际贯通且含C1": len(axial(5, 1, 7.5)) == 1 and len(axial(5, 22.5, 29)) == 1,
    "中央D14孔含两端R2": len(axial(7, 9.5, 20.5)) == 1,
    "偏置D5盲孔从左端到8且距轴9": len(axial(2.5, 0, 8, 0, -9)) == 1,
    "盲孔口无自行添加倒角": m.FeatureByName("盲孔口C1") is None,
    "图示各圆角倒角特征持久存在": all(m.FeatureByName(name) is not None for name in
        ["底板四角R5", "十字筋内角R2", "筋脚R2", "套筒筋根R2", "内孔肩部R2", "图示C1", "左端外缘C1"]),
}
corners = [f for f in surfaces if "Cylinder" in f and near(f["Cylinder"][6], .005)
    and abs(sum(a*b for a, b in zip(f["Cylinder"][3:6], n))) > .999999]
checks["底板四角实际R5"] = len(corners) == 4
plate_holes = []
for face in surfaces:
    if "Cylinder" not in face:
        continue
    p = face["Cylinder"]
    if not near(p[6], .0025) or abs(sum(a*b for a, b in zip(p[3:6], n))) < .999999:
        continue
    p = [v*1000 for v in p[:3]]
    h = (p[0]-15)*n[0]+(p[1]+52)*n[1]
    q = [p[i]-h*n[i] for i in range(3)]
    plate_holes.append({"t": (q[0]-15)*u[0]+(q[1]+52)*u[1], "z": q[2]})
checks["四个D5孔按35乘20分布且方向垂直斜底板"] = len(plate_holes) == 4 and all(
    sum(near(p["t"], t) and near(p["z"], z) for p in plate_holes) == 1
    for t in [-17.5, 17.5] for z in [-10, 10])
save_preview(m, str(out/"支架_等轴测.bmp"), "isometric")
m.ShowNamedView2("*Left", 3)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"支架_B向.bmp"), 1600, 1000)
m.ShowNamedView2("*Front", 1)
m.ViewDisplayWireframe()
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"支架_线框正视.bmp"), 1600, 1000)
m.ViewDisplayShaded()
m.ShowNamedView2("*Isometric", 7)
m.GraphicsRedraw2()
report = {"started": started, "finished": datetime.now().astimezone().isoformat(),
          "seconds": time.perf_counter()-timer, "passed": all(checks.values()),
          "checks": checks, "quality": quality, "health": health, "surfaces": surfaces,
          "base_dimensions": base_dims, "plate_holes": plate_holes,
          "limits": "待目视对照；未制作工程图、装配，STEP未重导入；本轮未做改参实验。"}
(out/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"passed": report["passed"], "failed": [k for k, v in checks.items() if not v],
                  "seconds": report["seconds"]}, ensure_ascii=False))
assert report["passed"]
