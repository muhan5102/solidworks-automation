"""LJT04.33透盖：一张轴向截面，孔向导母孔加六孔阵列。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client
import win32gui
import win32process

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, mm, sketch, sketch_line, new_document, save_document, dimension_existing_relations
from sw_connect import create_empty_dispatch_variant
from sw_part import revolve_boss, _select_first_candidate, _get_plane_name_candidates

out = Path(sys.argv[1]).resolve()
target = out / "透盖_LJT04.33.SLDPRT"
assert out.is_dir() and not target.exists()
s = win32com.client.Dispatch("SldWorks.Application.32")
pid = g(s, "GetProcessID")
assert pid == 12356
window = 69324
assert win32gui.IsWindow(window) and win32process.GetWindowThreadProcessId(window)[1] == pid
win32gui.ShowWindow(window, 9)
assert win32gui.IsWindowVisible(window) and not win32gui.IsIconic(window)
assert g(s, "GetFirstDocument") is None
(out / "session.json").write_text(json.dumps({"process_id": pid, "window_visible": True}), encoding="utf-8")
s.Visible = True
m = new_document(s, "part")
start = time.perf_counter()

def relation(items, kind):
    m.ClearSelection2(True)
    data = g(m.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    m.SketchAddConstraints(kind)

def origin(point):
    m.ClearSelection2(True)
    assert point.Select4(False, g(m.SelectionManager, "CreateSelectData"))
    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
    m.SketchAddConstraints("sgCOINCIDENT")

def progress(stage):
    assert m.ForceRebuild3(False)
    assert save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    print(f"[{stage}] 已保存，脚本累计{time.perf_counter()-start:.2f}秒", flush=True)
    (out / "checkpoint.json").write_text(json.dumps({
        "stage": stage, "saved": datetime.now().astimezone().isoformat(),
        "file": target.name, "next": "孔群或重开终检"
    }, ensure_ascii=False, indent=2), encoding="utf-8")

# 轴向Y：平端0、法兰后面13、凸台后面18。
# 槽根4居中于13厚度，起点(13-4)/2=4.5；槽口5.5，同中心。
points = [(17.5, 0), (60, 0), (60, 13), (40, 13), (40, 18),
          (34, 18), (34, 13), (17.5, 13), (17.5, 9.25),
          (24, 8.5), (24, 4.5), (17.5, 3.75)]
with sketch(m, "Front Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0, 0, 0, 0, mm(18), 0)
        edges = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                 for a, b in zip(points, points[1:] + points[:1])]
    finally:
        sm.AddToDB = old
    for a, b in zip(edges, edges[1:] + edges[:1]):
        relation([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
    for edge, a, b in zip(edges, points, points[1:] + points[:1]):
        if a[0] == b[0] or a[1] == b[1]:
            relation([edge], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    origin(g(axis, "GetStartPoint2"))
    relation([axis], "sgVERTICAL")
    relation([g(axis, "GetStartPoint2"), g(edges[0], "GetStartPoint2")], "sgHORIZONTALPOINTS2D")
    relation([g(axis, "GetEndPoint2"), g(edges[4], "GetStartPoint2")], "sgHORIZONTALPOINTS2D")
    relation([edges[8], edges[10]], "sgSAMELENGTH")
    old_prompt = s.GetUserPreferenceToggle(10)
    s.SetUserPreferenceToggle(10, False)
    try:
        data = g(m.SelectionManager, "CreateSelectData")
        m.ClearSelection2(True)
        assert edges[9].Select4(False, data)
        dd = m.AddDimension2(.028, .0065, 0)
        assert dd is not None and abs(g(g(dd, "GetDimension2", 0), "SystemValue") - .004) < 1e-8
        m.ClearSelection2(True)
        assert g(edges[8], "GetStartPoint2").Select4(False, data)
        assert g(edges[10], "GetEndPoint2").Select4(True, data)
        dd = m.AddVerticalDimension2(.012, .0065, 0)
        assert dd is not None and abs(g(g(dd, "GetDimension2", 0), "SystemValue") - .0055) < 1e-8
        dimension_existing_relations(m)
    finally:
        s.SetUserPreferenceToggle(10, old_prompt)
    for e, p in zip(edges, points):
        q = g(e, "GetStartPoint2")
        assert math.dist([g(q, "X") * 1000, g(q, "Y") * 1000], p) < 1e-5
m.FeatureByName(str(ref)).Name = "轴向截面"
f = revolve_boss(m, "轴向截面", 2 * math.pi)
assert f is not None
f.Name = "主体回转"
progress("主体及梯形内环槽")

m.ClearSelection2(True)
assert m.Extension.SelectByRay(0, -.010, .049, 0, 1, 0, .0001, 2, False, 0, 0)
f = m.FeatureManager.HoleWizard5(0, 8, 139, "M8", 1, .009, .013, -1,
    .016, .006, 0, -1, -1, 0, 0, 0, 0, 0, 0, -1,
    "", False, False, True, False, False, False)
assert f is not None
f.Name = "安装沉孔"
sf = g(f, "GetFirstSubFeature")
sf.Name = "安装孔位"
g(sf, "GetNextSubFeature").Name = "沉孔向导轮廓"
point = g(g(sf, "GetSpecificFeature2"), "GetSketchPoints2")[0]
m.ClearSelection2(True)
assert sf.Select2(False, 0)
m.EditSketch()
# 首孔在主轴上，仅需主轴关系与49半径驱动，不新增45度定位线。
x, y = g(point, "X"), g(point, "Y")
assert abs(x) < 1e-8 and abs(abs(y) - .049) < 1e-8
m.ClearSelection2(True)
assert point.Select4(False, g(m.SelectionManager, "CreateSelectData"))
assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
m.SketchAddConstraints("sgVERTICALPOINTS2D")
old_prompt = s.GetUserPreferenceToggle(10)
s.SetUserPreferenceToggle(10, False)
try:
    m.ClearSelection2(True)
    data = g(m.SelectionManager, "CreateSelectData")
    assert point.Select4(False, data)
    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
    dd = m.AddDimension2(.012, 0, .025)
    assert dd is not None
    assert abs(g(g(dd, "GetDimension2", 0), "SystemValue") - .049) < 1e-8
    assert g(g(sm, "ActiveSketch"), "GetConstrainedStatus") == 3
finally:
    s.SetUserPreferenceToggle(10, old_prompt)
    sm.InsertSketch(True)
m.ClearSelection2(True)
assert _select_first_candidate(m.Extension, _get_plane_name_candidates("Front Plane"), "PLANE")
assert _select_first_candidate(m.Extension, _get_plane_name_candidates("Right Plane"), "PLANE", append=True)
assert m.InsertAxis2(True)
axis = g(m, "FirstFeature")
while axis is not None:
    if g(axis, "GetTypeName2") == "RefAxis":
        axis.Name = "回转轴"
    axis = g(axis, "GetNextFeature")
m.ClearSelection2(True)
assert f.Select2(False, 4)
assert m.Extension.SelectByID2("回转轴", "AXIS", 0, 0, 0, True, 1, create_empty_dispatch_variant(), 0)
p = m.FeatureManager.FeatureCircularPattern4(6, 2 * math.pi, False, "", False, True, False)
assert p is not None
p.Name = "六组安装孔阵列"
m.ClearSelection2(True)
assert m.FeatureByName("回转轴").Select2(False, 0)
m.BlankRefGeom()
m.SetMaterialPropertyName2("", "", "")
for name, value in [("名称", "透盖"), ("图号", "LJT04.33"), ("材料", "HT150"),
                    ("技术要求", "去毛刺、飞边；未注尺寸公差GB/T 1804—2000-c；未注几何公差GB/T 1184—1996-L")]:
    m.Extension.CustomPropertyManager("").Add3(name, 30, value, 2)
progress("六组沉孔完成")
(out / "build.json").write_text(json.dumps({
    "finished": datetime.now().astimezone().isoformat(), "execution_seconds": time.perf_counter() - start
}, ensure_ascii=False, indent=2), encoding="utf-8")
