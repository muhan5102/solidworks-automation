"""LJT04.34：轴套回转与双耳相切轮廓；原图角度与中心距驱动。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, mm, sketch, sketch_line, new_document, save_document, dimension_existing_relations, offset_plane
from sw_connect import create_empty_dispatch_variant
from sw_part import revolve_boss, extrude_boss
from sw_review import save_preview

out = Path(sys.argv[1]).resolve()
target = out/"填料压盖_LJT04.34.SLDPRT"
assert out.is_dir() and not target.exists()
s = win32com.client.Dispatch("SldWorks.Application.32")
assert g(s, "GetProcessID") == 12356 and g(s, "GetFirstDocument") is None
s.Visible = True
m = new_document(s, "part")
start = time.perf_counter()

def rel(items, kind):
    m.ClearSelection2(True)
    data = g(m.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    m.SketchAddConstraints(kind)

def origin(p):
    m.ClearSelection2(True)
    assert p.Select4(False, g(m.SelectionManager, "CreateSelectData"))
    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
    m.SketchAddConstraints("sgCOINCIDENT")

def axis_zero(p, horizontal):
    m.ClearSelection2(True)
    assert p.Select4(False, g(m.SelectionManager, "CreateSelectData"))
    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
    m.SketchAddConstraints("sgHORIZONTALPOINTS2D" if horizontal else "sgVERTICALPOINTS2D")

def dimension(items, pos, value):
    m.ClearSelection2(True)
    data = g(m.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    d = m.AddDimension2(*pos)
    assert d is not None
    assert abs(g(g(d, "GetDimension2", 0), "SystemValue")-value) < 1e-8

def checkpoint(stage):
    assert m.ForceRebuild3(False) and save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    print(f"[{stage}] 已保存，脚本累计{time.perf_counter()-start:.2f}秒", flush=True)
    (out/"checkpoint.json").write_text(json.dumps({
        "stage": stage, "saved": datetime.now().astimezone().isoformat(), "file": target.name
    }, ensure_ascii=False, indent=2), encoding="utf-8")

old_prompt = s.GetUserPreferenceToggle(10)
s.SetUserPreferenceToggle(10, False)
try:
    # 图示角30°相对端面；轴向2意味着径向展开2/tan30°。
    mouth = 11+2*math.sqrt(3)
    points = [(mouth, 0), (17, 0), (17, 25), (11, 25), (11, 2)]
    with sketch(m, "Front Plane") as ref:
        sm = m.SketchManager
        old = sm.AddToDB
        sm.AddToDB = True
        try:
            axis = sm.CreateCenterLine(0, 0, 0, 0, .025, 0)
            end_face = sm.CreateCenterLine(mm(mouth), 0, 0, mm(mouth)-.010, 0, 0)
            ee = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                  for a, b in zip(points, points[1:]+points[:1])]
        finally:
            sm.AddToDB = old
        for a, b in zip(ee, ee[1:]+ee[:1]):
            rel([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
        for e, a, b in zip(ee, points, points[1:]+points[:1]):
            if a[0] == b[0] or a[1] == b[1]:
                rel([e], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
        origin(g(axis, "GetStartPoint2"))
        rel([axis], "sgVERTICAL")
        axis_zero(g(ee[0], "GetStartPoint2"), True)
        rel([g(axis, "GetEndPoint2"), g(ee[2], "GetStartPoint2")], "sgHORIZONTALPOINTS2D")
        rel([end_face], "sgHORIZONTAL")
        rel([g(end_face, "GetStartPoint2"), g(ee[0], "GetStartPoint2")], "sgCOINCIDENT")
        dimension([ee[4], end_face], (.009, .002, 0), math.pi/6)
        # 直接用原图轴向2与30°，不另加换算斜边长度。
        m.ClearSelection2(True)
        data = g(m.SelectionManager, "CreateSelectData")
        assert g(ee[4], "GetStartPoint2").Select4(False, data)
        assert g(ee[4], "GetEndPoint2").Select4(True, data)
        dd = m.AddVerticalDimension2(.020, .001, 0)
        assert dd is not None
        d = g(dd, "GetDimension2", 0)
        assert abs(g(d, "SystemValue")-.002) < 1e-8
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name = "轴套截面"
    f = revolve_boss(m, "轴套截面", 2*math.pi)
    assert f is not None
    f.Name = "轴套回转"
    checkpoint("轴套与30度入口")

    plane = offset_plane(m, "Top Plane", 25, "法兰截面")
    c, h = math.sqrt(45)/7, 2/7
    pts = [(-12*c, 35+12*h), (12*c, 35+12*h),
           (22*c, 22*h), (22*c, -22*h), (12*c, -35-12*h),
           (-12*c, -35-12*h), (-22*c, -22*h), (-22*c, 22*h)]
    with sketch(m, plane) as ref:
        sm = m.SketchManager
        inv = list(g(g(g(sm, "ActiveSketch"), "ModelToSketchTransform"), "Inverse").ArrayData)
        old = sm.AddToDB
        sm.AddToDB = True
        try:
            ee = []
            centers = {0: (0, 35), 2: (0, 0), 4: (0, -35), 6: (0, 0)}
            for i, (a, b) in enumerate(zip(pts, pts[1:]+pts[:1])):
                if i in centers:
                    center = centers[i]
                    e = sm.CreateArc(mm(center[0]), mm(center[1]), 0, mm(a[0]), mm(a[1]), 0,
                                     mm(b[0]), mm(b[1]), 0, -1)
                else:
                    e = sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                assert e is not None
                ee.append(e)
            arm_top = sm.CreateCenterLine(0, 0, 0, 0, .035, 0)
            arm_bottom = sm.CreateCenterLine(0, 0, 0, 0, -.035, 0)
            holes = [sm.CreateCircleByRadius(0, mm(y), 0, mm(r)) for y, r in [(0, 11), (35, 6), (-35, 6)]]
        finally:
            sm.AddToDB = old
        for a, b in zip(ee, ee[1:]+ee[:1]):
            rel([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
            rel([a, b], "sgTANGENT")
        origin(g(ee[2], "GetCenterPoint2"))
        origin(g(ee[6], "GetCenterPoint2"))
        for arm, arc in [(arm_top, ee[0]), (arm_bottom, ee[4])]:
            origin(g(arm, "GetStartPoint2"))
            rel([arm], "sgVERTICAL")
            rel([g(arm, "GetEndPoint2"), g(arc, "GetCenterPoint2")], "sgCOINCIDENT")
        rel([arm_top, arm_bottom], "sgSAMELENGTH")
        rel([ee[0], ee[4]], "sgSAMELENGTH")
        rel([ee[2], ee[6]], "sgSAMELENGTH")
        origin(g(holes[0], "GetCenterPoint2"))
        rel([g(holes[1], "GetCenterPoint2"), g(ee[0], "GetCenterPoint2")], "sgCOINCIDENT")
        rel([g(holes[2], "GetCenterPoint2"), g(ee[4], "GetCenterPoint2")], "sgCOINCIDENT")
        rel([holes[1], holes[2]], "sgSAMELENGTH")
        dimension([g(ee[0], "GetCenterPoint2"), g(ee[4], "GetCenterPoint2")], (.032, 0, 0), .070)
        dimension([ee[0]], (.017, .045, 0), .012)
        dimension([ee[2]], (.028, 0, 0), .022)
        dimension_existing_relations(m)
        for edge, expected in zip(ee, pts):
            p = g(edge, "GetStartPoint2")
            assert math.dist([g(p, "X")*1000, g(p, "Y")*1000], expected) < 1e-5
    m.FeatureByName(str(ref)).Name = "双耳相切轮廓"
    f = extrude_boss(m, "双耳相切轮廓", .015, direction=inv[7] < 0)
    assert f is not None
    f.Name = "双耳法兰"
    m.ClearSelection2(True)
    assert m.FeatureByName("法兰截面").Select2(False, 0)
    m.BlankRefGeom()
    m.SetMaterialPropertyName2("", "", "")
    for key, value in [("名称", "填料压盖"), ("图号", "LJT04.34"), ("材料", "HT150"),
                       ("技术要求", "锐边倒钝；未注尺寸公差GB/T 1804—2000-m；未注几何公差GB/T 1184—1996-K")]:
        m.Extension.CustomPropertyManager("").Add3(key, 30, value, 2)
    checkpoint("双耳法兰及两安装孔")
    save_preview(m, str(out/"填料压盖_等轴测.bmp"), "isometric")
finally:
    s.SetUserPreferenceToggle(10, old_prompt)
(out/"build.json").write_text(json.dumps({
    "finished": datetime.now().astimezone().isoformat(), "execution_seconds": time.perf_counter()-start
}, ensure_ascii=False, indent=2), encoding="utf-8")
