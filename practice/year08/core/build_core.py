"""第8届内芯练习：阶梯回转、两端扁位、盲孔和装饰螺纹。"""
import argparse
import json
import math
from pathlib import Path
import sys
import time
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import (
    member, mm, new_document, save_document, sketch, sketch_line,
    revolve_boss, circle_cut, offset_plane, round_edges,
    dimension_existing_relations,
)
from sw_part import extrude_cut
from sw_connect import open_document, create_empty_dispatch_variant
from sw_sketch_quality import inspect_model_sketches
from sw_review import collect_geometry_measurements, save_preview
from sw_assembly import iter_feature_tree
from finish_core import finish
import win32com.client

parser = argparse.ArgumentParser()
parser.add_argument("output", type=Path)
parser.add_argument("--process-id", required=True, type=int)
args = parser.parse_args()
out = args.output.resolve()
dest = out / "内芯.SLDPRT"
assert out.is_dir() and not dest.exists()
started = time.monotonic()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == args.process_id
sw.Visible = True
sw.UserControl = True
model = new_document(sw, "part")

def checkpoint(name):
    assert model.ForceRebuild3(False)
    assert save_document(model, str(dest))
    model.ClearSelection2(True)
    model.ShowNamedView2("*Isometric", 7)
    model.ViewZoomtofit2()
    model.GraphicsRedraw2()
    print(f"[{name}] 已保存，执行累计{time.monotonic()-started:.1f}秒", flush=True)

points = [(0, 0), (5, 0), (5, 2), (8.5, 2), (8.5, 4),
          (7, 4), (7, 7.5), (5.6, 7.5), (5.6, 9.5), (7, 9.5),
          (7, 13), (5, 13), (5, 17.8), (3.5, 17.8),
          (3.5, 19.8), (0, 19.8)]
with sketch(model, "Front Plane") as ref:
    sm = model.SketchManager
    old_add = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0, 0, 0, 0, mm(19.8), 0)
        segments = [sketch_line(model, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                    for a, b in zip(points, points[1:] + points[:1])]
    finally:
        sm.AddToDB = old_add
    assert axis is not None and all(s is not None for s in segments)
    sd = member(model.SelectionManager, "CreateSelectData")
    for a, b in zip(segments, segments[1:] + segments[:1]):
        model.ClearSelection2(True)
        assert member(a, "GetEndPoint2").Select4(False, sd)
        assert member(b, "GetStartPoint2").Select4(True, sd)
        model.SketchAddConstraints("sgCOINCIDENT")
    for segment, a, b in zip(segments, points, points[1:] + points[:1]):
        model.ClearSelection2(True)
        assert segment.Select4(False, sd)
        model.SketchAddConstraints("sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    model.ClearSelection2(True)
    assert axis.Select4(False, sd)
    model.SketchAddConstraints("sgVERTICAL")
    for point in (member(axis, "GetStartPoint2"), member(segments[0], "GetStartPoint2")):
        model.ClearSelection2(True)
        assert point.Select4(False, sd)
        assert model.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0,
                                           True, 0, create_empty_dispatch_variant(), 0)
        model.SketchAddConstraints("sgCOINCIDENT")
    dimension_existing_relations(model)
    for segment, expected in zip(segments, points):
        p = member(segment, "GetStartPoint2")
        assert math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) < 1e-5
    contours = list(member(member(sm, "ActiveSketch"), "GetSketchContours"))
    assert contours and all(member(c, "IsClosed") for c in contours)
body = revolve_boss(model, ref, 2*math.pi)
assert body is not None
body.Name = "主体回转"
member(model, "FeatureByName", str(ref)).Name = "轴向截面"
checkpoint("主体")

# 端视图按投影对应：短Φ10端宽5、Φ4深2；长Φ10轴肩宽8、Φ7端M5深7。
with sketch(model, "Front Plane") as ref:
    # 切除工具伸出前端1mm，避免零距离定位尺寸；不改变实体的0..2扁位。
    rectangles = [(2.5, -1, 12, 2), (-12, -1, -2.5, 2),
                  (4, 13, 12, 20.8), (-12, 13, -4, 20.8)]
    sm = model.SketchManager
    old_add = sm.AddToDB
    sm.AddToDB = True
    all_segments = []
    try:
        for x1, y1, x2, y2 in rectangles:
            corners = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
            lines = [sketch_line(model, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                     for a, b in zip(corners, corners[1:]+corners[:1])]
            assert all(line is not None for line in lines)
            all_segments.append((lines, corners))
    finally:
        sm.AddToDB = old_add
    sd = member(model.SelectionManager, "CreateSelectData")
    for lines, corners in all_segments:
        for a, b in zip(lines, lines[1:]+lines[:1]):
            model.ClearSelection2(True)
            assert member(a, "GetEndPoint2").Select4(False, sd)
            assert member(b, "GetStartPoint2").Select4(True, sd)
            model.SketchAddConstraints("sgCOINCIDENT")
        for line, a, b in zip(lines, corners, corners[1:]+corners[:1]):
            model.ClearSelection2(True)
            assert line.Select4(False, sd)
            model.SketchAddConstraints("sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    dimension_existing_relations(model)
    allowed = [(x, y) for x1, y1, x2, y2 in rectangles
               for x, y in [(x1, y1), (x1, y2), (x2, y1), (x2, y2)]]
    active = member(model.SketchManager, "ActiveSketch")
    for segment in member(active, "GetSketchSegments"):
        for p in (member(segment, "GetStartPoint2"), member(segment, "GetEndPoint2")):
            assert min(math.dist([member(p, "X")*1000, member(p, "Y")*1000], q)
                       for q in allowed) < 1e-5
cut = extrude_cut(model, ref, 0, both_directions=True)
assert cut is not None
cut.Name = "两端扁位"
member(model, "FeatureByName", str(ref)).Name = "扁位草图"

circle_cut(model, sw, "Top Plane", [(0, 0, 0)], 2, 2, 1, "短端盲孔")
back = offset_plane(model, "Top Plane", 19.8, "螺纹端基准面")
circle_cut(model, sw, back, [(0, 19.8, 0)], 2.1, 7, -1, "螺纹底孔")
mouth = [e for e, p, r in round_edges(model)
         if math.dist(p, [0, 19.8, 0]) < .001 and abs(r-2.1) < .001]
assert len(mouth) == 1
model.ClearSelection2(True)
assert mouth[0].Select2(False, 0)
thread = model.FeatureManager.InsertCosmeticThread3(
    8, "Tapped Hole", "M5x0.8", mm(5), 0, mm(7), "M5 深7")
assert thread is not None
thread.Name = "M5螺纹"
checkpoint("扁位与孔")
treatments = finish(model)
model.SetMaterialPropertyName2("", "", "")
for name, value in [("名称", "内芯"), ("图号", "SDF-05"), ("材料", "YL12"),
                    ("技术要求", "未注圆角R0.5；未注倒角C0.5；表面涂银粉漆"),
                    ("配合", "短Φ10端扁位5h9；长Φ10轴肩扁位8h9；轴颈Φ7h9、Φ14d9"),
                    ("假设", "M5按深7作简化螺纹孔，不另设底孔工艺深度；无钻尖及实体牙型")]:
    model.Extension.CustomPropertyManager("").Add3(name, 30, value, 2)
if member(model, "GetVisibilityOfConstructPlanes"):
    member(model, "ViewDispRefplanes")
checkpoint("收尾")
sw.CloseDoc(member(model, "GetTitle"))
model = open_document(sw, str(dest), silent=True)
assert model.ForceRebuild3(False)
geometry = collect_geometry_measurements(model)
quality = inspect_model_sketches(model)
plane_faces = []
for solid in member(model, "GetBodies2", 0, False):
    for face in member(solid, "GetFaces"):
        surface = member(face, "GetSurface")
        if member(surface, "IsPlane"):
            plane_faces.append(list(member(surface, "PlaneParams")))
checks = {"one_solid": len(member(model, "GetBodies2", 0, False)) == 1,
          "fully_defined": quality["status"] == "pass",
          "short_end_flat_5": all(any(abs(abs(p[0])-1) < 1e-7 and abs(p[3]-x) < 1e-8
                                     for p in plane_faces) for x in [-.0025, .0025]),
          "shoulder_flat_8": all(any(abs(abs(p[0])-1) < 1e-7 and abs(p[3]-x) < 1e-8
                                    for p in plane_faces) for x in [-.004, .004]),
          "rear_hole_D4": any(f["internal"] and abs(f["diameter_mm"]-4) < 1e-5
                              and abs(f["axial_length_mm"]-2) < 1e-5
                              for f in geometry["cylindrical_faces"]),
          "M5_native_thread": sum(member(f, "GetTypeName2") == "CosmeticThread"
                                  for f, _ in iter_feature_tree(model)) == 1}
save_preview(model, str(out/"内芯_等轴测.bmp"), "isometric")
save_preview(model, str(out/"内芯_正视.bmp"), "front")
model.ShowNamedView2("*Isometric", 7)
model.ViewZoomtofit2()
model.GraphicsRedraw2()
report = {"checks": checks, "geometry": geometry, "sketches": quality,
          "planes": plane_faces, "expected_profile_mm": points,
          "execution_seconds": round(time.monotonic()-started, 3),
          "end": datetime.now().astimezone().isoformat(),
          "treatments": treatments,
          "remaining": ["M5为深7简化螺纹表达，无钻尖及实体牙型",
                        "材料、公差、表面处理未正式落图"],
          "assumptions": []}
(out/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"checks": checks, "execution_seconds": report["execution_seconds"]}, ensure_ascii=False))
