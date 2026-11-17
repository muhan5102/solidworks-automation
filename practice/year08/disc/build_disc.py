"""第8届配气盘SDF-09：薄盘、四通孔、槽面方向按原图。单位输入mm。"""
import sys, math, json, time, argparse
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import (member, mm, sketch, sketch_circle, sketch_line,
                             new_document, save_document, circle_cut,
                             round_edges)
from sw_part import extrude_boss, extrude_cut, fillet
from sw_sketch_quality import fully_define_sketch, inspect_model_sketches
from sw_connect import open_document
from sw_review import collect_geometry_measurements, save_preview
import win32com.client

parser = argparse.ArgumentParser()
parser.add_argument("output", type=Path)
parser.add_argument("--process-id", required=True, type=int)
parser.add_argument("--name", default="配气盘.SLDPRT")
args = parser.parse_args()
out = args.output.resolve()
assert Path(args.name).name == args.name and Path(args.name).suffix.lower() == ".sldprt"
dest = out / args.name
assert out.is_dir() and not dest.exists()
started = time.monotonic()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == args.process_id
sw.Visible = True
sw.UserControl = True
model = new_document(sw, "part")

with sketch(model, "Top Plane") as ref:
    assert sketch_circle(model, 0, 0, mm(17)) is not None
    active = member(model.SketchManager, "ActiveSketch")
    normal_y = list(member(member(member(active, "ModelToSketchTransform"), "Inverse"), "ArrayData"))[7]
    fully_define_sketch(model)
boss = extrude_boss(model, ref, mm(4), direction=(normal_y < 0))
assert boss is not None
boss.Name = "圆盘主体"
member(model, "FeatureByName", str(ref)).Name = "外圆草图"
model.ForceRebuild3(False)
assert save_document(model, str(dest))
model.ShowNamedView2("*Isometric", 7)
model.ViewZoomtofit2()
model.GraphicsRedraw2()

# 正视图：上、右为Φ4通孔；左上和右下为Φ3H9孔。
# D向视图四孔均可见，未标孔深，按通孔处理，不擅自设盲深。
d = 11 / math.sqrt(2)
circle_cut(model, sw, "Top Plane", [(0, 0, 11), (11, 0, 0)],
           2, 0, 1, "通孔")
circle_cut(model, sw, "Top Plane", [(-d, 0, d), (d, 0, -d)],
           1.5, 0, 1, "定位孔")
# From the slot side, +X is right and +Z is up: D4 holes must be up/right.
# The solid occupies Y=0..4, so that drawing side is Y=0, not Y=4.
back = "Top Plane"
outer = [edge for edge, pos, radius in round_edges(model)
         if math.dist(pos, [0, 0, 0]) < .001 and abs(radius - 17) < .001]
assert len(outer) == 1
model.ClearSelection2(True)
assert outer[0].Select2(False, 0)
rounding = fillet(model, mm(1))
assert rounding is not None
rounding.Name = "外缘圆角"

# 槽总长14、宽5，两个半圆中心间距9；A-A给定背面深2。
with sketch(model, back) as ref:
    sm = model.SketchManager
    old_add = sm.AddToDB
    sm.AddToDB = True
    try:
        left = sketch_line(model, mm(-2.5), mm(-4.5), mm(-2.5), mm(4.5))
        top = sm.CreateArc(0, mm(4.5), 0, mm(-2.5), mm(4.5), 0,
                           mm(2.5), mm(4.5), 0, -1)
        right = sketch_line(model, mm(2.5), mm(4.5), mm(2.5), mm(-4.5))
        bottom = sm.CreateArc(0, mm(-4.5), 0, mm(2.5), mm(-4.5), 0,
                              mm(-2.5), mm(-4.5), 0, -1)
    finally:
        sm.AddToDB = old_add
    segments = [left, top, right, bottom]
    assert all(s is not None for s in segments)
    sd = member(model.SelectionManager, "CreateSelectData")
    def relation(a, b, kind):
        model.ClearSelection2(True)
        assert a.Select4(False, sd) and b.Select4(True, sd)
        model.SketchAddConstraints(kind)
    for a, b in zip(segments, segments[1:] + segments[:1]):
        relation(member(a, "GetEndPoint2"), member(b, "GetStartPoint2"), "sgCOINCIDENT")
        relation(a, b, "sgTANGENT")
    for line in (left, right):
        model.ClearSelection2(True)
        assert line.Select4(False, sd)
        model.SketchAddConstraints("sgVERTICAL")
    fully_define_sketch(model)
    expected_points = {(-2.5, -4.5), (-2.5, 4.5), (2.5, -4.5), (2.5, 4.5)}
    for segment in segments:
        for point in (member(segment, "GetStartPoint2"), member(segment, "GetEndPoint2")):
            xy = (member(point, "X") * 1000, member(point, "Y") * 1000)
            assert min(math.dist(xy, expected) for expected in expected_points) < 1e-5
    active = member(sm, "ActiveSketch")
    assert all(member(c, "IsClosed") for c in member(active, "GetSketchContours"))
    normal_y = list(member(member(member(active, "ModelToSketchTransform"), "Inverse"), "ArrayData"))[7]
slot = extrude_cut(model, ref, mm(2), direction=(normal_y > 0))
assert slot is not None
slot.Name = "背面长圆槽"
member(model, "FeatureByName", str(ref)).Name = "长圆槽草图"
model.SetMaterialPropertyName2("", "", "")
for key, value in [("名称", "配气盘"), ("图号", "SDF-09"), ("材料", "LY12"),
                   ("技术要求", "未注圆角R1；氧化处理"),
                   ("配合", "2×Φ3H9；槽宽5H9，公称尺寸建模")]:
    model.Extension.CustomPropertyManager("").Add3(key, 30, value, 2)
assert model.ForceRebuild3(False)
assert save_document(model, str(dest))
sw.CloseDoc(member(model, "GetTitle"))
model = open_document(sw, str(dest), silent=True)
assert model.ForceRebuild3(False)
geo = collect_geometry_measurements(model)
quality = inspect_model_sketches(model)
checks = {"one_solid": len(member(model, "GetBodies2", 0, False)) == 1,
          "fully_defined": quality["status"] == "pass"}
for diameter, centers in [(4, [(0, 11), (11, 0)]),
                          (3, [(-d, d), (d, -d)])]:
    faces = [f for f in geo["cylindrical_faces"]
             if f["internal"] and abs(f["diameter_mm"] - diameter) < 1e-5]
    checks[f"D{diameter}_positions"] = (
        len(faces) == 2 and all(any(math.dist(
            [f["origin_mm"][0], f["origin_mm"][2]], center) < 1e-4
            and abs(f["axial_length_mm"] - 4) < 1e-4 for f in faces) for center in centers))
slot_bounds, torus_radii, plane_positions = [], [], []
for body in member(model, "GetBodies2", 0, False):
    for face in member(body, "GetFaces"):
        surface = member(face, "GetSurface")
        if member(surface, "IsTorus"):
            torus_radii.append(list(member(surface, "TorusParams"))[7])
        if member(surface, "IsPlane"):
            plane_positions.append(list(member(surface, "PlaneParams")))
        if not member(surface, "IsCylinder"):
            continue
        cylinder = list(member(surface, "CylinderParams"))
        if abs(cylinder[6] - .0025) > 1e-8:
            continue
        points = []
        for edge in member(face, "GetEdges"):
            for accessor in ("GetStartVertex", "GetEndVertex"):
                vertex = member(edge, accessor)
                if vertex is not None:
                    points.append(list(member(vertex, "GetPoint")))
        slot_bounds.append({"center_xz": [cylinder[0], cylinder[2]],
                            "y_bounds": [min(p[1] for p in points), max(p[1] for p in points)]})
checks["slot_width5_length14_depth2"] = len(slot_bounds) == 2 and all(
    any(math.dist(b["center_xz"], [0, z]) < 1e-8 and
        math.dist(b["y_bounds"], [0.0, .002]) < 1e-8 for b in slot_bounds)
    for z in [-.0045, .0045])
checks["slot_floor_y2"] = any(abs(abs(p[1]) - 1) < 1e-7 and abs(p[4] - .002) < 1e-8
                               for p in plane_positions)
checks["outer_R1"] = len(torus_radii) == 1 and abs(torus_radii[0] - .001) < 1e-8
assert all(checks.values()), checks
if member(model, "GetVisibilityOfConstructPlanes"):
    member(model, "ViewDispRefplanes")
model.ClearSelection2(True)
save_preview(model, str(out / "配气盘_isometric.bmp"), "isometric")
save_preview(model, str(out / "配气盘_槽面.bmp"), "bottom")
assert save_document(model)
model.ShowNamedView2("*Isometric", 7)
model.ViewZoomtofit2()
model.GraphicsRedraw2()
report = {"checks": checks, "geometry": geo, "sketches": quality,
          "execution_seconds": round(time.monotonic() - started, 3),
          "end": datetime.now().astimezone().isoformat(),
          "slot_bounds_m": slot_bounds,
          "pending": "材料、公差及表面处理仅属性，未出工程图"}
(out / "verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"checks": checks, "execution_seconds": report["execution_seconds"]}, ensure_ascii=False))
