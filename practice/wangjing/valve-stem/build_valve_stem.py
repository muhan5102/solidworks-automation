"""王静图集LJT03.03阀杆：回转、45°四平面、两端平面。"""
import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "year08" / "connector"))
from build_connector import member, mm, sketch, sketch_line, new_document, save_document, revolve_boss, dimension_existing_relations
from sw_connect import open_document, create_empty_dispatch_variant
from sw_part import extrude_cut
from sw_edge_select import _edge_curve
from sw_sketch_quality import inspect_model_sketches
from sw_review import save_preview, collect_model_summary
from sw_export import export_to_step
import win32com.client

parser = argparse.ArgumentParser()
parser.add_argument("output", type=Path)
parser.add_argument("--process-id", type=int, required=True)
args = parser.parse_args()
out = args.output.resolve()
path = out / "阀杆_LJT03.03.SLDPRT"
assert out.is_dir() and not path.exists()
started = time.monotonic()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == args.process_id
sw.Visible = True
sw.UserControl = True
m = new_document(sw, "part")
cone_y = 1.5*math.tan(math.pi/6)
sphere_y = 30+math.sqrt(20**2-9**2)
points = [(0, 0), (5.5, 0), (7, cone_y), (7, 38), (9, 38), (9, sphere_y)]

def constrain_chain(entities):
    sd = member(m.SelectionManager, "CreateSelectData")
    for a, b in zip(entities, entities[1:]+entities[:1]):
        m.ClearSelection2(True)
        assert member(a, "GetEndPoint2").Select4(False, sd)
        assert member(b, "GetStartPoint2").Select4(True, sd)
        m.SketchAddConstraints("sgCOINCIDENT")
    return sd

def line_relation(line, kind, sd):
    m.ClearSelection2(True)
    assert line.Select4(False, sd)
    m.SketchAddConstraints(kind)

with sketch(m, "Front Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0, 0, 0, 0, mm(50), 0)
        lines = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1])) for a, b in zip(points, points[1:])]
        arc = sm.CreateArc(0, mm(30), 0, mm(9), mm(sphere_y), 0, 0, mm(50), 0, 1)
        closing = sketch_line(m, 0, mm(50), 0, 0)
    finally:
        sm.AddToDB = old
    entities = lines+[arc, closing]
    assert all(e is not None for e in entities) and axis is not None
    sd = constrain_chain(entities)
    for line, relation in [(lines[0], "sgHORIZONTAL"), (lines[2], "sgVERTICAL"),
                           (lines[3], "sgHORIZONTAL"), (lines[4], "sgVERTICAL"),
                           (closing, "sgVERTICAL"), (axis, "sgVERTICAL")]:
        line_relation(line, relation, sd)
    for p in [member(axis, "GetStartPoint2"), member(lines[0], "GetStartPoint2")]:
        m.ClearSelection2(True)
        assert p.Select4(False, sd)
        assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
        m.SketchAddConstraints("sgCOINCIDENT")
    m.ClearSelection2(True)
    assert member(arc, "GetCenterPoint2").Select4(False, sd)
    assert axis.Select4(True, sd)
    m.SketchAddConstraints("sgCOINCIDENT")
    dimension_existing_relations(m)
    assert abs(member(arc, "GetRadius")-mm(20)) < 1e-9
    for line, expected in zip(lines, points):
        p = member(line, "GetStartPoint2")
        assert math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) < 1e-5
feature = revolve_boss(m, ref, 2*math.pi)
assert feature is not None
feature.Name = "主体回转"
m.FeatureByName(str(ref)).Name = "轴向截面"
assert m.ForceRebuild3(False) and save_document(m, str(path))
m.ShowNamedView2("*Isometric", 7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
print("[主体] Φ14轴、Φ18头、30°端部及SR20已回转", flush=True)

# 四个不相交的三角形只裁去圆截面在45°法向上的超出部分。
a = 11/math.sqrt(2)
with sketch(m, "Top Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    groups = []
    try:
        for sx, sy in [(1, 1), (-1, 1), (-1, -1), (1, -1)]:
            pts = [(sx*.5, sy*(a-.5)), (sx*(a-.5), sy*.5), (sx*(a-.5), sy*(a-.5))]
            entities = [sketch_line(m, mm(p[0]), mm(p[1]), mm(q[0]), mm(q[1])) for p, q in zip(pts, pts[1:]+pts[:1])]
            groups.append((entities, pts))
    finally:
        sm.AddToDB = old
    for entities, pts in groups:
        assert all(e is not None for e in entities)
        sd = constrain_chain(entities)
        line_relation(entities[1], "sgVERTICAL", sd)
        line_relation(entities[2], "sgHORIZONTAL", sd)
    dimension_existing_relations(m)
    for entities, pts in groups:
        for line, expected in zip(entities, pts):
            p = member(line, "GetStartPoint2")
            assert math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) < 1e-5
    active = member(sm, "ActiveSketch")
    normal_y = list(member(member(member(active, "ModelToSketchTransform"), "Inverse"), "ArrayData"))[7]
cut = extrude_cut(m, ref, mm(14), direction=normal_y > 0)
assert cut is not None
cut.Name = "四方铣平"
m.FeatureByName(str(ref)).Name = "四方截面"

with sketch(m, "Right Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    groups = []
    try:
        for left, right in [(4.25, 20), (-20, -4.25)]:
            pts = [(left, 43), (right, 43), (right, 51), (left, 51)]
            entities = [sketch_line(m, mm(p[0]), mm(p[1]), mm(q[0]), mm(q[1])) for p, q in zip(pts, pts[1:]+pts[:1])]
            groups.append((entities, pts))
    finally:
        sm.AddToDB = old
    for entities, pts in groups:
        sd = constrain_chain(entities)
        for line, p, q in zip(entities, pts, pts[1:]+pts[:1]):
            line_relation(line, "sgVERTICAL" if p[0] == q[0] else "sgHORIZONTAL", sd)
    dimension_existing_relations(m)
    for entities, pts in groups:
        for line, expected in zip(entities, pts):
            p = member(line, "GetStartPoint2")
            assert math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) < 1e-5
cut = extrude_cut(m, ref, 0, both_directions=True)
assert cut is not None
cut.Name = "端头扁位"
m.FeatureByName(str(ref)).Name = "端头扁位草图"
for key, value in [("名称", "阀杆"), ("图号", "LJT03.03"), ("材料", "40Cr"),
                   ("技术要求", "调质220～250HBW；去毛刺锐边；按图公称尺寸"),
                   ("配合", "Φ14d11；Φ18a11；左端11×11；右端8.5两扁位")]:
    m.Extension.CustomPropertyManager("").Add3(key, 30, value, 2)
assert m.ForceRebuild3(False) and save_document(m)
sw.CloseDoc(member(m, "GetTitle"))
m = open_document(sw, str(path), silent=True)
assert m.ForceRebuild3(False)
bodies = list(member(m, "GetBodies2", 0, False))
planes, spheres, cylinders, circles = [], [], [], []
for f in member(bodies[0], "GetFaces"):
    s = member(f, "GetSurface")
    if member(s, "IsPlane"):
        params = list(member(s, "PlaneParams"))
        ys = []
        for e in member(f, "GetEdges"):
            for accessor in ["GetStartVertex", "GetEndVertex"]:
                v = member(e, accessor)
                if v is not None:
                    ys.append(member(v, "GetPoint")[1]*1000)
        planes.append({"params": params, "vertex_y_mm": ys})
    if member(s, "IsSphere"):
        spheres.append(list(member(s, "SphereParams")))
    if member(s, "IsCylinder"):
        cylinders.append(list(member(s, "CylinderParams")))
for e in member(bodies[0], "GetEdges"):
    curve = _edge_curve(e)
    if curve is not None and member(curve, "IsCircle"):
        p = list(member(curve, "CircleParams"))
        if abs(abs(p[4])-1) < 1e-7:
            circles.append([p[1]*1000, p[6]*1000])
quality = inspect_model_sketches(m)
health = collect_model_summary(m)
square_faces = [p for p in planes if abs(abs(p["params"][0])-1/math.sqrt(2)) < 1e-7
                and abs(abs(p["params"][2])-1/math.sqrt(2)) < 1e-7
                and abs(sum(n*v for n, v in zip(p["params"][:3], p["params"][3:]))*1000) > 5.49]
head_faces = [p for p in planes if abs(abs(p["params"][2])-1) < 1e-7
              and abs(abs(p["params"][5])*1000-4.25) < 1e-5]
def edge_at(y, r):
    return any(math.dist(c, [y, r]) < 1e-5 for c in circles)
checks = {"单实体": len(bodies) == 1, "草图全约束无Fix": quality["status"] == "pass",
          "无错误警告": not health["faulty_features"] and not health["warning_features"],
          "SR20中心Y30及总长50": len(spheres) == 1 and math.dist(spheres[0], [0, .03, 0, .02]) < 1e-8,
          "Φ14及Φ18": all(any(abs(c[6]*2000-d) < 1e-5 for c in cylinders) for d in [14, 18]),
          "Φ11端及30度锥段": edge_at(0, 5.5) and edge_at(cone_y, 7),
          "四平面相对主视转45度且跨距11长14": len(square_faces) == 4 and all(
              abs(abs(sum(n*v for n, v in zip(p["params"][:3], p["params"][3:])))*1000-5.5) < 1e-5
              and abs(max(p["vertex_y_mm"])-14) < 1e-5 for p in square_faces),
          "端头两平面跨距8.5始于Y43": len(head_faces) == 2 and all(abs(min(p["vertex_y_mm"])-43) < 1e-5 for p in head_faces),
          "头部台阶始于Y38": edge_at(38, 7) and edge_at(38, 9)}
assert all(checks.values()), checks
assert export_to_step(m, str(out/"阀杆_LJT03.03.step"))
for view in ["isometric", "right", "top"]:
    save_preview(m, str(out/("阀杆_"+view+".bmp")), view)
m.ShowNamedView2("*Isometric", 7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert save_document(m)
report = {"checks": checks, "planes": planes, "spheres": spheres, "cylinders": cylinders,
          "circular_edges_y_r_mm": circles, "sketches": quality, "native_health": health,
          "reopened": True, "execution_seconds": round(time.monotonic()-started, 3),
          "finished": datetime.now().astimezone().isoformat(),
          "scope": "原生公称几何，材料热处理和公差仅属性；未出工程图"}
(out/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"checks": checks, "execution_seconds": report["execution_seconds"]}, ensure_ascii=False))
