"""LJT03.04滑阀：回转含盲孔、凸十字、两扁位、原生装饰螺纹。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member, mm, sketch, sketch_line, revolve_boss, new_document, save_document, dimension_existing_relations, round_edges
from sw_part import extrude_cut
from sw_connect import create_empty_dispatch_variant, open_document
from sw_sketch_quality import inspect_model_sketches
from sw_review import collect_model_summary, run_review
from sw_assembly import iter_feature_tree
from sw_edge_select import _edge_curve
from sw_export import export_to_step
import win32com.client

out = Path(sys.argv[1]).resolve()
path = out/"滑阀_LJT03.04.SLDPRT"
assert out.is_dir() and not path.exists()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == int(sys.argv[2])
sw.Visible = True
sw.UserControl = True
started = time.monotonic()
m = new_document(sw, "part")

def closed_lines(points, origin=False):
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        segments = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                    for a, b in zip(points, points[1:]+points[:1])]
    finally:
        sm.AddToDB = old
    assert all(s is not None for s in segments)
    sd = member(m.SelectionManager, "CreateSelectData")
    for a, b in zip(segments, segments[1:]+segments[:1]):
        m.ClearSelection2(True)
        assert member(a, "GetEndPoint2").Select4(False, sd)
        assert member(b, "GetStartPoint2").Select4(True, sd)
        m.SketchAddConstraints("sgCOINCIDENT")
    for segment, a, b in zip(segments, points, points[1:]+points[:1]):
        if a[0] == b[0] or a[1] == b[1]:
            m.ClearSelection2(True)
            assert segment.Select4(False, sd)
            m.SketchAddConstraints("sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    if origin:
        m.ClearSelection2(True)
        assert member(segments[0], "GetStartPoint2").Select4(False, sd)
        assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
        m.SketchAddConstraints("sgCOINCIDENT")
    return segments

def check_points(groups):
    for segments, points in groups:
        for segment, expected in zip(segments, points):
            p = member(segment, "GetStartPoint2")
            assert math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) < 1e-5

tip_y = 45-4.25/math.tan(math.pi/3)
profile = [(0, 0), (10, 0), (10, 12), (8, 12), (8, 15), (12, 15),
           (14, 17), (14, 21), (8, 21), (8, 59), (6, 61), (4.25, 61),
           (4.25, 45), (0, tip_y)]
with sketch(m, "Front Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0, 0, 0, 0, mm(61), 0)
    finally:
        sm.AddToDB = old
    assert axis is not None
    segments = closed_lines(profile, origin=True)
    sd = member(m.SelectionManager, "CreateSelectData")
    m.ClearSelection2(True)
    assert axis.Select4(False, sd)
    m.SketchAddConstraints("sgVERTICAL")
    m.ClearSelection2(True)
    assert member(axis, "GetStartPoint2").Select4(False, sd)
    assert member(segments[0], "GetStartPoint2").Select4(True, sd)
    m.SketchAddConstraints("sgCOINCIDENT")
    dimension_existing_relations(m)
    check_points([(segments, profile)])
boss = revolve_boss(m, ref, 2*math.pi)
assert boss is not None
boss.Name = "主体回转"
m.FeatureByName(str(ref)).Name = "轴向截面"
assert m.ForceRebuild3(False) and save_document(m, str(path))
m.ShowNamedView2("*Isometric", 7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
print("[主体] 回转完成，含两处C2、深16底孔与图示钻尖", flush=True)

# 端视中的十字为实体：四个角区去料，轴心与四臂保留。
with sketch(m, "Top Plane") as ref:
    groups = []
    for sx, sz in [(1, 1), (-1, 1), (-1, -1), (1, -1)]:
        points = [(sx*2, sz*2), (sx*12, sz*2), (sx*12, sz*12), (sx*2, sz*12)]
        groups.append((closed_lines(points), points))
    dimension_existing_relations(m)
    check_points(groups)
    active = member(m.SketchManager, "ActiveSketch")
    normal_y = list(member(member(member(active, "ModelToSketchTransform"), "Inverse"), "ArrayData"))[7]
cut = extrude_cut(m, ref, mm(15), direction=normal_y > 0)
assert cut is not None
cut.Name = "十字头铣削"
m.FeatureByName(str(ref)).Name = "四角去料草图"

with sketch(m, "Front Plane") as ref:
    groups = []
    for left, right in [(6, 20), (-20, -6)]:
        points = [(left, 24), (right, 24), (right, 29), (left, 29)]
        groups.append((closed_lines(points), points))
    dimension_existing_relations(m)
    check_points(groups)
cut = extrude_cut(m, ref, 0, both_directions=True)
assert cut is not None
cut.Name = "扳手平面"
m.FeatureByName(str(ref)).Name = "扳手平面草图"

mouth = [e for e, p, r in round_edges(m) if math.dist(p, [0, 61, 0]) < 1e-5 and abs(r-4.25) < 1e-5]
assert len(mouth) == 1
m.ClearSelection2(True)
assert mouth[0].Select2(False, 0)
thread = m.FeatureManager.InsertCosmeticThread3(8, "Tapped Hole", "M10x1.5", mm(10), 0, mm(14), "M10-6H 深14")
assert thread is not None
thread.Name = "M10螺纹"
m.ClearSelection2(True)
for k, v in [("名称", "滑阀"), ("图号", "LJT03.04"), ("材料", "45"),
             ("热处理", "调质40～45HRC"), ("配合", "Φ20f7；Φ16f7；M10-6H"),
             ("孔深", "M10×1.5有效深14；Φ8.5圆柱底孔深16；图示钻尖按120°假设"),
             ("尺寸说明", "凸十字臂宽4，四角铣至Y15法兰端面；跨12两平面Y24..29；两处C2按原图")]:
    m.Extension.CustomPropertyManager("").Add3(k, 30, v, 2)
assert m.ForceRebuild3(False) and save_document(m)
sw.CloseDoc(member(m, "GetTitle"))
m = open_document(sw, str(path), silent=True)
assert m.ForceRebuild3(False)
bodies = list(member(m, "GetBodies2", 0, False))
planes, cylinders, circles, axial_vertices, cones = [], [], [], [], []
for face in member(bodies[0], "GetFaces"):
    s = member(face, "GetSurface")
    if member(s, "IsPlane"):
        ys = []
        for e in member(face, "GetEdges"):
            for accessor in ["GetStartVertex", "GetEndVertex"]:
                v = member(e, accessor)
                if v is not None:
                    ys.append(member(v, "GetPoint")[1]*1000)
        planes.append({"params": list(member(s, "PlaneParams")), "vertex_y_mm": ys,
                       "opposite_surface_normal": bool(member(face, "FaceInSurfaceSense"))})
    if member(s, "IsCylinder"):
        cylinders.append({"params": list(member(s, "CylinderParams")), "area_mm2": member(face, "GetArea")*1e6})
    if member(s, "IsCone"):
        cones.append({"params": list(member(s, "ConeParams")), "box": list(member(face, "GetBox"))})
for e in member(bodies[0], "GetEdges"):
    curve = _edge_curve(e)
    if curve is not None and member(curve, "IsCircle"):
        cp = list(member(curve, "CircleParams"))
        if abs(abs(cp[4])-1) < 1e-7:
            circles.append([cp[1]*1000, cp[6]*1000])
    for accessor in ["GetStartVertex", "GetEndVertex"]:
        v = member(e, accessor)
        if v is not None:
            p = list(member(v, "GetPoint"))
            if math.hypot(p[0], p[2]) < 1e-8:
                axial_vertices.append(p[1]*1000)

def edge(y, r):
    return any(math.dist(c, [y, r]) < 1e-5 for c in circles)
flat_faces = [p for p in planes if abs(abs(p["params"][0])-1) < 1e-7 and abs(abs(p["params"][3])*1000-6) < 1e-5]
cross_faces = [p for p in planes if (abs(abs(p["params"][0])-1) < 1e-7 and abs(abs(p["params"][3])*1000-2) < 1e-5)
               or (abs(abs(p["params"][2])-1) < 1e-7 and abs(abs(p["params"][5])*1000-2) < 1e-5)]
quality = inspect_model_sketches(m)
health = collect_model_summary(m)
threads = [member(f, "GetDefinition") for f, _ in iter_feature_tree(m) if member(f, "GetTypeName2") == "CosmeticThread"]
checks = {"单实体": len(bodies) == 1, "草图全约束无Fix": quality["status"] == "pass",
          "无错误警告": not health["faulty_features"] and not health["warning_features"],
          "总长61及末端C2": edge(0, 10) and edge(59, 8) and edge(61, 6),
          "Φ20长12及Φ16退刀槽长3": all(edge(y, r) for y, r in [(12, 10), (12, 8), (15, 8)]),
          "Φ28法兰长6与前缘C2": all(edge(y, r) for y, r in [(15, 12), (17, 14), (21, 14), (21, 8)]),
          "宽4凸十字到Y15端面": len(cross_faces) == 8 and all(abs(min(p["vertex_y_mm"])) < 1e-5 and abs(max(p["vertex_y_mm"])-15) < 1e-5 for p in cross_faces),
          "十字为实体而非槽的材料侧核验": len(cross_faces) == 8 and all(abs((1 if not p["opposite_surface_normal"] else -1)*sum(n*v for n,v in zip(p["params"][:3],p["params"][3:]))*1000-2) < 1e-5 for p in cross_faces),
          "跨12扁位始于24长5": len(flat_faces) == 2 and all(abs(min(p["vertex_y_mm"])-24) < 1e-5 and abs(max(p["vertex_y_mm"])-29) < 1e-5 for p in flat_faces),
          "Φ8.5圆柱底孔深16": edge(45, 4.25) and edge(61, 4.25) and any(abs(c["params"][6]*2000-8.5) < 1e-5 and abs(c["area_mm2"]-math.pi*8.5*16) < 1e-5 for c in cylinders),
          "M10原生装饰螺纹深14": len(threads) == 1 and abs(member(threads[0], "Diameter")-.01) < 1e-8 and abs(member(threads[0], "BlindDepth")-.014) < 1e-8,
          "钻尖120度表达": any(abs(c["params"][6]-.00425) < 1e-8 and abs(c["params"][7]-math.pi/3) < 1e-8 and abs(c["params"][1]-.045) < 1e-8 and abs(c["box"][1]*1000-tip_y) < 1e-5 for c in cones)}
assert all(checks.values()), checks
assert export_to_step(m, str(out/"滑阀_LJT03.04.step"))
review, _ = run_review(m, str(out), basename="滑阀", views=["isometric", "right", "bottom", "top"], expected_outputs=[str(path), str(out/"滑阀_LJT03.04.step")])
assert review["evaluation"]["status"] != "fail"
m.ShowNamedView2("*Isometric", 7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert save_document(m)
report = {"checks": checks, "planes": planes, "cylinders": cylinders, "circle_edges_y_r_mm": circles,
          "axial_vertices_mm": axial_vertices, "cones": cones, "sketches": quality, "native_health": health,
          "profile_mm": profile, "reopened": True, "execution_seconds": round(time.monotonic()-started, 3),
          "finished": datetime.now().astimezone().isoformat(),
          "representation": "原生装饰螺纹，非实体牙型；图示钻尖采用120°假设，16为圆柱孔深"}
(out/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"checks": checks, "execution_seconds": report["execution_seconds"]}, ensure_ascii=False))
