"""LJT05.14阶段续作：图示R2过渡、轴向台阶孔、底板孔及偏置盲孔。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, mm, sketch, sketch_line, dimension_existing_relations, save_document, chamfer_edges, round_edges
from sw_connect import create_empty_dispatch_variant
from sw_part import fillet, extrude_cut, _ensure_sketch_selected
from sw_sketch_quality import fully_define_sketch
from sw_review import save_preview

out = Path(sys.argv[1]).resolve()
target = out / "支架_LJT05.14.SLDPRT"
s = win32com.client.Dispatch("SldWorks.Application.32")
m = g(s, "ActiveDoc")
assert g(s, "GetProcessID") == 12356 and Path(g(m, "GetPathName")).resolve() == target
start = time.perf_counter()
history = []
n = (.5, math.sqrt(3)/2, 0)
u = (math.sqrt(3)/2, -.5, 0)

def rel(items, kind):
    m.ClearSelection2(True)
    data = g(m.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    m.SketchAddConstraints(kind)

def edges():
    rows, seen = [], set()
    for body in g(m, "GetBodies2", 0, False):
        for face in g(body, "GetFaces"):
            for edge in g(face, "GetEdges"):
                va, vb = g(edge, "GetStartVertex"), g(edge, "GetEndVertex")
                if va is None or vb is None:
                    continue
                a, b = [v*1000 for v in g(va, "GetPoint")], [v*1000 for v in g(vb, "GetPoint")]
                key = tuple(sorted((tuple(round(v, 7) for v in a), tuple(round(v, 7) for v in b))))
                if key in seen:
                    continue
                seen.add(key)
                rows.append((edge, a, b))
    return rows

def round_selection(name, rows, expected):
    assert len(rows) == expected, (name, len(rows), expected)
    m.ClearSelection2(True)
    f = fillet(m, .002, [row[0] for row in rows])
    assert f is not None, name
    f.Name = name
    assert m.ForceRebuild3(False)
    history.append({"name": name, "radius_mm": 2, "edges": [{"a": a, "b": b} for _, a, b in rows]})
    print(f"[过渡] {name}，{len(rows)}条语义选边", flush=True)

def checkpoint(stage):
    assert m.ForceRebuild3(False)
    assert save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    (out / "checkpoint.json").write_text(json.dumps({
        "stage": stage, "saved": datetime.now().astimezone().isoformat(), "file": target.name
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[{stage}] 已保存，续作累计{time.perf_counter()-start:.2f}秒", flush=True)

def on_top(p):
    return abs((p[0]-15)*n[0]+(p[1]+52)*n[1]) < 1e-5

if m.FeatureByName("十字筋内角R2") is None:
    rows = [row for row in edges() if all(any(abs(p[0]-x) < 1e-5 for x in [11, 19])
            and abs(abs(p[2])-4) < 1e-5 for p in row[1:])
            and abs(row[1][1]-row[2][1]) > 20]
    round_selection("十字筋内角R2", rows, 4)
if m.FeatureByName("筋脚R2") is None:
    rows = [row for row in edges() if all(on_top(p) and abs((p[0]-15)/u[0]) < 17
            and abs(p[2]) <= 12.50001 for p in row[1:])]
    round_selection("筋脚R2", rows, 16)
if m.FeatureByName("套筒筋根R2") is None:
    candidates, seen = [], set()
    for body in g(m, "GetBodies2", 0, False):
        for face in g(body, "GetFaces"):
            surface = g(face, "GetSurface")
            if not g(surface, "IsCylinder") or abs(g(surface, "CylinderParams")[6]-.0125) > 1e-8:
                continue
            for edge in g(face, "GetEdges"):
                va, vb = g(edge, "GetStartVertex"), g(edge, "GetEndVertex")
                if not va or not vb:
                    continue
                a, b = [v*1000 for v in g(va, "GetPoint")], [v*1000 for v in g(vb, "GetPoint")]
                key = tuple(sorted((tuple(round(v, 7) for v in a), tuple(round(v, 7) for v in b))))
                if key in seen:
                    continue
                seen.add(key)
                if all(.01 < p[0] < 29.99 and p[1] <= 1e-5 for p in [a, b]) and min(a[1], b[1]) < -1:
                    candidates.append((edge, a, b))
    round_selection("套筒筋根R2", candidates, 14)
checkpoint("图示筋过渡")

if m.FeatureByName("轴向台阶孔") is None:
    with sketch(m, "Front Plane") as ref:
        points = [(0, 0), (30, 0), (30, 5), (22.5, 5), (22.5, 7), (7.5, 7), (7.5, 5), (0, 5)]
        sm = m.SketchManager
        old = sm.AddToDB
        sm.AddToDB = True
        try:
            ee = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                  for a, b in zip(points, points[1:] + points[:1])]
            axis = sm.CreateCenterLine(0, 0, 0, .030, 0, 0)
        finally:
            sm.AddToDB = old
        for a, b in zip(ee, ee[1:]+ee[:1]):
            rel([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
        for edge, a, b in zip(ee, points, points[1:]+points[:1]):
            rel([edge], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
        m.ClearSelection2(True)
        assert g(ee[0], "GetStartPoint2").Select4(False, g(m.SelectionManager, "CreateSelectData"))
        assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
        m.SketchAddConstraints("sgCOINCIDENT")
        rel([g(axis, "GetStartPoint2"), g(ee[0], "GetStartPoint2")], "sgCOINCIDENT")
        rel([g(axis, "GetEndPoint2"), g(ee[0], "GetEndPoint2")], "sgCOINCIDENT")
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name = "内孔截面"
    _ensure_sketch_selected(m, "内孔截面")
    data = g(m.SelectionManager, "CreateSelectData")
    data.Mark = 16
    assert axis.Select4(True, data)
    f = m.FeatureManager.FeatureRevolve2(True, True, False, True, False, False, 0, 0,
        2*math.pi, 0, False, False, 0, 0, 0, 0, 0, True, False, True)
    assert f is not None
    f.Name = "轴向台阶孔"
    chamfer_edges(m, [([0, 0, 0], 5), ([30, 0, 0], 5), ([30, 0, 0], 12.5)], 1, "图示C1")

if m.FeatureByName("内孔肩部R2") is None:
    selected = [edge for edge, pos, radius in round_edges(m) if abs(radius-7) < 1e-5
                and any(math.dist(pos, [x, 0, 0]) < 1e-5 for x in [7.5, 22.5])]
    assert len(selected) == 2
    m.ClearSelection2(True)
    f = fillet(m, .002, selected)
    assert f is not None
    f.Name = "内孔肩部R2"
chamfer_edges(m, [([0, 0, 0], 12.5)], 1, "左端外缘C1")

def face_circles(name, ray_start, ray_direction, centers, radius, depth):
    m.ClearSelection2(True)
    assert m.Extension.SelectByRay(*[mm(v) for v in ray_start], *ray_direction, .0001, 2, False, 0, 0)
    m.SketchManager.InsertSketch(True)
    active = g(m.SketchManager, "ActiveSketch")
    assert active is not None
    sketch_name = g(active, "Name")
    transform = list(g(g(active, "ModelToSketchTransform"), "ArrayData"))
    try:
        sm = m.SketchManager
        old = sm.AddToDB
        sm.AddToDB = True
        try:
            circles = []
            for center in centers:
                p = [mm(v) for v in center]
                xyz = [sum(p[j]*transform[j*3+i] for j in range(3))*transform[12]+transform[9+i] for i in range(3)]
                assert abs(xyz[2]) < 1e-7
                circle = sm.CreateCircleByRadius(xyz[0], xyz[1], 0, mm(radius))
                assert circle is not None
                circles.append(circle)
        finally:
            sm.AddToDB = old
        # 首先补精确零坐标关系，避免自动补尺寸遗漏。
        for circle in circles:
            p = g(circle, "GetCenterPoint2")
            coords = [g(p, "X"), g(p, "Y")]
            for index, kind in [(0, "sgVERTICALPOINTS2D"), (1, "sgHORIZONTALPOINTS2D")]:
                if abs(coords[index]) < 1e-8:
                    m.ClearSelection2(True)
                    assert p.Select4(False, g(m.SelectionManager, "CreateSelectData"))
                    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
                    m.SketchAddConstraints(kind)
        fully_define_sketch(m)
    finally:
        m.SketchManager.InsertSketch(True)
    feature = m.FeatureByName(sketch_name)
    assert feature is not None
    feature.Name = name+"孔位"
    f = extrude_cut(m, feature.Name, mm(depth) if depth else 0, direction=False)
    assert f is not None
    f.Name = name

if m.FeatureByName("底板四安装孔") is None:
    centers = [(15+t*u[0], -52+t*u[1], z) for t in [-17.5, 17.5] for z in [-10, 10]]
    ray_start = tuple(centers[0][i]+10*n[i] for i in range(3))
    face_circles("底板四安装孔", ray_start, tuple(-v for v in n), centers, 2.5, 7)
if m.FeatureByName("偏置盲孔") is None:
    face_circles("偏置盲孔", (-10, 0, -9), (1, 0, 0), [(0, 0, -9)], 2.5, 8)
m.SetMaterialPropertyName2("", "", "")
for key, value in [("名称", "支架"), ("图号", "LJT05.14"), ("材料", "HT150"),
                   ("技术要求", "铸件需经时效处理；未注圆角R2～R3；未注倒角C1；未注表面粗糙度Ra12.5")]:
    m.Extension.CustomPropertyManager("").Add3(key, 30, value, 2)
checkpoint("全部孔系完成")
save_preview(m, str(out / "支架_等轴测.bmp"), "isometric")
(out / "finish.json").write_text(json.dumps({
    "finished": datetime.now().astimezone().isoformat(), "execution_seconds": time.perf_counter()-start,
    "fillets": history
}, ensure_ascii=False, indent=2), encoding="utf-8")
