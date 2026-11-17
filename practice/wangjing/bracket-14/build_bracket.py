"""LJT05.14：斜底板、十字筋和套筒；先成形再统一开孔。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, mm, sketch, sketch_line, new_document, save_document, dimension_existing_relations
from sw_connect import create_empty_dispatch_variant
from sw_part import extrude_midplane, revolve_boss, fillet
from sw_review import save_preview
from sw_edge_select import _edge_curve

out = Path(sys.argv[1]).resolve()
target = out / "支架_LJT05.14.SLDPRT"
resume = "--resume" in sys.argv[2:]
assert out.is_dir() and (target.exists() if resume else not target.exists())
s = win32com.client.Dispatch("SldWorks.Application.32")
assert g(s, "GetProcessID") == 12356
assert resume or g(s, "GetFirstDocument") is None
s.Visible = True
m = g(s, "ActiveDoc") if resume else new_document(s, "part")
if resume:
    assert Path(g(m, "GetPathName")).resolve() == target
start = time.perf_counter()
u = (math.sqrt(3) / 2, -.5, 0)
n = (.5, math.sqrt(3) / 2, 0)

def rel(items, kind):
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

def polygon(points):
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        edges = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                 for a, b in zip(points, points[1:] + points[:1])]
    finally:
        sm.AddToDB = old
    for a, b in zip(edges, edges[1:] + edges[:1]):
        rel([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
    for e, a, b in zip(edges, points, points[1:] + points[:1]):
        if a[0] == b[0] or a[1] == b[1]:
            rel([e], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    return edges

def dimension(items, text, expected, vertical=False):
    m.ClearSelection2(True)
    data = g(m.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    d = (m.AddVerticalDimension2 if vertical else m.AddDimension2)(*text)
    assert d is not None
    assert abs(g(g(d, "GetDimension2", 0), "SystemValue") - expected) < 1e-8

def checkpoint(stage):
    assert m.ForceRebuild3(False)
    assert save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    print(f"[{stage}] 已保存，脚本累计{time.perf_counter()-start:.2f}秒", flush=True)
    (out / "checkpoint.json").write_text(json.dumps({
        "stage": stage, "saved": datetime.now().astimezone().isoformat(), "file": target.name
    }, ensure_ascii=False, indent=2), encoding="utf-8")

old_prompt = s.GetUserPreferenceToggle(10)
s.SetUserPreferenceToggle(10, False)
try:
    if not resume:
        # 主轴X；左端X0、右端X30。底板上表面中心(15,-52,0)。
        a = (15 - 22.5*u[0], -52 - 22.5*u[1])
        b = (15 + 22.5*u[0], -52 + 22.5*u[1])
        points = [a, b, (b[0]-7*n[0], b[1]-7*n[1]), (a[0]-7*n[0], a[1]-7*n[1])]
        with sketch(m, "Front Plane") as ref:
            edges = polygon(points)
            rel([edges[0], edges[2]], "sgPARALLEL")
            rel([edges[1], edges[3]], "sgPARALLEL")
            rel([edges[0], edges[1]], "sgPERPENDICULAR")
            sm = m.SketchManager
            axis = sm.CreateCenterLine(mm(15), 0, 0, mm(15), mm(-52), 0)
            rel([axis], "sgVERTICAL")
            m.ClearSelection2(True)
            assert g(axis, "GetStartPoint2").Select4(False, g(m.SelectionManager, "CreateSelectData"))
            assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
            m.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            # 定义60°原图角度：底板左向射线与竖直向上射线。
            ray = sm.CreateCenterLine(mm(15), mm(-52), 0, mm(a[0]), mm(a[1]), 0)
            rel([g(ray, "GetStartPoint2"), g(axis, "GetEndPoint2")], "sgCOINCIDENT")
            rel([g(ray, "GetEndPoint2"), g(edges[0], "GetStartPoint2")], "sgCOINCIDENT")
            rel([ray, edges[0]], "sgCOLINEAR")
            dimension([axis, ray], (.007, -.037, 0), math.pi/3)
            dimension([edges[0]], (.015, -.043, 0), .045)
            dimension([edges[1]], (.042, -.063, 0), .007)
            dimension([axis], (.041, -.026, 0), .052)
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name = "底板截面"
        f = extrude_midplane(m, "底板截面", .030)
        assert f is not None
        f.Name = "倾斜底板"
        # R5仅选四条板厚方向直边，排除长边。
        selected = []
        seen = set()
        for body in g(m, "GetBodies2", 0, False):
            for face in g(body, "GetFaces"):
                for edge in g(face, "GetEdges"):
                    curve = _edge_curve(edge)
                    if not g(curve, "IsLine"):
                        continue
                    p, q = g(edge, "GetStartVertex"), g(edge, "GetEndVertex")
                    if p is None or q is None:
                        continue
                    p, q = list(g(p, "GetPoint")), list(g(q, "GetPoint"))
                    key = tuple(sorted((tuple(round(v, 9) for v in p), tuple(round(v, 9) for v in q))))
                    if key in seen:
                        continue
                    seen.add(key)
                    delta = [(q[i]-p[i])*1000 for i in range(3)]
                    if abs(math.dist(p, q)*1000-7) < 1e-5 and abs(delta[0]*n[1]-delta[1]*n[0]) < 1e-5:
                        selected.append(edge)
        assert len(selected) == 4
        m.ClearSelection2(True)
        f = fillet(m, .005, selected)
        assert f is not None
        f.Name = "底板四角R5"
        checkpoint("底板")

        # 套筒先为实心；筋与套筒合并后再统一切内孔，避免筋填堵孔腔。
        with sketch(m, "Front Plane") as ref:
            edges = polygon([(0, 0), (30, 0), (30, 12.5), (0, 12.5)])
            origin(g(edges[0], "GetStartPoint2"))
            axis = m.SketchManager.CreateCenterLine(0, 0, 0, .030, 0, 0)
            origin(g(axis, "GetStartPoint2"))
            rel([axis], "sgHORIZONTAL")
            rel([g(axis, "GetEndPoint2"), g(edges[0], "GetEndPoint2")], "sgCOINCIDENT")
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name = "套筒截面"
        f = revolve_boss(m, "套筒截面", 2*math.pi)
        assert f is not None
        f.Name = "套筒"

    if resume:
        # 本轮实际中断点：首张筋截面未消费，只补缺失Y0关系。
        orphan = m.FeatureByName("草图3")
        assert orphan is not None
        m.ClearSelection2(True)
        assert orphan.Select2(False, 0)
        m.EditSketch()
        try:
            point = g(g(orphan, "GetSpecificFeature2"), "GetSketchPoints2")[0]
            m.ClearSelection2(True)
            assert point.Select4(False, g(m.SelectionManager, "CreateSelectData"))
            assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
            m.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            assert g(g(m.SketchManager, "ActiveSketch"), "GetConstrainedStatus") == 3
        finally:
            m.SketchManager.InsertSketch(True)
        orphan.Name = "纵筋截面"
        f = extrude_midplane(m, "纵筋截面", .008)
        assert f is not None
        f.Name = "纵筋"
    for name, x0, x1, width in [("纵筋", 2.5, 27.5, 8), ("横筋", 11, 19, 25)]:
        if m.FeatureByName(name) is not None:
            continue
        with sketch(m, "Front Plane") as ref:
            points = [(x0, 0), (x1, 0), (x1, -52-(x1-15)/math.sqrt(3)), (x0, -52-(x0-15)/math.sqrt(3))]
            edges = polygon(points)
            m.ClearSelection2(True)
            assert g(edges[0], "GetStartPoint2").Select4(False, g(m.SelectionManager, "CreateSelectData"))
            assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
            m.SketchAddConstraints("sgHORIZONTALPOINTS2D")
            # 草图底边与底板上表面一致，使用原生共线关系关联。
            m.ClearSelection2(True)
            assert edges[2].Select4(False, g(m.SelectionManager, "CreateSelectData"))
            assert m.Extension.SelectByID2("Line1@底板截面", "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
            m.SketchAddConstraints("sgCOLINEAR")
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name = name+"截面"
        f = extrude_midplane(m, name+"截面", mm(width))
        assert f is not None
        f.Name = name
    checkpoint("套筒与十字筋")
    save_preview(m, str(out / "结构阶段.bmp"), "isometric")
finally:
    s.SetUserPreferenceToggle(10, old_prompt)
(out / "build-structure.json").write_text(json.dumps({
    "finished": datetime.now().astimezone().isoformat(), "execution_seconds": time.perf_counter()-start
}, ensure_ascii=False, indent=2), encoding="utf-8")
