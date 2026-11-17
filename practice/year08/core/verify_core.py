"""回读内芯最终原生几何，封闭圆边使用圆心而非不存在的端点。"""
import json
import math
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member, save_document
from sw_connect import open_document
from sw_sketch_quality import inspect_model_sketches
from sw_assembly import iter_feature_tree
from sw_review import save_preview, collect_model_summary, run_review
from sw_export import export_to_step
from sw_edge_select import _edge_curve
import win32com.client

out = Path(sys.argv[1]).resolve()
path = out / "内芯.SLDPRT"
sw = win32com.client.Dispatch("SldWorks.Application.32")
m = sw.ActiveDoc
assert Path(member(m, "GetPathName")).resolve() == path
assert save_document(m)
sw.CloseDoc(member(m, "GetTitle"))
m = open_document(sw, str(path), silent=True)
assert m.ForceRebuild3(False)
bodies = list(member(m, "GetBodies2", 0, False))
assert len(bodies) == 1
planes, cylinders, face_boundaries = [], [], []
for face in member(bodies[0], "GetFaces"):
    surf = member(face, "GetSurface")
    points, circles = [], []
    for edge in member(face, "GetEdges"):
        for accessor in ("GetStartVertex", "GetEndVertex"):
            vertex = member(edge, accessor)
            if vertex is not None:
                points.append([c * 1000 for c in member(vertex, "GetPoint")])
        curve = _edge_curve(edge)
        if curve is not None and member(curve, "IsCircle"):
            cp = list(member(curve, "CircleParams"))
            points.append([c * 1000 for c in cp[:3]])
            circles.append([round(cp[1]*1000, 6), round(cp[6]*1000, 6)])
    ys = [min(p[1] for p in points), max(p[1] for p in points)] if points else None
    if member(surf, "IsPlane"):
        planes.append({"params": list(member(surf, "PlaneParams")), "y_mm": ys})
    if member(surf, "IsCylinder"):
        cylinders.append({"diameter_mm": list(member(surf, "CylinderParams"))[6] * 2000,
                          "y_mm": ys, "area_mm2": member(face, "GetArea")*1e6})
    face_boundaries.append({"circle_y_radius_mm": sorted(set(map(tuple, circles))), "y_mm": ys})

def near(a, b):
    return abs(a - b) < 1e-5

def cylinder(diameter, start, end):
    return any(near(c["diameter_mm"], diameter) and c["y_mm"] and
               near(c["y_mm"][0], start) and near(c["y_mm"][1], end) for c in cylinders)

def flats(width, start, end):
    return all(any(abs(abs(p["params"][0]) - 1) < 1e-7 and
                   near(p["params"][3] * 1000, x) and p["y_mm"] and
                   near(p["y_mm"][0], start) and near(p["y_mm"][1], end)
                   for p in planes) for x in [-width / 2, width / 2])

def transition(boundaries):
    expected_edges = set(boundaries)
    return any(set(map(tuple, f["circle_y_radius_mm"])) == expected_edges for f in face_boundaries)

expected = [(0, 0), (5, 0), (5, 2), (8.5, 2), (8.5, 4), (7, 4), (7, 7.5),
            (5.6, 7.5), (5.6, 9.5), (7, 9.5), (7, 13), (5, 13), (5, 17.8),
            (3.5, 17.8), (3.5, 19.8), (0, 19.8)]
sketch = member(m.FeatureByName("轴向截面"), "GetSpecificFeature2")
profile = []
for segment in member(sketch, "GetSketchSegments"):
    if not member(segment, "ConstructionGeometry"):
        p = member(segment, "GetStartPoint2")
        profile.append(tuple(round(member(p, k) * 1000, 6) for k in ("X", "Y")))
threads = []
for feature, _ in iter_feature_tree(m):
    if member(feature, "GetTypeName2") == "CosmeticThread":
        data = member(feature, "GetDefinition")
        threads.append({"diameter_mm": member(data, "Diameter") * 1000,
                        "depth_mm": member(data, "BlindDepth") * 1000})
quality = inspect_model_sketches(m)
summary = collect_model_summary(m)
round_feature = m.FeatureByName("末端圆角")
chamfer_feature = m.FeatureByName("台阶倒角")
radius = member(member(round_feature, "GetDefinition"), "DefaultRadius") * 1000
chamfer_distance = member(m.Parameter("D1@台阶倒角"), "SystemValue") * 1000
chamfer_angle = member(m.Parameter("D2@台阶倒角"), "SystemValue")
checks = {
    "单实体": len(bodies) == 1,
    "主体16点及总长19.8": set(profile) == set(expected),
    "短Φ10端宽5扁位长2": cylinder(10, 0, 2) and flats(5, 0, 2),
    "Φ17凸缘及指定边C0.5": cylinder(17, 2, 3.5),
    "两段Φ14及指定边C0.5": cylinder(14, 4, 7) and cylinder(14, 9.5, 12.5),
    "Φ11.2槽长2": cylinder(11.2, 7.5, 9.5),
    "长Φ10轴肩宽8扁位长4.8": cylinder(10, 13, 17.8) and flats(8, 13, 17.8),
    "Φ7端保留完整圆柱及R0.5": cylinder(7, 17.8, 19.3) and near(radius, .5) and any(
        near(c["diameter_mm"], 7) and near(c["area_mm2"], 2*math.pi*3.5*1.5) for c in cylinders),
    "短Φ10端Φ4盲孔深2": cylinder(4, 0, 2),
    "Φ7端M5孔向内深7": cylinder(4.2, 12.8, 19.8),
    "原生M5螺纹深7": len(threads) == 1 and near(threads[0]["diameter_mm"], 5)
                       and near(threads[0]["depth_mm"], 7),
    "草图全约束无Fix": quality["status"] == "pass",
    "三处台阶C0.5乘45度": near(chamfer_distance, .5) and near(chamfer_angle, math.pi/4),
    "实体三处倒角边界各退0.5": all(transition(pair) for pair in [
        [(3.5, 8.5), (4, 8)], [(7, 7), (7.5, 6.5)], [(12.5, 7), (13, 6.5)]]),
    "实体末端圆角边界位置": transition([(19.3, 3.5), (19.8, 3)]),
    "无批量根部圆角及额外端缘倒角": all(m.FeatureByName(n) is None for n in
                              ["台阶圆角", "台阶根部圆角", "端缘倒角", "螺纹孔口倒角"]),
    "特征树无错误警告": not summary["faulty_features"] and not summary["warning_features"],
}
assert all(checks.values()), checks
m.ClearSelection2(True)
save_preview(m, str(out / "内芯_等轴测.bmp"), "isometric")
save_preview(m, str(out / "内芯_正视.bmp"), "front")
save_preview(m, str(out / "内芯_题面主视方向.bmp"), "right")
save_preview(m, str(out / "内芯_端视.bmp"), "top")
m.ShowNamedView2("*Isometric", 7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
assert save_document(m)
assert export_to_step(m, str(out / "内芯.step"))
review, review_path = run_review(m, str(out), basename="内芯",
                                views=["isometric", "right", "top"],
                                expected_outputs=[str(path), str(out / "内芯.step")])
report = {"status": "按两端投影对应修正扁位及孔端归属，核实圆倒角实体位置", "checks": checks,
          "profile_mm": profile, "native_cylinders": cylinders, "threads": threads,
          "sketches": quality, "reopen_verified": True,
          "features": summary, "radius_mm": radius,
          "chamfer_mm": chamfer_distance, "chamfer_angle_rad": chamfer_angle,
          "transition_face_boundaries": face_boundaries,
          "source_mapping": {"short_D10_y0_end": "题面右侧端视：5h9、Φ4深2",
                             "D7_y19_8_end": "题面左侧端视：Φ10轴肩8h9、末端M5深7",
                             "side_view_confirmation": "宽8扁位在Φ10轴肩的交线距轴线3mm，与主视图对应"},
          "review_report": str(Path(review_path).name),
          "representation": "M5原生装饰螺纹；不生成实体牙型、额外钻尖或独立工艺孔深",
          "scope": "原生零件模型；材料、公差及表面要求以属性记录，未出工程图",
          "supersedes": "verification.json及drawing-correction.json当前状态，旧记录保留为历史"}
(out / "final-verification.json").write_text(
    json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"status": report["status"], "checks": checks}, ensure_ascii=False))
