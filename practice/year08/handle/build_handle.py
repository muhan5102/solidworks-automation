"""第8届手柄单件试跑；复用连接件已验证的草图关系与尺寸路线。"""
import argparse
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "connector"))
from build_connector import (
    member, mm, new_document, save_document, sketch, sketch_line,
    revolve_boss, dimension_existing_relations, round_edges, chamfer_edges,
)
from sw_connect import open_document, create_empty_dispatch_variant
from sw_part import fillet
from sw_sketch_quality import inspect_model_sketches
from sw_assembly import iter_feature_tree
from sw_review import collect_geometry_measurements, save_preview
import win32com.client


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--process-id", required=True, type=int)
    args = parser.parse_args()
    out = args.output.resolve()
    if not out.is_dir():
        raise RuntimeError("输出目录须预先核对并创建")
    dest = out / "SDF-03-handle.SLDPRT"
    if dest.exists():
        raise FileExistsError("不覆盖首次产物")
    started = time.monotonic()
    times = {}
    sw = win32com.client.Dispatch("SldWorks.Application.32")
    if int(member(sw, "GetProcessID")) != args.process_id:
        raise RuntimeError("非预期SolidWorks实例")
    sw.Visible = True
    sw.UserControl = True
    model = new_document(sw, "part")
    if model is None:
        raise RuntimeError("新建零件失败")

    def checkpoint(label):
        if not model.ForceRebuild3(False):
            raise RuntimeError(f"{label}重建失败")
        model.ClearSelection2(True)
        model.ShowNamedView2("*Isometric", 7)
        model.ViewZoomtofit2()
        model.GraphicsRedraw2()
        if not save_document(model, str(dest)):
            raise RuntimeError(f"{label}保存失败")
        times[label] = round(time.monotonic() - started, 3)
        print(f"[{label}] 已重建、保存、刷新，累计{times[label]:.1f}秒", flush=True)

    print("[主体] 总长62；退刀槽在轴向6～8；锥段26～34", flush=True)
    # 从题面尺寸链独立计算：肩部8；槽宽2；细轴18；末段28。
    points = [(0, 0), (4, 0), (4, 6), (3.25, 6), (3.25, 8),
              (4.5, 8), (4.5, 26), (6, 34), (6, 62), (0, 62)]
    with sketch(model, "Front Plane") as ref:
        manager = model.SketchManager
        old_add = manager.AddToDB
        manager.AddToDB = True
        try:
            axis = manager.CreateCenterLine(0, 0, 0, 0, mm(62), 0)
            segments = [sketch_line(model, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                        for a, b in zip(points, points[1:] + points[:1])]
        finally:
            manager.AddToDB = old_add
        if axis is None or any(s is None for s in segments):
            raise RuntimeError("草图几何创建失败")
        sd = member(model.SelectionManager, "CreateSelectData")
        sd.Mark = 0
        for first, second in zip(segments, segments[1:] + segments[:1]):
            model.ClearSelection2(True)
            assert member(first, "GetEndPoint2").Select4(False, sd)
            assert member(second, "GetStartPoint2").Select4(True, sd)
            model.SketchAddConstraints("sgCOINCIDENT")
        for segment, a, b in zip(segments, points, points[1:] + points[:1]):
            relation = "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL" if a[1] == b[1] else None
            if relation:
                model.ClearSelection2(True)
                assert segment.Select4(False, sd)
                model.SketchAddConstraints(relation)
        model.ClearSelection2(True)
        assert axis.Select4(False, sd)
        model.SketchAddConstraints("sgVERTICAL")
        for point in (member(axis, "GetStartPoint2"), member(segments[0], "GetStartPoint2")):
            model.ClearSelection2(True)
            assert point.Select4(False, sd)
            assert model.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT",
                                              0, 0, 0, True, 0,
                                              create_empty_dispatch_variant(), 0)
            model.SketchAddConstraints("sgCOINCIDENT")
        quality = dimension_existing_relations(model)
        for segment, expected in zip(segments, points):
            p = member(segment, "GetStartPoint2")
            assert math.dist([member(p, "X") * 1000, member(p, "Y") * 1000], expected) < 1e-5
        contours = list(member(member(manager, "ActiveSketch"), "GetSketchContours") or [])
        assert contours and all(member(c, "IsClosed") for c in contours)
    body = revolve_boss(model, ref, 2 * math.pi)
    if body is None:
        raise RuntimeError("主体回转失败")
    body.Name = "主体回转"
    member(model, "FeatureByName", str(ref)).Name = "轴向截面"
    checkpoint("主体")

    edges = round_edges(model)
    end = [edge for edge, pos, r in edges if math.dist(pos, [0, 62, 0]) < .001 and abs(r - 6) < .001]
    assert len(end) == 1
    model.ClearSelection2(True)
    assert end[0].Select2(False, 0)
    rounding = fillet(model, mm(2))
    if rounding is None:
        raise RuntimeError("端部R2失败")
    rounding.Name = "端部圆角"

    # 自定义装饰螺纹：M8粗牙螺距1.25，外螺纹基本小径约6.4664。
    # 仅表达原生螺纹语义，不创建实体牙型；失败时明确保留未完成状态。
    thread_status = "未完成"
    thread_errors = []
    mouth = [edge for edge, pos, r in round_edges(model)
             if math.dist(pos, [0, 0, 0]) < .001 and abs(r - 4) < .001]
    assert len(mouth) == 1
    try:
        model.ClearSelection2(True)
        assert mouth[0].Select2(False, 0)
        thread = model.FeatureManager.InsertCosmeticThread3(
            0, "", "", mm(6.4664), 0, mm(6), "M8")
        if thread is None:
            raise RuntimeError("InsertCosmeticThread3返回空")
        thread.Name = "M8外螺纹"
        thread_status = "装饰螺纹，待回读"
    except Exception as exc:
        thread_errors.append(str(exc))
    chamfer_edges(model, [([0, 0, 0], 4), ([0, 6, 0], 4), ([0, 8, 0], 4.5)], .5, "端口倒角")
    props = model.Extension.CustomPropertyManager("")
    for name, value in [("名称", "手柄"), ("图号", "SDF-03"), ("材料", "45"),
                        ("技术要求", "表面镀锌；未标注倒角C0.5"),
                        ("螺纹", "M8，粗牙1.25；名义螺纹段长6，含入口倒角"),
                        ("来源", "第8届计算机绘图真题，第1题，PDF第3页")]:
        props.Add3(name, 30, value, 2)
    checkpoint("收尾")
    sw.CloseDoc(member(model, "GetTitle"))
    model = open_document(sw, str(dest), silent=True)
    assert Path(member(model, "GetPathName")).resolve() == dest
    assert model.ForceRebuild3(False)
    geometry = collect_geometry_measurements(model)
    threads = [(member(f, "Name"), member(f, "GetTypeName2"))
               for f, _ in iter_feature_tree(model)
               if member(f, "GetTypeName2") == "CosmeticThread"]
    sketches = inspect_model_sketches(model)
    checks = {"one_solid": len(member(model, "GetBodies2", 0, False)) == 1,
              "sketches_constrained": sketches["status"] == "pass",
              "one_cosmetic_thread": len(threads) == 1,
              "profile_checked_before_revolve": quality["status"] == "pass"}
    # 原图尺寸链的独立核验；倒角后的圆柱长度不直接当作原始设计段长。
    final_edges = [(pos, r) for _, pos, r in round_edges(model)]
    def has_edge(y, radius):
        return any(abs(pos[1] - y) < 1e-5 and abs(r - radius) < 1e-5
                   for pos, r in final_edges)
    cylinders = geometry["cylindrical_faces"]
    checks["external_diameters"] = (
        len(cylinders) == 4 and all(not f["internal"] for f in cylinders)
        and sorted(round(f["diameter_mm"], 5) for f in cylinders) == [6.5, 8, 9, 12])
    checks["neck_and_cone_positions"] = all(
        has_edge(y, r) for y, r in [(6, 3.25), (8, 3.25), (26, 4.5), (34, 6)])
    checks["total_length_62"] = all(
        any(abs(pos[1] - y) < 1e-5 for pos, _ in final_edges) for y in [0, 62])
    checks["three_C0p5_pairs"] = all(
        has_edge(y, r) for y, r in [(0, 3.5), (.5, 4), (5.5, 4),
                                    (6, 3.5), (8, 4), (8.5, 4.5)])
    tori = []
    for solid in member(model, "GetBodies2", 0, False):
        for face in member(solid, "GetFaces"):
            surface = member(face, "GetSurface")
            if member(surface, "IsTorus"):
                tori.append(list(member(surface, "TorusParams")))
    checks["tip_R2"] = len(tori) == 1 and abs(tori[0][7] - .002) < 1e-8
    thread_data = []
    for feature, _ in iter_feature_tree(model):
        if member(feature, "GetTypeName2") == "CosmeticThread":
            definition = member(feature, "GetDefinition")
            thread_data.append({key: member(definition, key)
                                for key in ("Diameter", "BlindDepth", "ThreadCallout", "EndCondition")})
    checks["M8_callout_depth_6"] = (
        len(thread_data) == 1 and thread_data[0]["ThreadCallout"] == "M8"
        and abs(thread_data[0]["BlindDepth"] - .006) < 1e-8)
    save_preview(model, str(out / "手柄_isometric.bmp"), "isometric")
    save_preview(model, str(out / "手柄_front.bmp"), "front")
    model.ShowNamedView2("*Isometric", 7)
    model.ViewZoomtofit2()
    model.GraphicsRedraw2()
    if thread_data and checks["M8_callout_depth_6"]:
        thread_status = "原生装饰外螺纹，标注和深度已回读"
    report = {"checks": checks, "geometry": geometry, "sketches": sketches,
              "threads": threads, "thread_status": thread_status,
              "thread_definition": thread_data,
              "thread_errors": thread_errors, "times_cumulative": times,
              "execution_seconds": round(time.monotonic() - started, 3),
              "material_id": str(member(model, "MaterialIdName")),
              "note": "首轮简单件；仍须目视对照题面；螺纹为自定义装饰表达；材料45仅属性；无工程图"}
    (out / "verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: report[k] for k in ("checks", "thread_errors", "execution_seconds", "material_id")}, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
