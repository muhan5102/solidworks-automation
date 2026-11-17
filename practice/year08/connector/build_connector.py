"""第8届真题连接件练习。分阶段执行，重跑前先核对当前文档与保存点。"""
import argparse
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import (
    SKILL, member, mm, new_document, save_document, sketch, sketch_line,
    sketch_circle, revolve_boss, extrude_cut, _ensure_sketch_selected,
    fully_define_sketch, dimension_existing_relations, offset_plane, circle_cut,
    round_edges, chamfer_edges,
)
import win32com.client
from sw_review import save_preview, collect_geometry_measurements


def add_holes(model, sw):
    top = offset_plane(model, "Top Plane", 23, "后端基准面")
    # 以几何变换验证实际偏移方向，避免依赖基准面名称猜法向。
    return {
        "rear": circle_cut(model, sw, top, [(0, 23, -14)], 2, 4, -1, "定位孔"),
        "tap": circle_cut(model, sw, "Top Plane",
                          [(20.75, 0, 0), (-20.75, 0, 0), (0, 0, 20.75), (0, 0, -20.75)],
                          1.65, 9.5, 1, "螺纹底孔"),
    }


def add_spheres(model):
    result = []
    diagonal = 14 / math.sqrt(2)
    for index, (z, xs) in enumerate([(0, [14, -14]), (diagonal, [diagonal, -diagonal]), (-diagonal, [diagonal, -diagonal])]):
        plane = "Front Plane" if z == 0 else offset_plane(model, "Front Plane", z, f"球窝截面基准面_{index}")
        for x in xs:
            feature_name = f"球窝{len(result)+1}"
            if member(model, "FeatureByName", feature_name) is not None:
                result.append({"x_mm": x, "z_mm": z, "resumed_existing": True})
                continue
            with sketch(model, plane) as ref:
                cx, cy, radius = mm(x), mm(24.8), mm(2.8)
                arc = model.SketchManager.CreateArc(cx, cy, 0, cx, cy + radius, 0, cx, cy - radius, 0, -1)
                line = sketch_line(model, cx, cy - radius, cx, cy + radius)
                if any(item is None for item in (arc, line)):
                    raise RuntimeError("球窝截面创建失败")
                sd = member(model.SelectionManager, "CreateSelectData")
                sd.Mark = 0
                for a, b in [(member(arc, "GetStartPoint2"), member(line, "GetEndPoint2")),
                             (member(arc, "GetEndPoint2"), member(line, "GetStartPoint2"))]:
                    model.ClearSelection2(True)
                    if not a.Select4(False, sd) or not b.Select4(True, sd):
                        raise RuntimeError("球窝端点选择失败")
                    model.SketchAddConstraints("sgCOINCIDENT")
                quality = fully_define_sketch(model)
                active = member(model.SketchManager, "ActiveSketch")
                if not all(member(c, "IsClosed") for c in member(active, "GetSketchContours")):
                    raise RuntimeError("球窝截面未闭合")
            _ensure_sketch_selected(model, ref)
            sd.Mark = 16
            if not line.Select4(True, sd):
                raise RuntimeError("无法选择球窝回转轴")
            cut = model.FeatureManager.FeatureRevolve2(
                True, True, False, True, False, False, 0, 0,
                2 * math.pi, 0.0, False, False, 0.0, 0.0, 0, 0.0, 0.0, True, False, True)
            if cut is None:
                raise RuntimeError("球面凹窝回转切除失败")
            cut.Name = feature_name
            member(model, "FeatureByName", str(ref)).Name = f"球窝草图{len(result)+1}"
            result.append({"x_mm": x, "z_mm": z, "quality": quality})
            print(f"[球窝] {len(result)}/6，SR2.8，深1 mm", flush=True)
    return result


def finish(model):
    # 原题未要求实体牙型，用原生装饰螺纹保留 M4×0.7、深8 的语义。
    mouth = [(edge, pos) for edge, pos, r in round_edges(model)
             if abs(pos[1]) < .001 and abs(r-1.65) < .001]
    if len(mouth) != 4:
        raise RuntimeError("M4孔口数量不为4")
    threads = []
    for index, (edge, _pos) in enumerate(mouth, 1):
        name = f"M4螺纹{index}"
        if member(model, "FeatureByName", name) is not None:
            threads.append(name)
            continue
        model.ClearSelection2(True)
        if not edge.Select2(False, 0):
            raise RuntimeError("螺纹入口选择失败")
        feature = model.FeatureManager.InsertCosmeticThread3(8, "Tapped Hole", "M4x0.7", mm(4), 0, mm(8), "M4 深8")
        if feature is None:
            raise RuntimeError("原生装饰螺纹创建失败，需保留为未完成项")
        feature.Name = name
        threads.append(name)
    chamfer_edges(model, [([0, 0, 0], 17.5)], 1, "内腔口倒角")
    chamfer_edges(model, [([0, 0, 0], 24), ([0, 23, 0], 18), ([0, 23, 0], 7)], .5, "端面倒角")
    chamfer_edges(model, [(pos, 1.65) for _edge, pos in mouth], .5, "螺纹孔口倒角")
    chamfer_edges(model, [([0, 23, -14], 2)], .5, "定位孔口倒角")
    # 模板带有45钢，清除错误材料，原图牌号单独保留，不编造材料库参数。
    model.SetMaterialPropertyName2("", "", "")
    props = model.Extension.CustomPropertyManager("")
    for name, value in [
        ("图号", "SDF-02"), ("名称", "连接件"), ("材料", "YL12"),
        ("技术要求", "未注倒角C0.5；表面涂银粉漆"),
        ("螺纹", "4×M4×0.7，螺纹深8；底孔深9.5为工艺假设"),
        ("来源", "第8届计算机绘图真题，第1题，PDF第4页"),
    ]:
        props.Add3(name, 30, value, 2)
    if member(model, "GetVisibilityOfConstructPlanes"):
        member(model, "ViewDispRefplanes")
    model.ClearSelection2(True)
    return {"cosmetic_threads": threads, "material": "YL12（图纸属性，未配置物性）",
            "assumptions": ["M4粗牙螺距0.7，底孔直径3.3、深9.5；底孔深度为工艺选择"] }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--stage", choices=["body", "holes", "spheres", "finish"], required=True)
    parser.add_argument("--use-empty-active", action="store_true")
    parser.add_argument("--process-id", type=int, required=True)
    args = parser.parse_args()
    out = args.output.resolve()
    out.mkdir(parents=True, exist_ok=True)
    dest = out / "SDF-02-connector.SLDPRT"
    if dest.exists() and args.stage == "body":
        raise FileExistsError("主体阶段不覆盖已有模型，请检查保存点")
    started = time.monotonic()
    print(f"[{args.stage}] 连接可见会话，核对已有文档", flush=True)
    # 本轮已验证类工厂可返回现有实例，而 ROT GetActiveObject 不可用。
    # 核对 PID 后复用它；不修改全局连接封装或反复启动。
    sw = win32com.client.Dispatch("SldWorks.Application.32")
    if int(member(sw, "GetProcessID")) != args.process_id:
        raise RuntimeError("连接到非预期进程，停止写入")
    sw.Visible = True
    sw.UserControl = True
    if args.stage != "body":
        model = sw.ActiveDoc
        if model is None or Path(member(model, "GetPathName")).resolve() != dest:
            raise RuntimeError("当前文档与本轮保存路径不一致")
        guard = {"holes": "定位孔", "spheres": "球窝6", "finish": "定位孔口倒角"}[args.stage]
        if member(model, "FeatureByName", guard) is not None:
            raise RuntimeError("本阶段已有产物，禁止盲目重放")
        data = (add_holes(model, sw) if args.stage == "holes" else
                add_spheres(model) if args.stage == "spheres" else finish(model))
        model.ForceRebuild3(False)
        model.ClearSelection2(True)
        model.ShowNamedView2("*Isometric", 7)
        model.ViewZoomtofit2()
        model.GraphicsRedraw2()
        if not save_document(model):
            raise RuntimeError("阶段保存失败")
        save_preview(model, str(out / f"{args.stage}-isometric.bmp"), "isometric")
        evidence = {"stage": args.stage, "elapsed_seconds": round(time.monotonic()-started,3),
                    "details": data, "geometry": collect_geometry_measurements(model)}
        (out / f"{args.stage}-evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[完成] {args.stage}，实测 {evidence['elapsed_seconds']:.1f} 秒", flush=True)
        return
    existing = list(member(sw, "GetDocuments") or [])
    if existing and not args.use_empty_active:
        raise RuntimeError("主体首建要求空会话；请先核对已有文档")
    if args.use_empty_active:
        model = sw.ActiveDoc
        if model is None or member(model, "GetPathName") or member(model, "GetBodies2", 0, False):
            raise RuntimeError("不是本轮新建的未保存空零件")
        feature = member(model, "FirstFeature")
        while feature is not None:
            if member(feature, "GetTypeName2") in {"ProfileFeature", "3DProfileFeature"}:
                raise RuntimeError("空零件已有草图，拒绝覆盖")
            feature = member(feature, "GetNextFeature")
    else:
        model = new_document(sw, "part")
    if model is None:
        raise RuntimeError("新建零件失败")
    print("[主体] 前视基准面绘制回转截面，主轴为模型 Y 轴", flush=True)
    points = [
        (17.5, 0), (24, 0), (24, 14), (20, 21), (18, 21),
        (18, 23), (7, 23), (7, 14), (9, 14), (9, 15), (17.5, 15),
    ]
    with sketch(model, "Front Plane") as ref:
        old_add = model.SketchManager.AddToDB
        model.SketchManager.AddToDB = True
        segments = []
        try:
            center = model.SketchManager.CreateCenterLine(0, 0, 0, 0, mm(23), 0)
            if center is None:
                raise RuntimeError("回转轴创建失败")
            for first, second in zip(points, points[1:] + points[:1]):
                segment = sketch_line(model, mm(first[0]), mm(first[1]), mm(second[0]), mm(second[1]))
                if segment is None:
                    raise RuntimeError("截面线创建失败")
                segments.append(segment)
        finally:
            model.SketchManager.AddToDB = old_add
        sd = member(model.SelectionManager, "CreateSelectData")
        sd.Mark = 0
        for one, two in zip(segments, segments[1:] + segments[:1]):
            model.ClearSelection2(True)
            member(one, "GetEndPoint2").Select4(False, sd)
            member(two, "GetStartPoint2").Select4(True, sd)
            model.SketchAddConstraints("sgCOINCIDENT")
        for segment, first, second in zip(segments, points, points[1:] + points[:1]):
            relation = "sgVERTICAL" if first[0] == second[0] else "sgHORIZONTAL" if first[1] == second[1] else None
            if relation:
                model.ClearSelection2(True)
                segment.Select4(False, sd)
                model.SketchAddConstraints(relation)
        model.ClearSelection2(True)
        center.Select4(False, sd)
        model.SketchAddConstraints("sgVERTICAL")
        from sw_connect import create_empty_dispatch_variant
        for entity in (member(center, "GetStartPoint2"), segments[0]):
            model.ClearSelection2(True)
            entity.Select4(False, sd)
            if not model.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0,
                                             create_empty_dispatch_variant(), 0):
                raise RuntimeError("原点关系选择失败")
            model.SketchAddConstraints("sgCOINCIDENT")
        quality = dimension_existing_relations(model)
        for segment, expected in zip(segments, points):
            p = member(segment, "GetStartPoint2")
            if math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) > 1e-5:
                raise RuntimeError("草图定义改变了原始尺寸，停止回转")
        print("[草图] " + json.dumps(quality, ensure_ascii=False), flush=True)
    feature = revolve_boss(model, ref, 2 * math.pi)
    if feature is None:
        raise RuntimeError("回转主体失败")
    feature.Name = "主体回转"
    member(model, "FeatureByName", str(ref)).Name = "主体轴向截面"
    model.ForceRebuild3(False)
    model.ShowNamedView2("*Isometric", 7)
    model.ViewZoomtofit2()
    model.GraphicsRedraw2()
    if not save_document(model, str(dest)):
        raise RuntimeError("保存主体失败")
    save_preview(model, str(out / "01-body-isometric.bmp"), "isometric")
    evidence = {
        "stage": "body", "elapsed_seconds": round(time.monotonic() - started, 3),
        "revision": member(sw, "RevisionNumber"),
        "sketch": quality, "profile_mm": points,
        "geometry": collect_geometry_measurements(model),
    }
    (out / "01-body-evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[完成] 主体已保存；阶段实测 {evidence['elapsed_seconds']:.1f} 秒", flush=True)
    # 保留可见文档供用户观察，不进入自动清理会话上下文。


if __name__ == "__main__":
    main()
