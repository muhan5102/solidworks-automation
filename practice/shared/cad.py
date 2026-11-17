"""项目内共享的CAD辅助函数；题目专属结构留在各案例。"""
import math
import os
from pathlib import Path
import sys

SKILL = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex")) / "skills" / "solidworks-automation"
sys.path.insert(0, str(SKILL / "scripts"))
from sw_connect import get_com_member as member, mm, new_document, save_document
from sw_part import sketch, sketch_line, sketch_circle, revolve_boss, extrude_cut, _ensure_sketch_selected
from sw_sketch_quality import fully_define_sketch


def dimension_existing_relations(model):
    """只补尺寸，避免自动推断把相近台阶误判为共线。"""
    from sw_connect import create_empty_dispatch_variant
    from sw_sketch_quality import inspect_sketch
    model.ClearSelection2(True)
    if not model.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, False, 6,
                                     create_empty_dispatch_variant(), 0):
        raise RuntimeError("原点选择失败")
    model.SketchManager.FullyDefineSketch(True, False, 0, True, 1, create_empty_dispatch_variant(),
                                         1, create_empty_dispatch_variant(), 1, 1)
    result = inspect_sketch(member(model.SketchManager, "ActiveSketch"))
    if result["status"] != "pass":
        raise RuntimeError(f"尺寸定义未通过: {result}")
    model.ClearSelection2(True)
    return result


def offset_plane(model, base, offset_mm, name):
    from sw_part import _get_plane_name_candidates, _select_first_candidate
    if member(model, "FeatureByName", name) is not None:
        return name
    model.ClearSelection2(True)
    if not _select_first_candidate(model.Extension, _get_plane_name_candidates(base), "PLANE"):
        raise RuntimeError(f"基准面选择失败: {base}")
    plane = model.FeatureManager.InsertRefPlane(8 | (256 if offset_mm < 0 else 0), mm(abs(offset_mm)), 0, 0.0, 0, 0.0)
    if plane is None:
        raise RuntimeError("偏移面创建失败")
    feature = member(model, "FirstFeature")
    last_plane = None
    while feature is not None:
        if member(feature, "GetTypeName2") == "RefPlane":
            last_plane = feature
        feature = member(feature, "GetNextFeature")
    if last_plane is None:
        raise RuntimeError("无法回读新增基准面")
    last_plane.Name = name
    model.ClearSelection2(True)
    return name


def circle_cut(model, sw, plane, centers_mm, radius_mm, depth_mm, direction_y, name):
    with sketch(model, plane) as ref:
        active = member(model.SketchManager, "ActiveSketch")
        transform = member(active, "ModelToSketchTransform")
        inv = member(transform, "Inverse")
        matrix = list(member(transform, "ArrayData"))
        normal_y = list(member(inv, "ArrayData"))[7]
        for center in centers_mm:
            point = [mm(v) for v in center]
            xyz = [sum(point[j] * matrix[j*3+i] for j in range(3)) * matrix[12] + matrix[9+i]
                   for i in range(3)]
            if abs(xyz[2]) > 1e-7:
                raise RuntimeError(f"孔中心不在草图平面: {xyz}")
            if sketch_circle(model, xyz[0], xyz[1], mm(radius_mm)) is None:
                raise RuntimeError("孔草图创建失败")
        quality = fully_define_sketch(model)
    feature = extrude_cut(model, ref, mm(depth_mm), direction=(normal_y * direction_y > 0))
    if feature is None:
        raise RuntimeError(f"切除失败: {name}")
    feature.Name = name
    member(model, "FeatureByName", str(ref)).Name = "孔位草图_" + name
    print(f"[孔槽] {name}，草图完全定义={quality['status']}", flush=True)
    return quality


def round_edges(model):
    from sw_edge_select import _edge_curve
    seen = set()
    result = []
    for body in member(model, "GetBodies2", 0, False):
        for face in member(body, "GetFaces"):
            for edge in member(face, "GetEdges"):
                curve = _edge_curve(edge)
                if not member(curve, "IsCircle"):
                    continue
                values = list(member(curve, "CircleParams"))
                key = tuple(round(v, 9) for v in values[:3] + values[6:7])
                if key not in seen:
                    seen.add(key)
                    result.append((edge, [v * 1000 for v in values[:3]], values[6] * 1000))
    return result


def chamfer_edges(model, targets, distance, name):
    if member(model, "FeatureByName", name) is not None:
        return
    edges = round_edges(model)
    model.ClearSelection2(True)
    sd = member(model.SelectionManager, "CreateSelectData")
    sd.Mark = 0
    for center, radius in targets:
        matches = [edge for edge, pos, r in edges if math.dist(pos, center) < .001 and abs(r-radius) < .001]
        if len(matches) != 1 or not matches[0].Select4(True, sd):
            raise RuntimeError(f"倒角边不唯一或不可选: {center}, {radius}")
    from sw_part import chamfer
    feature = chamfer(model, mm(distance))
    if feature is None:
        raise RuntimeError(f"倒角失败: {name}")
    feature.Name = name
    model.ClearSelection2(True)
