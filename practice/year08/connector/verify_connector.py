"""读取本轮连接件原生几何、重开与代表性尺寸修改结果。"""
import argparse
import json
from pathlib import Path
import shutil
import sys
import time
import win32com.client

from build_connector import member, SKILL
from sw_connect import open_document, save_document
from sw_assembly import iter_feature_tree
from sw_sketch_quality import inspect_model_sketches
from sw_review import run_review, collect_geometry_measurements


def inspect(model):
    bodies = list(member(model, "GetBodies2", 0, False) or [])
    spheres = []
    for body in bodies:
        for face in member(body, "GetFaces"):
            surface = member(face, "GetSurface")
            if member(surface, "IsSphere"):
                spheres.append([v*1000 for v in member(surface, "SphereParams")])
    sketch_feature = member(model, "FeatureByName", "主体轴向截面")
    sk = member(sketch_feature, "GetSpecificFeature2")
    vertices = []
    for segment in member(sk, "GetSketchSegments"):
        if not member(segment, "ConstructionGeometry"):
            p = member(segment, "GetStartPoint2")
            vertices.append([round(member(p, k)*1000, 6) for k in ("X", "Y")])
    expected = [(17.5, 0), (24, 0), (24, 14), (20, 21), (18, 21), (18, 23),
                (7, 23), (7, 14), (9, 14), (9, 15), (17.5, 15)]
    geo = collect_geometry_measurements(model)
    names = [(member(f, "Name"), member(f, "GetTypeName2")) for f, _ in iter_feature_tree(model)]
    tapped = [name for name, kind in names if kind == "CosmeticThread"]
    cylinders = geo["cylindrical_faces"]
    d4 = [f for f in cylinders if f["internal"] and abs(f["diameter_mm"]-4)<1e-5]
    d33 = [f for f in cylinders if f["internal"] and abs(f["diameter_mm"]-3.3)<1e-5]
    quality = inspect_model_sketches(model)
    checks = {
        "one_solid": len(bodies) == 1,
        "profile_matches_drawing": set(map(tuple, vertices)) == set(expected),
        "six_SR2p8_seats": len(spheres) == 6 and all(abs(s[3]-2.8)<1e-6 and abs(s[1]-24.8)<1e-6 for s in spheres),
        "one_D4_hole": len(d4) == 1,
        "four_D3p3_pilot_holes": len(d33) == 4,
        "four_native_cosmetic_threads": len(tapped) == 4,
        "constrained_editable_sketches": quality["status"] == "pass",
        "wrong_template_material_removed": not bool(member(model, "MaterialIdName")),
    }
    return {"checks": checks, "profile_mm": vertices, "spheres_mm": spheres, "geometry": geo,
            "sketches": quality, "threads": tapped}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--process-id", type=int, required=True)
    args = parser.parse_args()
    out = args.output.resolve()
    path = out / "SDF-02-connector.SLDPRT"
    start = time.monotonic()
    sw = win32com.client.Dispatch("SldWorks.Application.32")
    assert int(member(sw, "GetProcessID")) == args.process_id
    model = sw.ActiveDoc
    assert Path(member(model, "GetPathName")).resolve() == path
    assert model.ForceRebuild3(False)
    assert save_document(model)
    before = inspect(model)
    print("[核验] " + json.dumps(before["checks"], ensure_ascii=False), flush=True)
    assert all(before["checks"].values()), "原生几何或草图检查未通过"
    copy_path = out / "checkpoints/editability-check.SLDPRT"
    assert not copy_path.exists()
    shutil.copy2(path, copy_path)
    copy = open_document(sw, str(copy_path), silent=True)
    sk = member(member(copy, "FeatureByName", "主体轴向截面"), "GetSpecificFeature2")
    candidates = []
    for relation in member(member(sk, "RelationManager"), "GetRelations", 0):
        display = member(relation, "GetDisplayDimension")
        if display is not None:
            dimension = member(display, "GetDimension2", 0)
            if abs(member(dimension, "SystemValue")-.018) < 1e-9:
                candidates.append(dimension)
    assert len(candidates) == 1, "后端半径驱动尺寸不唯一"
    dim = candidates[0]
    dim.SystemValue = .0185
    assert copy.ForceRebuild3(False)
    edited = collect_geometry_measurements(copy)
    editable = any(abs(f["diameter_mm"]-37)<1e-5 for f in edited["cylindrical_faces"])
    assert editable, "代表尺寸修改没有反映到模型"
    sw.CloseDoc(member(copy, "GetTitle"))
    sw.CloseDoc(member(model, "GetTitle"))
    model = open_document(sw, str(path), silent=True)
    assert model.ForceRebuild3(False)
    after = inspect(model)
    assert all(after["checks"].values()), "保存重开后检查不一致"
    review, review_path = run_review(model, str(out/"review"), basename="连接件",
                                    views=["isometric", "top", "bottom", "front"], expected_outputs=[str(path)])
    model.ShowNamedView2("*Isometric", 7)
    model.ViewZoomtofit2()
    model.GraphicsRedraw2()
    result = {"before": before, "after_reopen": after, "representative_edit": {
        "copy_only": True, "diameter_change_mm": [36,37], "rebuild_and_geometry_passed": editable,
        "original_restored_by_reopen": True}, "review_path": str(Path(review_path).relative_to(out)),
        "review_evaluation": review.get("evaluation"), "elapsed_seconds": round(time.monotonic()-start,3)}
    (out/"verification.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
    print("[完成] 保存重开与副本尺寸修改通过；模型已留在可见窗口", flush=True)


if __name__ == "__main__":
    main()
