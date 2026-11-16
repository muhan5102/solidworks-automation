"""第8届第2页钢球和O型圈简化图；自行原生建模，不调用标准件库。"""
import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "connector"))
from build_connector import member, mm, new_document, save_document, sketch, sketch_line, sketch_circle, revolve_boss, dimension_existing_relations
from sw_part import extrude_boss
from sw_connect import open_document, create_empty_dispatch_variant
from sw_sketch_quality import fully_define_sketch, inspect_model_sketches
from sw_review import save_preview, collect_model_summary
from sw_export import export_to_step
import win32com.client

parser = argparse.ArgumentParser()
parser.add_argument("output", type=Path)
parser.add_argument("--process-id", type=int, required=True)
args = parser.parse_args()
out = args.output.resolve()
assert out.is_dir()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == args.process_id
sw.Visible = True
sw.UserControl = True
reports = []

for name, kind, values, material, quantity in [
    ("钢球", "sphere", (5.5,), "GCr15", 2),
    ("O型圈10", "ring", (14, 11, 2), "丁腈橡胶", 1),
    ("O型圈13", "ring", (35, 32, 2), "丁腈橡胶", 1),
]:
    started = time.monotonic()
    path = out / (name + ".SLDPRT")
    assert not path.exists()
    model = new_document(sw, "part")
    if kind == "sphere":
        diameter = values[0]
        radius = diameter / 2
        with sketch(model, "Front Plane") as ref:
            sm = model.SketchManager
            old = sm.AddToDB
            sm.AddToDB = True
            try:
                axis = sm.CreateCenterLine(0, 0, 0, 0, mm(diameter), 0)
                arc = sm.CreateArc(0, mm(radius), 0, 0, 0, 0, 0, mm(diameter), 0, 1)
                line = sketch_line(model, 0, mm(diameter), 0, 0)
            finally:
                sm.AddToDB = old
            assert axis is not None and arc is not None and line is not None
            sd = member(model.SelectionManager, "CreateSelectData")
            for a, b in [(member(arc, "GetEndPoint2"), member(line, "GetStartPoint2")),
                         (member(line, "GetEndPoint2"), member(arc, "GetStartPoint2"))]:
                model.ClearSelection2(True)
                assert a.Select4(False, sd) and b.Select4(True, sd)
                model.SketchAddConstraints("sgCOINCIDENT")
            for segment in [axis, line]:
                model.ClearSelection2(True)
                assert segment.Select4(False, sd)
                model.SketchAddConstraints("sgVERTICAL")
            for point in [member(axis, "GetStartPoint2"), member(arc, "GetStartPoint2")]:
                model.ClearSelection2(True)
                assert point.Select4(False, sd)
                assert model.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0,
                                                   True, 0, create_empty_dispatch_variant(), 0)
                model.SketchAddConstraints("sgCOINCIDENT")
            dimension_existing_relations(model)
            assert abs(member(arc, "GetRadius") - mm(radius)) < 1e-9
        feature = revolve_boss(model, ref, 2*math.pi)
        assert feature is not None
        feature.Name = "球体回转"
    else:
        outer, inner, thickness = values
        with sketch(model, "Front Plane") as ref:
            assert sketch_circle(model, 0, 0, mm(outer/2)) is not None
            assert sketch_circle(model, 0, 0, mm(inner/2)) is not None
            fully_define_sketch(model)
        feature = extrude_boss(model, ref, mm(thickness))
        assert feature is not None
        feature.Name = "环形主体"
    model.FeatureByName(str(ref)).Name = "截面草图"
    for key, value in [("名称", name), ("材料", material), ("数量", str(quantity)),
                       ("题面尺寸mm", "×".join(str(v) for v in values)),
                       ("表达方式", "完整球面" if kind == "sphere" else "按原卷D/d/L表的矩形截面简化环，不作圆截面O圈")]:
        model.Extension.CustomPropertyManager("").Add3(key, 30, value, 2)
    assert model.ForceRebuild3(False)
    assert save_document(model, str(path))
    model.ShowNamedView2("*Isometric", 7)
    model.ViewZoomtofit2()
    model.GraphicsRedraw2()
    sw.CloseDoc(member(model, "GetTitle"))
    model = open_document(sw, str(path), silent=True)
    assert Path(member(model, "GetPathName")).resolve() == path
    assert model.ForceRebuild3(False)
    bodies = list(member(model, "GetBodies2", 0, False))
    quality = inspect_model_sketches(model)
    health = collect_model_summary(model)
    surfaces = []
    for face in member(bodies[0], "GetFaces"):
        surface = member(face, "GetSurface")
        if member(surface, "IsSphere"):
            surfaces.append({"sphere": list(member(surface, "SphereParams")), "area_mm2": member(face, "GetArea")*1e6})
        if member(surface, "IsCylinder"):
            cp = list(member(surface, "CylinderParams"))
            surfaces.append({"cylinder": cp, "area_mm2": member(face, "GetArea")*1e6})
    checks = {"single_solid": len(bodies) == 1, "fully_defined_no_fix": quality["status"] == "pass",
              "no_error_or_warning": not health["faulty_features"] and not health["warning_features"]}
    if kind == "sphere":
        checks["D5.5_complete_sphere"] = len(surfaces) == 1 and "sphere" in surfaces[0] and (
            abs(surfaces[0]["sphere"][3]*2000-5.5) < 1e-6 and
            abs(surfaces[0]["area_mm2"]-4*math.pi*2.75**2) < 1e-5)
    else:
        for diameter in [outer, inner]:
            checks[f"D{diameter}_coaxial_height2"] = any("cylinder" in s and
                abs(s["cylinder"][6]*2000-diameter) < 1e-6 and
                math.hypot(s["cylinder"][0], s["cylinder"][1]) < 1e-8 and
                abs(s["area_mm2"]-math.pi*diameter*2) < 1e-5 for s in surfaces)
    assert all(checks.values()), checks
    assert export_to_step(model, str(out/(name+".step")))
    save_preview(model, str(out/(name+"_等轴测.bmp")), "isometric")
    save_preview(model, str(out/(name+"_正视.bmp")), "front")
    assert save_document(model)
    report = {"name": name, "checks": checks, "surfaces": surfaces, "sketches": quality,
              "native_health": health, "reopened": True,
              "execution_seconds": round(time.monotonic()-started, 3),
              "finished": datetime.now().astimezone().isoformat(), "material": material+"仅属性"}
    (out/(name+"_verification.json")).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    reports.append({k: report[k] for k in ["name", "checks", "execution_seconds", "finished"]})
    print(json.dumps(reports[-1], ensure_ascii=False), flush=True)
    if name != "O型圈13":
        sw.CloseDoc(member(model, "GetTitle"))

model.ShowNamedView2("*Isometric", 7)
model.ViewZoomtofit2()
model.GraphicsRedraw2()
assert save_document(model)
(out/"verification.json").write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
