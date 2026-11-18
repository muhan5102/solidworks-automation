"""本项目单件练习的重复代码；复用已安装Skill，不替代题面识读或目视复核。"""
from contextlib import contextmanager
from datetime import datetime
import hashlib
import json
from pathlib import Path
import sys
import time

import win32com.client
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g, create_empty_dispatch_variant
from sw_export import export_to_step
from sw_review import collect_model_summary, save_preview
from sw_sketch_quality import inspect_model_sketches


def relation(model, items, kind):
    model.ClearSelection2(True)
    data = g(model.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    model.SketchAddConstraints(kind)


def anchor(model, point, *, x=False, y=False):
    """明确相对原点的零偏移；调用者决定设计关系，不自动猜测所有点。"""
    assert x or y
    if x:
        assert abs(g(point, "X")) < 1e-8
    if y:
        assert abs(g(point, "Y")) < 1e-8
    model.ClearSelection2(True)
    assert point.Select4(False, g(model.SelectionManager, "CreateSelectData"))
    assert model.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0,
                                      create_empty_dispatch_variant(), 0)
    model.SketchAddConstraints("sgCOINCIDENT" if x and y else
                              "sgVERTICALPOINTS2D" if x else "sgHORIZONTALPOINTS2D")


@contextmanager
def quiet_dimensions(sw):
    old = sw.GetUserPreferenceToggle(10)
    sw.SetUserPreferenceToggle(10, False)
    try:
        yield
    finally:
        sw.SetUserPreferenceToggle(10, old)


def drive(model, items, position, expected, kind="normal"):
    """回读原图驱动值和驱动状态；角度的两条参照仍由题面决定。"""
    model.ClearSelection2(True)
    data = g(model.SelectionManager, "CreateSelectData")
    for i, item in enumerate(items):
        assert item.Select4(i > 0, data)
    methods = {"normal": "AddDimension2", "vertical": "AddVerticalDimension2",
               "horizontal": "AddHorizontalDimension2", "diameter": "AddDiameterDimension2"}
    dd = getattr(model, methods[kind])(*position)
    assert dd is not None
    dim = g(dd, "GetDimension2", 0)
    assert abs(g(dim, "SystemValue")-expected) < 1e-8
    assert g(dim, "DrivenState") == 2
    return dim


def circle(model, x, y, radius):
    """避免创建圆时暗中推断外部关系；后续显式定位和加驱动。"""
    manager = model.SketchManager
    old = manager.AddToDB
    manager.AddToDB = True
    try:
        result = manager.CreateCircleByRadius(x, y, 0, radius)
        assert result is not None
        return result
    finally:
        manager.AddToDB = old


def dimensions(model, feature_name):
    feature = model.FeatureByName(feature_name)
    assert feature is not None
    dd = g(feature, "GetFirstDisplayDimension")
    result = []
    while dd is not None:
        dim = g(dd, "GetDimension2", 0)
        result.append({"name": g(dim, "FullName"), "value": g(dim, "SystemValue"),
                       "driven": g(dim, "DrivenState")})
        dd = g(feature, "GetNextDisplayDimension", dd)
    return result


def surfaces(model, *, hole_evidence=False):
    result = []
    for body in g(model, "GetBodies2", 0, False):
        for face in g(body, "GetFaces"):
            surface = g(face, "GetSurface")
            row = {"box": [v*1000 for v in g(face, "GetBox")]}
            for kind in ["Cylinder", "Cone", "Torus", "Sphere", "Plane"]:
                if g(surface, "Is"+kind):
                    row[kind] = list(g(surface, kind+"Params"))
                    if hole_evidence and kind in ("Cylinder", "Cone"):
                        row["internal"] = bool(g(face, "FaceInSurfaceSense"))
                        row["area_mm2"] = g(face, "GetArea") * 1_000_000
                    break
            result.append(row)
    return result


def save(model, path):
    assert Path(g(model, "GetPathName")).resolve() == Path(path).resolve()
    errors = win32com.client.VARIANT(16387, 0)
    warnings = win32com.client.VARIANT(16387, 0)
    assert model.Save3(1, errors, warnings) and errors.value == 0
    assert not g(model, "GetSaveFlag")


def prepare_review(sw, target, check_geometry, *, sketch_count, views, sketch_quality_check=None):
    """保存/导出/重开/通用检查一次完成；题面专项检查通过回调保留。"""
    target = Path(target).resolve()
    out = target.parent
    report_path = out/"verification.json"
    assert target.is_file() and not report_path.exists()
    model = g(sw, "ActiveDoc")
    assert Path(g(model, "GetPathName")).resolve() == target
    started = datetime.now().astimezone()
    timer = time.perf_counter()
    save(model, target)
    step = target.with_suffix(".step")
    assert not step.exists() and export_to_step(model, str(step))
    assert step.is_file() and step.stat().st_size > 0
    assert Path(g(model, "GetPathName")).resolve() == target
    sw.CloseDoc(g(model, "GetTitle"))
    errors = win32com.client.VARIANT(16387, 0)
    warnings = win32com.client.VARIANT(16387, 0)
    model = sw.OpenDoc6(str(target), 1, 1, "", errors, warnings)
    assert model is not None and errors.value == 0
    assert Path(g(model, "GetPathName")).resolve() == target and model.ForceRebuild3(False)
    quality, health = inspect_model_sketches(model), collect_model_summary(model)
    rows = surfaces(model)
    quality_name = "手工草图数量与完全定义无Fix无孤立"
    quality_pass = len(quality["sketches"]) == sketch_count and quality["status"] == "pass"
    if sketch_quality_check is not None:
        quality_name, quality_pass = sketch_quality_check(quality)
    checks = {
        "单实体": len(g(model, "GetBodies2", 0, False)) == 1,
        "特征健康读取完整且无错误警告": not health.get("feature_error")
            and not health["faulty_features"] and not health["warning_features"]
            and not health["unknown_feature_health"],
        quality_name: quality_pass,
    }
    checks.update(check_geometry(model, rows))
    images = []
    try:
        for view in views:
            wire = view.endswith("_wire")
            name = view.removesuffix("_wire")
            if wire:
                model.ViewDisplayWireframe()
            path = save_preview(model, str(out/f"review_{view}.bmp"), name)
            if wire:
                model.ViewDisplayShaded()
            images.append((view, path))
    finally:
        model.ViewDisplayShaded()
        model.ShowNamedView2("*Isometric", 7)
        model.GraphicsRedraw2()
    # 一张对照图集中查看，避免逐视图产生多次工具周转。
    grid = Image.new("RGB", (800*len(images), 525), "white")
    draw = ImageDraw.Draw(grid)
    for i, (view, path) in enumerate(images):
        with Image.open(path) as image:
            grid.paste(image.convert("RGB").resize((800, 500)), (i*800, 25))
        draw.text((i*800+12, 7), view, fill="black")
    grid.save(out/"review.png")
    with Image.open(images[0][1]) as image:
        image.save(out/"isometric.png")
    report = {
        "started": started.isoformat(), "finished": datetime.now().astimezone().isoformat(),
        "execution_seconds": time.perf_counter()-timer, "target": target.name,
        "process_id": g(sw, "GetProcessID"), "passed": all(checks.values()), "checks": checks,
        "status": "awaiting_visual_review" if all(checks.values()) else "failed",
        "quality": quality, "health": health, "surfaces": rows,
        "limits": "仍需与原图目视对照；STEP未重导入，未制作工程图或装配。",
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"review": report["status"], "seconds": report["execution_seconds"],
                      "failed": [k for k, v in checks.items() if not v]}, ensure_ascii=False), flush=True)
    return report


def finish_review(sw, target, *, visual_notes):
    """仅在实际看过对照图后调用，核对身份并完成保存关闭。"""
    target = Path(target).resolve()
    out = target.parent
    report_path = out/"verification.json"
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["passed"] and report["status"] == "awaiting_visual_review" and visual_notes
    assert report["process_id"] == g(sw, "GetProcessID") and report["target"] == target.name
    model = g(sw, "ActiveDoc")
    assert Path(g(model, "GetPathName")).resolve() == target
    save(model, target)
    sw.CloseDoc(g(model, "GetTitle"))
    remaining = []
    doc = g(sw, "GetFirstDocument")
    while doc is not None:
        path = g(doc, "GetPathName")
        assert not path or Path(path).resolve() != target
        remaining.append(path or g(doc, "GetTitle"))
        doc = g(doc, "GetNext")
    now = datetime.now().astimezone()
    result = {
        "finished": now.isoformat(), "visual_notes": visual_notes, "remaining_documents": len(remaining),
        "remaining": remaining, "files": [{"name": p.name, "bytes": p.stat().st_size,
            "sha256": hashlib.sha256(p.read_bytes()).hexdigest()} for p in [target, target.with_suffix(".step")]],
    }
    start_path = out/"start.json"
    if start_path.exists():
        start = json.loads(start_path.read_text(encoding="utf-8")).get("reading_start")
        if start:
            result["reading_to_saved_closed_seconds"] = (now-datetime.fromisoformat(start)).total_seconds()
    report.update(status="passed", visual_notes=visual_notes)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    (out/"completion.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"saved_closed": target.name, "remaining_documents": len(remaining),
                      "reading_seconds": result.get("reading_to_saved_closed_seconds")}, ensure_ascii=False))
    return result
