"""用现有参数化压缩弹簧的副本派生第8届两种右旋弹簧。"""
from datetime import datetime
import hashlib
import json
import math
from pathlib import Path
import shutil
import sys

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.single_part import g, save
from sw_connect import _attach_existing_with_pid_check, create_empty_dispatch_variant, solidworks_processes
from sw_export import export_to_step
from sw_review import collect_model_summary, save_preview
from sw_spring import exact_body_bounds_mm


def open_part(sw, path):
    errors, warnings = (win32com.client.VARIANT(16387, 0) for _ in range(2))
    doc = sw.OpenDoc6(str(path), 1, 1, "", errors, warnings)
    assert doc is not None and errors.value == 0 and Path(g(doc, "GetPathName")).resolve() == path
    return doc


def equations(doc):
    manager = g(doc, "GetEquationMgr")
    member = manager._oleobj_.GetIDsOfNames("Equation")
    return manager, member


def read_equations(manager, member):
    return [manager._oleobj_.Invoke(member, 0, pythoncom.DISPATCH_PROPERTYGET, True, i)
            for i in range(g(manager, "GetCount"))]


def build(sw, seed, out, code, name, wire, mean, total, active, height, pitch, *, resume=False):
    part = out / f"{name}_{code}.SLDPRT"
    assert not part.with_suffix(".step").exists()
    started = datetime.now().astimezone().isoformat()
    if resume:
        doc = g(sw, "ActiveDoc")
        assert doc is not None and Path(g(doc, "GetPathName")).resolve() == part
    else:
        assert not part.exists()
        shutil.copy2(seed, part)
        doc = open_part(sw, part)
    manager, member = equations(doc)
    manager.AutomaticSolveOrder = True
    for key, value in (("D2", f"{mean}mm"), ("p", f"{pitch}mm")):
        if not any(entry.split("=")[0].strip().strip('"') == key for entry in read_equations(manager, member)):
            assert manager.Add2(-1, f'"{key}" = {value}', True) >= 0
    updates = {
        "d": f"{wire}mm", "D1": '"D2" + "d"', "n1": str(total), "n": str(active),
        "H": f"{height}mm", "D40002@弹簧": '"p"', "D40003@弹簧": '"p"',
        "D40001@弹簧": '("H" - "d" - ("n" - 0.5) * "p") / 1.5',
        "D40004@弹簧": '"D40001@弹簧"',
    }
    old = read_equations(manager, member)
    indexes = {entry.split("=")[0].strip().strip('"'): i for i, entry in enumerate(old)}
    assert updates.keys() <= indexes.keys()
    for key, value in updates.items():
        manager._oleobj_.Invoke(member, 0, pythoncom.DISPATCH_PROPERTYPUT, False, indexes[key], f'"{key}" = {value}')
    g(manager, "EvaluateAll")
    assert doc.ForceRebuild3(False)
    helix = doc.FeatureByName("弹簧")
    definition = g(helix, "GetDefinition")
    # 此模板螺旋线的基准面与路径方向下，False对应实体右旋。
    if g(definition, "Clockwise"):
        definition.Clockwise = False
        assert helix.ModifyDefinition(definition, doc, create_empty_dispatch_variant())
        assert doc.ForceRebuild3(False)
    props = g(doc.Extension, "CustomPropertyManager", "")
    for key, value in [("图号", code), ("名称", name), ("材料", "65Mn"),
                       ("端部", "Y1，两端并紧，每段磨3/4圈；端圈用模板变节距及磨平表达"),
                       ("热处理要求", "45–50 HRC")]:
        props.Add3(key, 30, value, 2)
    save(doc, part)
    print(f"{name} 已保存，检查变节距驱动和外廓", flush=True)
    sw.CloseDoc(g(doc, "GetTitle"))
    doc = open_part(sw, part)
    assert doc.ForceRebuild3(False)
    body, = g(doc, "GetBodies2", 0, False)
    spans = sorted(exact_body_bounds_mm(body))
    check_spans = all(math.isclose(a, b, abs_tol=.02) for a, b in zip(spans, sorted([height, mean+wire, mean+wire])))
    expected_dims = {
        "D3@扫描1": wire, "D1@草图94": height, "D2@草图94": mean+wire,
        "D30003@弹簧": mean, "D40003@弹簧": pitch, "D40002@弹簧": pitch,
    }
    actual_dims = {key: g(doc.Parameter(key), "SystemValue")*1000 for key in expected_dims}
    params_ok = all(math.isclose(actual_dims[k], v, abs_tol=.001) for k, v in expected_dims.items())
    helix = doc.FeatureByName("弹簧")
    definition = g(helix, "GetDefinition")
    clockwise = bool(g(definition, "Clockwise"))
    health = collect_model_summary(doc)
    warnings = health["warning_features"] + health["faulty_features"] + health["unknown_feature_health"]
    turns = g(doc.Parameter("D20005@弹簧"), "SystemValue")
    turns_ok = math.isclose(turns, total, abs_tol=1e-8)
    status = check_spans and params_ok and not clockwise and turns_ok and not warnings and not health.get("feature_error")
    info = {"started": started, "time": datetime.now().astimezone().isoformat(), "part": part.name,
            "figure": code, "passed": status, "size_mm": spans, "clockwise": clockwise, "turns_raw": turns,
            "target_turns": total, "active_turns": active, "dimensions_mm": actual_dims,
            "equations": read_equations(*equations(doc)), "warnings": warnings}
    (out / f"{code}-review.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(code, "检查", status, "外廓", spans, "节距", actual_dims["D40003@弹簧"], "总圈", turns, "Clockwise", clockwise, flush=True)
    assert status, "弹簧驱动尺寸或外廓不符，保留已保存零件"
    assert export_to_step(doc, str(part.with_suffix(".step")))
    save_preview(doc, str(out / f"{code}-spring.bmp"), "isometric")
    save(doc, part)
    sw.CloseDoc(g(doc, "GetTitle"))
    assert g(sw, "GetFirstDocument") is None


if __name__ == "__main__":
    output, template = [Path(v).resolve() for v in sys.argv[1:3]]
    pid = int(sys.argv[3])
    state = solidworks_processes()
    app = _attach_existing_with_pid_check(state, state[pid])
    resume = sys.argv[4:] == ["--resume"]
    assert g(app, "GetProcessID") == pid and (resume or g(app, "GetFirstDocument") is None)
    before = hashlib.sha256(template.read_bytes()).hexdigest()
    build(app, template, output, "SDF-07", "钢球压紧弹簧", .7, 4.5, 7, 5, 9.4, 1.6, resume=resume)
    build(app, template, output, "SDF-08", "压紧弹簧", .5, 3, 4, 2, 3, 1)
    assert hashlib.sha256(template.read_bytes()).hexdigest() == before
