"""第8届手动换向阀题面自绘螺钉，按序号分别创建并重开审查。"""
import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import (
    _ensure_sketch_selected, circle_cut, dimension_existing_relations,
    member as g, mm, round_edges, save_document, sketch, sketch_line,
)
from shared.single_part import anchor, circle, drive, finish_review, prepare_review, quiet_dimensions, relation
from sw_connect import _attach_existing_with_pid_check, solidworks_processes
from sw_part import extrude_cut, fillet, sketch_arc
from sw_sketch_quality import inspect_sketch


NAMES = {"m5": "内六角螺钉_M5x8", "m4": "十字槽螺钉_M4x20"}


def make_feature(feature, name):
    if feature is None:
        raise RuntimeError(f"特征创建失败：{name}")
    feature.Name = name


def draw_profile(doc, points, *, crown=False, orthogonal=True):
    """点为当前草图局部坐标，单位毫米。主回转轴沿模型Y。"""
    manager = doc.SketchManager
    old = manager.AddToDB
    manager.AddToDB = True
    edges = []
    try:
        for i, (a, b) in enumerate(zip(points, points[1:] + points[:1])):
            if i == 0 and crown:
                edge = sketch_arc(doc, 0, mm(6.5), mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
            else:
                edge = sketch_line(doc, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
            if edge is None:
                raise RuntimeError("截面或轮廓线段创建失败")
            edges.append(edge)
    finally:
        manager.AddToDB = old
    for i, (a, b) in enumerate(zip(points, points[1:] + points[:1])):
        edge = edges[i]
        relation(doc, [g(edge, "GetEndPoint2"), g(edges[(i + 1) % len(edges)], "GetStartPoint2")],
                 "sgCOINCIDENT")
        if orthogonal and a[0] == b[0]:
            relation(doc, [edge], "sgVERTICAL")
        elif orthogonal and a[1] == b[1]:
            relation(doc, [edge], "sgHORIZONTAL")
    return edges


def verify_profile(points, edges):
    for expected, edge in zip(points, edges):
        start = g(edge, "GetStartPoint2")
        actual = g(start, "X") * 1000, g(start, "Y") * 1000
        if math.dist(expected, actual) > .002:
            raise RuntimeError(f"草图尺寸回读不一致：{expected} / {actual}")


def define_profile(doc, points, edges):
    dimension_existing_relations(doc)
    verify_profile(points, edges)


def round_head(doc, axial, radius, rounding):
    doc.ClearSelection2(True)
    options = [(edge, center) for edge, center, r in round_edges(doc)
               if abs(radius - r) < .002 and math.dist(center, (0, axial, 0)) < .002]
    assert len(options) == 1, f"螺钉头图示圆边不唯一：{options}"
    data = g(doc.SelectionManager, "CreateSelectData")
    assert options[0][0].Select4(False, data)
    make_feature(fillet(doc, mm(rounding)), "头部过渡")
    doc.ClearSelection2(True)


def body(doc, kind):
    head, stick, length, rounding = (4.25, 2.5, 13, .3) if kind == "m5" else (4, 2, 23.1, .4)
    height = 5 if kind == "m5" else 3.1
    crown = kind == "m4"
    axial = 6.5 - math.sqrt(6.5 ** 2 - head ** 2) if crown else 0
    end = [(stick, length)] if crown else [(stick, length - .5), (stick - .5, length)]
    points = [(0, 0), (head, axial), (head, height), (stick, height), *end, (0, length)]
    with sketch(doc, "Front Plane") as ref:
        edges = draw_profile(doc, points, crown=crown)
        anchor(doc, g(edges[0], "GetStartPoint2"), x=True, y=True)
        if crown:
            anchor(doc, g(edges[0], "GetCenterPoint2"), x=True)
        define_profile(doc, points, edges)
    g(doc, "FeatureByName", str(ref)).Name = "螺钉轴向轮廓"
    _ensure_sketch_selected(doc, "螺钉轴向轮廓")
    sd = g(doc.SelectionManager, "CreateSelectData")
    sd.Mark = 16
    assert edges[-1].Select4(True, sd)
    revolve = doc.FeatureManager.FeatureRevolve2(
        True, True, False, False, False, False, 0, 0, math.tau, 0,
        False, False, 0., 0., 0, 0., 0., True, True, True,
    )
    make_feature(revolve, "回转主体")
    round_head(doc, axial, head, rounding)


def socket(doc, sw, kind):
    if kind == "m5":
        radius = 4.1 / math.sqrt(3)
        points = [(radius * math.cos(math.radians(30 + 60 * i)),
                   radius * math.sin(math.radians(30 + 60 * i))) for i in range(6)]
        with sketch(doc, "Top Plane") as ref:
            active = g(doc.SketchManager, "ActiveSketch")
            axis = g(g(g(active, "ModelToSketchTransform"), "Inverse"), "ArrayData")[7]
            assert abs(abs(axis) - 1) < 1e-8
            edges = draw_profile(doc, points, orthogonal=False)
            polygon_circle = circle(doc, 0.0, 0.0, mm(radius))
            polygon_circle.ConstructionGeometry = True
            anchor(doc, g(polygon_circle, "GetCenterPoint2"), x=True, y=True)
            for edge in edges:
                relation(doc, [g(edge, "GetStartPoint2"), polygon_circle], "sgCOINCIDENT")
            for edge in edges[1:]:
                relation(doc, [edges[0], edge], "sgSAMELENGTH")
            anchor(doc, g(edges[1], "GetStartPoint2"), x=True)
            drive(doc, [polygon_circle], (.003, 0, .003), mm(2 * radius), kind="diameter")
            quality = inspect_sketch(active)
            assert quality["status"] == "pass", f"正六角未完全定义：{quality}"
            verify_profile(points, edges)
        g(doc, "FeatureByName", str(ref)).Name = "内六角对边4点1"
        make_feature(extrude_cut(doc, "内六角对边4点1", mm(2.5), direction=axis > 0), "内六角凹槽")
        return

    # 原剖视：底端宽2、槽深2.3、两斜壁夹角53°；两正交切除的厚度均为1。
    depth, start = 2.3, -.1
    mouth = 1 + (depth - start) * math.tan(math.radians(53 / 2))
    section = [(-mouth, start), (mouth, start), (1, depth), (-1, depth)]
    for plane, name in [("Front Plane", "十字槽横臂"), ("Right Plane", "十字槽纵臂")]:
        with sketch(doc, plane) as ref:
            edges = draw_profile(doc, section, orthogonal=False)
            manager = doc.SketchManager
            old = manager.AddToDB
            manager.AddToDB = True
            try:
                datum = manager.CreateCenterLine(0., 0., 0., 0., mm(depth), 0.)
            finally:
                manager.AddToDB = old
            assert datum is not None
            anchor(doc, g(datum, "GetStartPoint2"), x=True, y=True)
            relation(doc, [datum], "sgVERTICAL")
            relation(doc, [g(datum, "GetEndPoint2"), edges[2]], "sgCOINCIDENT")
            for edge in (edges[0], edges[2]):
                relation(doc, [g(edge, "GetStartPoint2"), g(edge, "GetEndPoint2"), datum], "sgSYMMETRIC")
            define_profile(doc, section, edges)
        g(doc, "FeatureByName", str(ref)).Name = name + "截面"
        make_feature(extrude_cut(doc, name + "截面", mm(.5), both_directions=True), name)
    circle_cut(doc, sw, "Top Plane", [(0, 0, 0)], .85, depth, 1, "十字槽中心")


def checkpoint(doc, target, stage):
    assert doc.ForceRebuild3(False) and len(g(doc, "GetBodies2", 0, False)) == 1
    if not g(doc, "GetPathName"):
        assert save_document(doc, str(target))
    else:
        from shared.single_part import save
        save(doc, target)
    doc.ShowNamedView2("*Isometric", 7)
    doc.ViewZoomtofit2()
    doc.GraphicsRedraw2()
    with (target.parent / "stages.jsonl").open("a", encoding="utf-8") as output:
        output.write(json.dumps({"time": datetime.now().astimezone().isoformat(),
                                 "stage": stage}, ensure_ascii=False) + "\n")
    print(stage, flush=True)


def geometry(doc, rows, kind):
    height, radius, stick, length = (5, 4.25, 2.5, 13) if kind == "m5" else (3.1, 4, 2, 23.1)
    depth, tip = (2.5, 12.5) if kind == "m5" else (2.3, 23.1)

    def near(a, b):
        return abs(a - b) < .01

    def outer_cylinder(row, r):
        p = row.get("Cylinder")
        return p and near(p[6] * 1000, r) and near(abs(p[4]), 1) and near(p[0], 0) and near(p[2], 0)

    check = {
        "头部直径及止口": any(outer_cylinder(s, radius) and near(s["box"][4], height) for s in rows),
        "杆径及长度": any(outer_cylinder(s, stick) and near(s["box"][1], height)
                       and near(s["box"][4], tip) for s in rows),
        "总长度": near(g(doc, "GetPartBox", True)[4] * 1000, length),
        "驱动凹槽未穿头": any((p := s.get("Plane")) and near(abs(p[1]), 1)
                           and near(s["box"][1], depth) and near(s["box"][4], depth)
                           and max(abs(s["box"][i]) for i in (0, 2, 3, 5)) <= 2.5 for s in rows),
    }
    if kind == "m5":
        walls = [s for s in rows if (p := s.get("Plane")) and abs(p[1]) < 1e-8 and
                 near(s["box"][1], 0) and near(s["box"][4], depth)]
        check["六条内六角侧壁及4点1对边"] = len(walls) == 6 and sum(
            near(s["box"][3], 2.05) and near(s["box"][0], 2.05)
            or near(s["box"][3], -2.05) and near(s["box"][0], -2.05) for s in walls) == 2
    else:
        check["R6点5球冠"] = any((p := s.get("Sphere")) and near(p[-1] * 1000, 6.5) for s in rows)
        check["十字槽四个斜端壁"] = sum(
            near(abs(p[1]), math.sin(math.radians(53 / 2))) and near(s["box"][4], depth)
            for s in rows if (p := s.get("Plane"))
        ) == 4
    return check


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("kind", choices=NAMES)
    parser.add_argument("directory", type=Path)
    parser.add_argument("pid", type=int)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--visual-notes", help="已目视核对的简述。仅用于结束当前已完成重开检查的文档")
    mode.add_argument("--resume-body", action="store_true", help="从已重开的干净回转主体保存点续建凹槽")
    args = parser.parse_args()
    out = args.directory.resolve()
    target = out / (NAMES[args.kind] + ".SLDPRT")
    assert out.is_dir()
    state = solidworks_processes()
    sw = _attach_existing_with_pid_check(state, state[args.pid])
    doc = g(sw, "ActiveDoc")
    assert g(sw, "GetProcessID") == args.pid and doc is not None and g(doc, "GetType") == 1
    if args.visual_notes:
        finish_review(sw, target, visual_notes=args.visual_notes)
        return
    assert not target.with_suffix(".step").exists() and not (out / "verification.json").exists()
    if args.resume_body:
        assert target.is_file() and Path(g(doc, "GetPathName")).resolve() == target
        assert not g(doc, "GetSaveFlag") and g(doc.SketchManager, "ActiveSketch") is None
        stages = (out / "stages.jsonl").read_text(encoding="utf-8").splitlines()
        assert len(stages) == 1 and json.loads(stages[0])["stage"] == "回转主体及头部过渡"
        assert len(g(doc, "GetBodies2", 0, False)) == 1
        assert g(doc, "FeatureByName", "回转主体") and g(doc, "FeatureByName", "头部过渡")
    else:
        assert not target.exists() and not g(doc, "GetPathName") and not g(doc, "GetBodies2", 0, False)
    with quiet_dimensions(sw):
        if not args.resume_body:
            body(doc, args.kind)
            checkpoint(doc, target, "回转主体及头部过渡")
        socket(doc, sw, args.kind)
    props = g(doc.Extension, "CustomPropertyManager", "")
    kind = args.kind
    for key, value in [
        ("名称", NAMES[kind]),
        ("材料", "45"),
        ("表面处理", "发黑" if kind == "m5" else "镀锌"),
        ("螺纹", "M5×0.8，图示螺纹段5.6" if kind == "m5" else "M4×0.7，图示螺纹段18.6"),
        ("表达范围", "原卷允许主要结构表达，螺杆采用公称直径"),
    ]:
        props.Add3(key, 30, value, 2)
    checkpoint(doc, target, "驱动凹槽完成")
    report = prepare_review(sw, target, lambda model, rows: geometry(model, rows, kind),
                            sketch_count=2 if kind == "m5" else 4, views=["isometric", "bottom", "front_wire"])
    if not report["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
