"""王静LJT04.23调节盘；图面u=X,v=-Z，从法兰端Y64观察。"""
import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import (member, mm, sketch, sketch_line, new_document,
    save_document, revolve_boss, dimension_existing_relations, offset_plane, circle_cut)
from sw_part import extrude_cut
from sw_connect import create_empty_dispatch_variant, open_document
from sw_pattern import circular_pattern
from sw_sketch_quality import inspect_model_sketches, fully_define_sketch
from sw_review import collect_model_summary, save_preview
import win32com.client

parser = argparse.ArgumentParser()
parser.add_argument("output", type=Path)
parser.add_argument("--process-id", type=int, required=True)
args = parser.parse_args()
out = args.output.resolve()
dest = out / "调节盘_LJT04.23.SLDPRT"
assert out.is_dir() and not dest.exists()
started = time.monotonic()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == args.process_id
sw.Visible = True
sw.UserControl = True
m = new_document(sw, "part")


def checkpoint(stage):
    assert m.ForceRebuild3(False)
    assert save_document(m, str(dest))
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    (out / "checkpoint.json").write_text(json.dumps({
        "stage": stage, "path": str(dest), "pid": args.process_id,
        "time": datetime.now().astimezone().isoformat(),
        "elapsed_seconds": time.monotonic() - started,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[{stage}] 已重建并保存，累计{time.monotonic()-started:.1f}秒", flush=True)


# 截面直接包含内环槽、外退刀槽和三处图示C2；不额外选边倒角。
points = [(32, 0), (32, 11), (37, 11), (37, 24), (30, 24),
          (30, 56), (32, 58), (52, 58), (52, 64),
          (88, 64), (90, 62), (90, 42), (49, 42), (49, 37),
          (50, 37), (50, 2), (48, 0)]
with sketch(m, "Front Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0, 0, 0, 0, mm(64), 0)
        lines = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                 for a, b in zip(points, points[1:] + points[:1])]
    finally:
        sm.AddToDB = old
    assert axis is not None and all(x is not None for x in lines)
    sd = member(m.SelectionManager, "CreateSelectData")
    for a, b in zip(lines, lines[1:] + lines[:1]):
        m.ClearSelection2(True)
        assert member(a, "GetEndPoint2").Select4(False, sd)
        assert member(b, "GetStartPoint2").Select4(True, sd)
        m.SketchAddConstraints("sgCOINCIDENT")
    for line, a, b in zip(lines, points, points[1:] + points[:1]):
        relation = "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL" if a[1] == b[1] else None
        if relation:
            m.ClearSelection2(True)
            assert line.Select4(False, sd)
            m.SketchAddConstraints(relation)
    m.ClearSelection2(True)
    assert axis.Select4(False, sd)
    m.SketchAddConstraints("sgVERTICAL")
    m.ClearSelection2(True)
    assert member(axis, "GetStartPoint2").Select4(False, sd)
    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0,
                                  create_empty_dispatch_variant(), 0)
    m.SketchAddConstraints("sgCOINCIDENT")
    m.ClearSelection2(True)
    assert member(lines[0], "GetStartPoint2").Select4(False, sd)
    assert member(axis, "GetStartPoint2").Select4(True, sd)
    m.SketchAddConstraints("sgHORIZONTALPOINTS2D")
    dimension_existing_relations(m)
    for line, expected in zip(lines, points):
        p = member(line, "GetStartPoint2")
        assert math.dist([member(p, "X")*1000, member(p, "Y")*1000], expected) < 1e-5
boss = revolve_boss(m, ref, 2*math.pi)
assert boss is not None
boss.Name = "主体回转"
m.FeatureByName(str(ref)).Name = "轴向截面"
checkpoint("主体")

plane = offset_plane(m, "Top Plane", 64, "法兰前端面")
with sketch(m, plane) as slot_ref:
    active = member(m.SketchManager, "ActiveSketch")
    transform = list(member(member(active, "ModelToSketchTransform"), "ArrayData"))
    inverse = list(member(member(member(active, "ModelToSketchTransform"), "Inverse"), "ArrayData"))
    normal_y = inverse[7]

    def xy(u, v):
        xyz = [mm(u), mm(64), mm(-v)]
        p = [sum(xyz[j]*transform[j*3+i] for j in range(3))*transform[12]+transform[9+i]
             for i in range(3)]
        assert abs(p[2]) < 1e-8
        return p[:2]

    def polar(r, deg):
        a = math.radians(deg)
        return r*math.cos(a), r*math.sin(a)

    # 沿边界依次走：外圆弧、上端盖、内圆弧、下端盖。
    c0, c1 = polar(71, 30), polar(71, 60)
    specs = [((0, 0), polar(78, 30), polar(78, 60), 1),
             (c1, polar(78, 60), polar(64, 60), 1),
             ((0, 0), polar(64, 60), polar(64, 30), -1),
             (c0, polar(64, 30), polar(78, 30), 1)]
    origin, ex, ey = xy(0, 0), xy(1, 0), xy(0, 1)
    orientation = 1 if ((ex[0]-origin[0])*(ey[1]-origin[1])-
                        (ex[1]-origin[1])*(ey[0]-origin[0])) > 0 else -1
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        arcs = []
        for center, start, end, direction in specs:
            c, a, b = xy(*center), xy(*start), xy(*end)
            arcs.append(sm.CreateArc(c[0], c[1], 0, a[0], a[1], 0,
                                     b[0], b[1], 0, direction*orientation))
    finally:
        sm.AddToDB = old
    assert all(a is not None for a in arcs)
    sd = member(m.SelectionManager, "CreateSelectData")
    for a, b in zip(arcs, arcs[1:] + arcs[:1]):
        m.ClearSelection2(True)
        assert member(a, "GetEndPoint2").Select4(False, sd)
        assert member(b, "GetStartPoint2").Select4(True, sd)
        m.SketchAddConstraints("sgCOINCIDENT")
    # 同心、相切及尺寸由现有FullyDefineSketch补足，随后回读圆心和半径。
    fully_define_sketch(m)
    for arc, (center, start, end, direction) in zip(arcs, specs):
        cp = member(arc, "GetCenterPoint2")
        expected = xy(*center)
        assert math.dist([member(cp, "X"), member(cp, "Y")], expected) < 1e-8
        assert abs(member(arc, "GetRadius")-mm(math.dist(center, start))) < 1e-8
    contours = list(member(active, "GetSketchContours") or [])
    assert len(contours) == 1 and member(contours[0], "IsClosed")
cut = extrude_cut(m, slot_ref, mm(22), direction=(normal_y*-1 > 0))
assert cut is not None
cut.Name = "母弧槽"
m.FeatureByName(str(slot_ref)).Name = "弧槽草图"
# 回转草图中心线是明确的轴实体，直接复用，避免另建基准轴。
pattern = circular_pattern(m, "母弧槽", axis, 2*math.pi, 6, True)
assert pattern is not None
pattern.Name = "六弧槽阵列"
checkpoint("弧槽")

for angle, diameter, name in [(15, 8, "铰制通孔"), (-45, 10, "销孔")]:
    a = math.radians(angle)
    circle_cut(m, sw, plane, [(71*math.cos(a), 64, -71*math.sin(a))],
               diameter/2, 22, -1, name)
props = m.Extension.CustomPropertyManager("")
for k, v in [
    ("名称", "调节盘"), ("图号", "LJT04.23"), ("材料", "45"),
    ("热处理", "调质230～250HBW；图示指定表面淬硬40～45HRC"),
    ("配合", "Φ100g6、Φ64H7、Φ60H7、Φ8H7通孔、Φ10H7销孔配作"),
    ("技术要求", "未注倒角C2；去毛刺、飞边；未注尺寸公差GB/T1804-2000-m；未注几何公差GB/T1184-1996-K"),
    ("来源", "王静图集PDF84/印刷78，第23件"),
]:
    props.Add3(k, 30, v, 2)
checkpoint("两孔")
sw.CloseDoc(member(m, "GetTitle"))
m = open_document(sw, str(dest), silent=True)
assert m is not None and m.ForceRebuild3(False)
assert str(member(m, "GetPathName")).lower() == str(dest).lower()
quality = inspect_model_sketches(m)
health = collect_model_summary(m)
faces = []
for body in member(m, "GetBodies2", 0, False):
    for face in member(body, "GetFaces"):
        surface = member(face, "GetSurface")
        item = {"box_m": list(member(face, "GetBox")), "area_m2": member(face, "GetArea")}
        for kind, test, params in [("cylinder", "IsCylinder", "CylinderParams"),
                                   ("cone", "IsCone", "ConeParams"),
                                   ("plane", "IsPlane", "PlaneParams")]:
            if member(surface, test):
                item[kind] = list(member(surface, params))
        faces.append(item)
report = {"path": str(dest), "faces": faces, "sketches": quality, "health": health,
          "execution_seconds": time.monotonic()-started,
          "finished": datetime.now().astimezone().isoformat()}
(out/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
assert quality["status"] == "pass", quality
assert not health["faulty_features"] and not health["warning_features"], health
if member(m, "GetVisibilityOfConstructPlanes"):
    member(m, "ViewDispRefplanes")
for name, view in [("调节盘_等轴测.bmp", "isometric"), ("调节盘_法兰端.bmp", "top"),
                   ("调节盘_侧视.bmp", "front")]:
    save_preview(m, str(out/name), view)
assert save_document(m)
print(json.dumps({"execution_seconds": report["execution_seconds"], "quality": quality["status"]},
                 ensure_ascii=False), flush=True)
