"""LJT05.01拨叉：开口叉体、连接杆、圆毂、筋与单壁径向沉头孔。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.single_part import g, relation as rel0, anchor as anchor0, drive as drive0, quiet_dimensions, circle as circle0
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import mm, sketch, sketch_line, new_document, save_document, dimension_existing_relations, offset_plane, chamfer_edges
from sw_part import extrude_boss, extrude_cut, _ensure_sketch_selected
from sw_connect import create_empty_dispatch_variant
from sw_review import save_preview
from rib_fork import create_rib
from round_fork import finish_rounds

out = Path(sys.argv[1]).resolve()
target = out/("拨叉_路线重放.SLDPRT" if "--replay" in sys.argv else "拨叉_LJT05.01.SLDPRT")
resume = "--resume" in sys.argv[2:]
assert target.exists() if resume else not target.exists()
s = win32com.client.Dispatch("SldWorks.Application.32")
assert g(s, "GetProcessID") == 12356
assert resume or g(s, "GetFirstDocument") is None
s.Visible = True
m = g(s, "ActiveDoc") if resume else new_document(s, "part")
if resume:
    assert Path(g(m, "GetPathName")).resolve() == target
timer = time.perf_counter()
def rel(items, kind): rel0(m, items, kind)
def anchor(p, **kwargs): anchor0(m, p, **kwargs)
def drive(items, pos, value, kind="normal"): return drive0(m, items, pos, value, kind)
def poly(points):
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        ee = [sketch_line(m, *[mm(v) for v in a], *[mm(v) for v in b])
              for a, b in zip(points, points[1:]+points[:1])]
    finally: sm.AddToDB = old
    for a, b in zip(ee, ee[1:]+ee[:1]):
        rel([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
    for e, a, b in zip(ee, points, points[1:]+points[:1]):
        if a[0] == b[0] or a[1] == b[1]:
            rel([e], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    return ee
def origin_dim(point, pos, value, kind):
    m.ClearSelection2(True)
    assert point.Select4(False, g(m.SelectionManager, "CreateSelectData"))
    assert m.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, create_empty_dispatch_variant(), 0)
    method = m.AddHorizontalDimension2 if kind == "horizontal" else m.AddVerticalDimension2
    dd = method(*pos)
    assert dd is not None and abs(g(g(dd, "GetDimension2", 0), "SystemValue")-value) < 1e-8
    assert g(g(dd, "GetDimension2", 0), "DrivenState") == 2
def checkpoint(stage):
    assert m.ForceRebuild3(False) and save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    row = {"stage": stage, "saved": datetime.now().astimezone().isoformat(),
           "execution_seconds": time.perf_counter()-timer, "file": target.name}
    (out/"checkpoint.json").write_text(json.dumps(row, ensure_ascii=False, indent=2), encoding="utf-8")
    with (out/"stages.jsonl").open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(row, ensure_ascii=False)+"\n")
    print(f"[{stage}] 已保存，脚本累计{row['execution_seconds']:.2f}秒", flush=True)

with quiet_dimensions(s):
    if not resume:
        # 平面X-Z：开口端在X1，避免将图示1误作完整半圆。
        inner_z = math.sqrt(13.5**2-1)
        outer_x = math.sqrt(27**2-22**2)
        pts = [(1, 22), (outer_x, 22), (outer_x, -22), (1, -22),
               (1, -inner_z), (1, inner_z)]
        with sketch(m, "Top Plane") as ref:
            sm = m.SketchManager
            inv = list(g(g(g(sm, "ActiveSketch"), "ModelToSketchTransform"), "Inverse").ArrayData)
            old = sm.AddToDB
            sm.AddToDB = True
            try:
                ee = []
                for i, (a, b) in enumerate(zip(pts, pts[1:]+pts[:1])):
                    if i in [1, 4]:
                        e = sm.CreateArc(0, 0, 0, mm(a[0]), mm(a[1]), 0, mm(b[0]), mm(b[1]), 0, -1 if i == 1 else 1)
                    else:
                        e = sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                    assert e is not None
                    ee.append(e)
            finally: sm.AddToDB = old
            for a, b in zip(ee, ee[1:]+ee[:1]):
                rel([g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
            for i in [0, 2]: rel([ee[i]], "sgHORIZONTAL")
            for i in [3, 5]: rel([ee[i]], "sgVERTICAL")
            anchor(g(ee[1], "GetCenterPoint2"), x=True, y=True)
            anchor(g(ee[4], "GetCenterPoint2"), x=True, y=True)
            rel([g(ee[0], "GetStartPoint2"), g(ee[2], "GetEndPoint2")], "sgVERTICALPOINTS2D")
            drive([ee[1]], (.030, 0, -.025), .027)
            drive([ee[4]], (.010, 0, -.014), .0135)
            drive([g(ee[0], "GetStartPoint2"), g(ee[2], "GetEndPoint2")], (0, 0, 0), .044, "vertical")
            origin_dim(g(ee[0], "GetStartPoint2"), (0, 0, -.030), .001, "horizontal")
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name = "叉口轮廓"
        f = extrude_boss(m, "叉口轮廓", .010, direction=inv[7] < 0)
        assert f is not None
        f.Name = "叉体"

        arm_plane = offset_plane(m, "Top Plane", 2, "杆下表面")
        with sketch(m, arm_plane) as ref:
            ee = poly([(13.5, -15), (120, -15), (120, 15), (13.5, 15)])
            origin_dim(g(ee[1], "GetStartPoint2"), (.120, .002, .025), .120, "horizontal")
            drive([ee[1]], (.130, .002, 0), .030)
            dimension_existing_relations(m)
            inv = list(g(g(g(m.SketchManager, "ActiveSketch"), "ModelToSketchTransform"), "Inverse").ArrayData)
        m.FeatureByName(str(ref)).Name = "连接杆轮廓"
        f = extrude_boss(m, "连接杆轮廓", .006, direction=inv[7] < 0)
        assert f is not None
        f.Name = "连接杆"
    with sketch(m, "Top Plane") as ref:
        circle = circle0(m, .120, 0, .015)
        p = g(circle, "GetCenterPoint2")
        anchor(p, y=True)
        origin_dim(p, (.060, 0, .025), .120, "horizontal")
        drive([circle], (.140, 0, .015), .030, "diameter")
        dimension_existing_relations(m)
        inv = list(g(g(g(m.SketchManager, "ActiveSketch"), "ModelToSketchTransform"), "Inverse").ArrayData)
    m.FeatureByName(str(ref)).Name = "圆毂轮廓"
    f = extrude_boss(m, "圆毂轮廓", .028, direction=inv[7] < 0)
    assert f is not None
    f.Name = "圆毂"
    create_rib(s, m)
    checkpoint("叉体杆圆毂和加强筋")
    save_preview(m, str(out/"结构阶段.bmp"), "isometric")
    # 圆毂统一开孔，清除杆、筋与圆毂合并时进入孔腔的材料。
    bore_plane = offset_plane(m, "Top Plane", 28, "圆毂端面")
    with sketch(m, bore_plane) as ref:
        circle = circle0(m, .120, 0, .0075)
        p = g(circle, "GetCenterPoint2")
        anchor(p, y=True)
        origin_dim(p, (.060, .028, .020), .120, "horizontal")
        drive([circle], (.140, .028, .010), .015, "diameter")
        dimension_existing_relations(m)
        normal = list(g(g(g(m.SketchManager, "ActiveSketch"), "ModelToSketchTransform"), "Inverse").ArrayData)[7]
    m.FeatureByName(str(ref)).Name = "圆毂孔位"
    f = extrude_cut(m, "圆毂孔位", .028, direction=normal*(-1) > 0)
    assert f is not None
    f.Name = "圆毂通孔"
    chamfer_edges(m, [([120, 0, 0], 7.5), ([120, 28, 0], 7.5)], 1, "图示孔口C1")
    # 曲面入口：Φ11口在外圆交线上，口面Z=sqrt(15²-5.5²)，锥段轴向3。
    mouth_z = math.sqrt(15**2-5.5**2)
    neck_z = mouth_z-3
    radial_plane = offset_plane(m, "Right Plane", 120, "径向孔轴截面")
    world = [(120, 14, 0), (120, 16.5, 0), (120, 16.5, neck_z),
             (120, 19.5, mouth_z), (120, 19.5, 20), (120, 14, 20)]
    with sketch(m, radial_plane) as ref:
        active = g(m.SketchManager, "ActiveSketch")
        matrix = list(g(g(active, "ModelToSketchTransform"), "ArrayData"))
        def local(point):
            p = [mm(v) for v in point]
            xyz = [sum(p[j]*matrix[j*3+i] for j in range(3))*matrix[12]+matrix[9+i] for i in range(3)]
            assert abs(xyz[2]) < 1e-7
            return [xyz[0]*1000, xyz[1]*1000]
        points = [local(p) for p in world]
        ee = poly(points)
        a, b = points[0], points[-1]
        axis = m.SketchManager.CreateCenterLine(mm(a[0]), mm(a[1]), 0, mm(b[0]), mm(b[1]), 0)
        rel([g(axis, "GetStartPoint2"), g(ee[0], "GetStartPoint2")], "sgCOINCIDENT")
        rel([g(axis, "GetEndPoint2"), g(ee[-1], "GetStartPoint2")], "sgCOINCIDENT")
        anchor(g(axis, "GetStartPoint2"), x=True)
        origin_dim(g(axis, "GetStartPoint2"), (.120, .014, -.004), .014, "vertical")
        drive([ee[2], axis], (.120, .016, .012), math.pi/4)
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name = "径向沉头孔截面"
    _ensure_sketch_selected(m, "径向沉头孔截面")
    data = g(m.SelectionManager, "CreateSelectData")
    data.Mark = 16
    assert axis.Select4(True, data)
    f = m.FeatureManager.FeatureRevolve2(True, True, False, True, False, False, 0, 0,
        2*math.pi, 0, False, False, 0, 0, 0, 0, 0, True, False, True)
    assert f is not None
    f.Name = "径向沉头孔"
    for name in ["杆下表面", "圆毂端面", "径向孔轴截面"]:
        m.ClearSelection2(True)
        assert m.FeatureByName(name).Select2(False, 0)
        m.BlankRefGeom()
    m.SetMaterialPropertyName2("", "", "")
    for key, value in [("名称", "拨叉"), ("图号", "LJT05.01"), ("材料", "HT200"),
                       ("技术要求", "未注倒角C1；未注圆角R2～R4；未注尺寸公差GB/T 1804—2000-m")]:
        m.Extension.CustomPropertyManager("").Add3(key, 30, value, 2)
    finish_rounds(m)
    checkpoint("孔系与图示过渡完成")
(out/"build.json").write_text(json.dumps({
    "finished": datetime.now().astimezone().isoformat(), "execution_seconds": time.perf_counter()-timer
}, ensure_ascii=False, indent=2), encoding="utf-8")
