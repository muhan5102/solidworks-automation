"""第8届SDF-10配气盘垫：原图尺寸及配气盘定位接口。"""
import json
import math
from datetime import datetime
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, sketch, dimension_existing_relations, save_document
from shared.single_part import anchor, circle, quiet_dimensions, relation, prepare_review
from sw_connect import _attach_existing_with_pid_check, solidworks_processes
from sw_part import extrude_boss, extrude_cut

out, pid = Path(sys.argv[1]).resolve(), int(sys.argv[2])
target = out / "配气盘垫_SDF-10.SLDPRT"
state = solidworks_processes()
sw = _attach_existing_with_pid_check(state, state[pid])
m = g(sw, "ActiveDoc")
assert g(sw, "GetProcessID") == pid and m is not None and g(m, "GetType") == 1
assert not g(m, "GetPathName") and not g(m, "GetBodies2", 0, False)


def checkpoint(stage):
    assert m.ForceRebuild3(False) and len(g(m, "GetBodies2", 0, False)) == 1
    assert save_document(m, str(target))
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    with (out / "stages.jsonl").open("a", encoding="utf-8") as log:
        log.write(json.dumps({"time": datetime.now().astimezone().isoformat(),
                              "stage": stage}, ensure_ascii=False) + "\n")
    print(stage, flush=True)


def finish_sketch(ref, name):
    m.FeatureByName(str(ref)).Name = name
    return name


with quiet_dimensions(sw):
    with sketch(m, "Front Plane") as ref:
        c = circle(m, 0, 0, .017)
        anchor(m, g(c, "GetCenterPoint2"), x=True, y=True)
        dimension_existing_relations(m)
    f = extrude_boss(m, finish_sketch(ref, "垫体轮廓"), .003, direction=False)
    assert f is not None
    f.Name = "垫体"
    box = g(m, "GetPartBox", True)
    assert abs(box[2]) < 1e-8 and abs(box[5] - .003) < 1e-8
    checkpoint("Φ34垫体厚3")

    m.ClearSelection2(True)
    assert m.Extension.SelectByRay(0, 0, .006, 0, 0, -1, .00001, 2, False, 0, 0)
    manager = m.SketchManager
    manager.InsertSketch(True)
    sk = g(manager, "ActiveSketch")
    matrix = g(g(sk, "ModelToSketchTransform"), "ArrayData")
    normal_z = g(g(g(sk, "ModelToSketchTransform"), "Inverse"), "ArrayData")[8]
    assert abs(abs(normal_z) - 1) < 1e-8
    for x in (-.011, .011):
        xyz = (x, 0, .003)
        a, b, c = [sum(xyz[j] * matrix[j*3+i] for j in range(3)) * matrix[12] + matrix[9+i] for i in range(3)]
        assert abs(c) < 1e-8
        made = circle(m, a, b, .0015)
        anchor(m, g(made, "GetCenterPoint2"), x=abs(a)<1e-8, y=abs(b)<1e-8)
    dimension_existing_relations(m)
    ref = g(sk, "Name")
    manager.InsertSketch(True)
    f = extrude_boss(m, finish_sketch(ref, "定位柱轮廓"), .002, direction=normal_z < 0)
    assert f is not None
    f.Name = "两定位柱"

    # 槽面从Z负侧观察；实体坐标孔群在Y正侧，定位柱位于X轴。
    with sketch(m, "Front Plane") as ref:
        side = .011 / math.sqrt(2)
        for x in (-side, side):
            circle(m, x, side, .002)
        dimension_existing_relations(m)
    f = extrude_cut(m, finish_sketch(ref, "两个Φ4通孔轮廓"), 0, direction=True)
    assert f is not None
    f.Name = "两个Φ4通孔"

    with sketch(m, "Front Plane") as ref:
        arcmgr = m.SketchManager
        for start in (45, 225):
            end = start + 90

            def point(r, angle):
                a = math.radians(angle)
                return r*math.cos(a), r*math.sin(a)

            outer_a, outer_b = point(.013, start), point(.013, end)
            inner_a, inner_b = point(.009, start), point(.009, end)
            a, b = point(.011, start), point(.011, end)
            spec = [((0, 0), outer_a, outer_b, 1), (b, outer_b, inner_b, 1),
                    ((0, 0), inner_b, inner_a, -1), (a, inner_a, outer_a, 1)]
            old = arcmgr.AddToDB
            arcmgr.AddToDB = True
            try:
                arcs = [arcmgr.CreateArc(*c, 0, *p, 0, *q, 0, sense) for c, p, q, sense in spec]
            finally:
                arcmgr.AddToDB = old
            for cur, nxt in zip(arcs, arcs[1:] + arcs[:1]):
                relation(m, [g(cur, "GetEndPoint2"), g(nxt, "GetStartPoint2")], "sgCOINCIDENT")
                relation(m, [cur, nxt], "sgTANGENT")
            relation(m, [arcs[0], arcs[2]], "sgCONCENTRIC")
            anchor(m, g(arcs[0], "GetCenterPoint2"), x=True, y=True)
        dimension_existing_relations(m)
    f = extrude_cut(m, finish_sketch(ref, "双90度弧槽轮廓"), .002, direction=True)
    assert f is not None
    f.Name = "双90度弧槽"

props = g(m.Extension, "CustomPropertyManager", "")
for name, value in [("图号", "SDF-10"), ("名称", "配气盘垫"), ("材料", "耐油橡胶"),
                    ("装配", "与配气盘粘接；定位柱随SDF-09的Φ3H9孔改Φ3f9"),
                    ("题面假设", "弧槽未注深度，取2，保留1厚膜")]:
    props.Add3(name, 30, value, 2)
checkpoint("定位柱、通孔及双弧槽")


def check(model, rows):
    def hits(radius, lo, hi):
        return [s for s in rows if (p := s.get("Cylinder"))
                and abs(p[6]*1000-radius) < .001 and abs(abs(p[5])-1) < 1e-8
                and abs(s["box"][2]-lo) < .001 and abs(s["box"][5]-hi) < .001]

    pins = hits(1.5, 3, 5)
    ends = [s for s in hits(2, 0, 2) if math.isclose(math.hypot(*s["Cylinder"][:2]), .011, abs_tol=1e-6)]
    side = 11/math.sqrt(2)
    passages = [s for s in rows if (p := s.get("Cylinder")) and abs(p[6]-.002)<1e-8
                and abs(abs(p[0]*1000)-side)<.001 and abs(p[1]*1000-side)<.001
                and abs(s["box"][5]-3)<.001]
    return {"Φ34厚3": bool(hits(17, 0, 3)),
            "两个定位柱厚2及分布圆Φ22": len(pins)==2 and all(abs(abs(s["Cylinder"][0])-.011)<1e-8 and abs(s["Cylinder"][1])<1e-8 for s in pins),
            "两个Φ4孔正向且贯通余膜": len(passages)==2 and all(
                min(abs(s["box"][2]), abs(s["box"][2]-2))<.001 for s in passages),
            "双弧槽内外壁及槽深2": len(hits(9,0,2))==2 and len(hits(13,0,2))==2 and bool(ends)}


prepare_review(sw, target, check, sketch_count=4, views=["isometric", "back", "front_wire"])
