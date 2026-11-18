"""第8届手动换向阀阀体。先由MCP建立空零件，随后按阶段执行。"""
from contextlib import contextmanager
from datetime import datetime
import json
import math
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import (member as g, chamfer_edges, dimension_existing_relations,
                        offset_plane, save_document, sketch, sketch_line)
from shared.single_part import (anchor, circle, finish_review, prepare_review,
                                quiet_dimensions, relation, surfaces)
from sw_connect import _attach_existing_with_pid_check, solidworks_processes
from sw_part import (_ensure_sketch_selected, _get_plane_name_candidates,
                     _select_first_candidate, extrude_boss, extrude_cut, fillet)
from sw_pattern import circular_pattern


@contextmanager
def face_sketch(model, origin, into):
    """在已知实体面建立草图，提供世界坐标转本地坐标及法向。"""
    model.ClearSelection2(True)
    source = tuple(v - .003*d for v, d in zip(origin, into))
    assert model.Extension.SelectByRay(*source, *into, 1e-6, 2, False, 0, 0)
    manager = model.SketchManager
    manager.InsertSketch(True)
    active = g(manager, "ActiveSketch")
    assert active is not None
    t = g(active, "ModelToSketchTransform")
    data = g(t, "ArrayData")
    normal = g(g(t, "Inverse"), "ArrayData")[6:9]

    def local(p):
        q = [sum(p[j]*data[j*3+i] for j in range(3))*data[12]+data[9+i] for i in range(3)]
        assert abs(q[2]) < 1e-7, ("目标点不在实际草图面", p, q)
        return q[:2]

    try:
        local(origin)
        yield g(active, "Name"), local, normal
    finally:
        if g(manager, "ActiveSketch") is not None:
            manager.InsertSketch(True)


def name_sketch(model, ref, name):
    model.FeatureByName(str(ref)).Name = name
    return name


def flat_box(model):
    box = g(model, "GetPartBox", True)
    assert box and len(box) == 6
    return [round(float(n)*1000, 6) for n in box]


def body(model, checkpoint):
    assert not g(model, "GetPathName") and not g(model, "GetBodies2", 0, False)
    with sketch(model, "Front Plane") as ref:
        assert model.SketchManager.CreateCenterRectangle(0, 0, 0, .0275, .0275, 0)
        dimension_existing_relations(model)
    f = extrude_boss(model, name_sketch(model, ref, "阀体外廓"), .020, direction=False)
    assert f is not None
    f.Name = "55平方基体"
    assert flat_box(model) == [-27.5, -27.5, 0.0, 27.5, 27.5, 20.0], flat_box(model)
    checkpoint("55×55×20方体")

    model.ClearSelection2(True)
    b, = g(model, "GetBodies2", 0, False)
    corners = []
    for edge in g(b, "GetEdges"):
        a, c = g(edge, "GetStartVertex"), g(edge, "GetEndVertex")
        if a is None or c is None:
            continue
        p, q = g(a, "GetPoint"), g(c, "GetPoint")
        if (abs(p[0]-q[0]) < 1e-8 and abs(p[1]-q[1]) < 1e-8
                and abs(abs(p[2]-q[2])-.020) < 1e-8):
            corners.append(edge)
    assert len(corners) == 4
    sd = g(model.SelectionManager, "CreateSelectData")
    for i, edge in enumerate(corners):
        assert edge.Select4(i > 0, sd)
    f = fillet(model, .001)
    assert f is not None
    f.Name = "方体四角R1"

    for seq, (z, diameter, depth) in enumerate(((20,35,1), (21,32.5,2), (23,35,1)), 1):
        origin = (0, 0, z/1000)
        with face_sketch(model, origin, (0,0,-1)) as (ref, local, normal):
            xy = local(origin)
            c = circle(model, *xy, diameter/2000)
            if all(abs(n) < 1e-8 for n in xy):
                anchor(model, g(c, "GetCenterPoint2"), x=True, y=True)
            dimension_existing_relations(model)
        sn = name_sketch(model, ref, f"止口轮廓{seq}")
        f = extrude_boss(model, sn, depth/1000, direction=normal[2] < 0)
        assert f is not None, sn
        f.Name = ("止口根台", "密封槽根径", "止口端台")[seq-1]
        assert abs(flat_box(model)[5] - z - depth) < 1e-5

    rows = surfaces(model)
    for radius, low, high in ((17.5,20,21), (16.25,21,23), (17.5,23,24)):
        assert any("Cylinder" in r and abs(r["Cylinder"][6]*1000-radius)<.001
                   and abs(r["box"][2]-low)<.001 and abs(r["box"][5]-high)<.001 for r in rows)
    props = g(model.Extension, "CustomPropertyManager", "")
    for key, val in (("名称","阀体"), ("图号","SDF-01"), ("材料","LY12"),
                     ("侧口记号","按C-C: 左上O、左下A、右上B、右下P")):
        props.Add3(key, 30, val, 2)
    checkpoint("四角R1及Φ35/Φ32.5止口，主体总高24")


def cut_circles(model, z, xy, diameter, depth, direction, label, *, pick=None):
    origin = (*(xy[0] if pick is None else pick), z)
    with face_sketch(model, origin, (0, 0, direction)) as (ref, local, normal):
        for x, y in xy:
            a, b = local((x, y, z))
            c = circle(model, a, b, diameter/2)
            if abs(a) < 1e-8 or abs(b) < 1e-8:
                anchor(model, g(c, "GetCenterPoint2"), x=abs(a)<1e-8, y=abs(b)<1e-8)
        dimension_existing_relations(model)
    sn = name_sketch(model, ref, label+"轮廓")
    f = extrude_cut(model, sn, depth, direction=normal[2]*direction > 0)
    assert f is not None, label
    f.Name = label
    return f


CORNER_XY = [(i*22, j*22) for i in (-1,1) for j in (-1,1)]
MOUNT_XY = [(20.75,0), (0,20.75), (-20.75,0), (0,-20.75)]
AIR_XY = [(i*11/math.sqrt(2),j*11/math.sqrt(2)) for i in (-1,1) for j in (-1,1)]


def cylinder_match(rows, axis, center_mm, radius, low, high):
    dimension = "xyz".index(axis)
    matches = []
    for row in rows:
        cp = row.get("Cylinder")
        if not cp or abs(abs(cp[dimension+3])-1)>1e-8 or abs(cp[6]*1000-radius)>.001:
            continue
        if any(abs(cp[i]*1000-center_mm[i])>.001 for i in range(3) if i != dimension):
            continue
        if abs(row["box"][dimension]-low)<.001 and abs(row["box"][dimension+3]-high)<.001:
            matches.append(row)
    return matches


def mount_holes(model, checkpoint, wizard=True):
    assert model.FeatureByName("止口端台") is not None
    if model.FeatureByName("四角Φ7贯穿") is None:
        cut_circles(model, 0, [(x/1000,y/1000) for x,y in CORNER_XY], .007, 0, 1, "四角Φ7贯穿")
    if model.FeatureByName("Φ10中心盲窝") is None:
        cut_circles(model, .024, [(0,0)], .010, .002, -1, "Φ10中心盲窝")
    checkpoint("四角孔和中心Φ10深2平底窝")

    candidates = [(x/1000,y/1000) for x,y in MOUNT_XY]
    if wizard:
        model.ClearSelection2(True)
        assert model.Extension.SelectByRay(.02075,0,-.003,0,0,1,1e-6,2,False,0,0)
        f = model.FeatureManager.HoleWizard5(0,13,361,"M4",1,.005,.020,-1,
            .0085,.003,0,-1,-1,0,0,0,0,0,0,-1,"",False,False,True,False,False,False)
        assert f is not None, "GB沉孔向导未创建。仅在确认无残留向导特征后执行 holes-cut"
        f.Name = "底面安装沉孔母孔"
        sub = g(f, "GetFirstSubFeature")
        sub.Name = "底面安装沉孔定位"
        model.ClearSelection2(True)
        assert sub.Select2(False,0)
        model.EditSketch()
        try:
            sk = g(model.SketchManager,"ActiveSketch")
            p, = g(sk,"GetSketchPoints2")
            q = [g(p,k) for k in ("X","Y","Z")]
            if abs(q[0])<1e-9 or abs(q[1])<1e-9:
                anchor(model,p,x=abs(q[0])<1e-9,y=abs(q[1])<1e-9)
            dimension_existing_relations(model)
            q = [g(p,k) for k in ("X","Y","Z")]
            a = g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData")
            xyz = [sum(q[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
            assert math.dist(xyz,(.02075,0,0))<1e-7, xyz
        finally:
            model.SketchManager.InsertSketch(True)
        assert model.ForceRebuild3(False)
        test = surfaces(model)
        assert cylinder_match(test,"z",(20.75,0,0),4.25,0,3), "沉孔母孔台阶不符"
        assert cylinder_match(test,"z",(20.75,0,0),2.5,3,20), "沉孔母孔实际通径不符"
        model.ClearSelection2(True)
        assert _select_first_candidate(model.Extension, _get_plane_name_candidates("Top Plane"), "PLANE")
        assert _select_first_candidate(model.Extension, _get_plane_name_candidates("Right Plane"), "PLANE", append=True)
        assert model.InsertAxis2(True)
        node = g(model,"FirstFeature")
        while node is not None and g(node,"GetTypeName2")!="RefAxis":
            node = g(node,"GetNextFeature")
        assert node is not None
        node.Name = "安装孔分布轴"
        circular_pattern(model,f.Name,node.Name,math.tau,4).Name = "四个底面安装沉孔"
    else:
        assert model.FeatureByName("底面安装沉孔母孔") is None
        cut_circles(model, 0, candidates, .005, 0, 1, "安装Φ5通孔")
        cut_circles(model, 0, candidates, .0085, .003, 1, "底面Φ8点5沉孔", pick=(0,0))

    # A-A明确标出的底面入口小倒角。四角孔在底面端视也画出入口外圈。
    model.ClearSelection2(True)
    chamfer_edges(model, [([x,y,0],r) for group,r in ((CORNER_XY,3.5),(MOUNT_XY,4.25))
                          for x,y in group], .5, "底面孔口C0点5")
    checkpoint("Φ41.5分布圆四安装孔及底面沉孔")


def side_ports(model, checkpoint):
    assert model.FeatureByName("底面孔口C0点5") is not None
    pilot, major = 5.9, 13.157/2
    run = 22
    tip = pilot/math.tan(math.radians(59))

    for label, offset in (("甲排",10.5), ("乙排",-10.5)):
        name = offset_plane(model, "Top Plane", offset, label+"侧口轴面")
        with sketch(model,name) as ref:
            sm = model.SketchManager
            active = g(sm,"ActiveSketch")
            transform = g(active,"ModelToSketchTransform")
            a = g(transform,"ArrayData")
            y = g(g(transform,"Inverse"),"ArrayData")[10]
            assert abs(abs(y)-.0105)<1e-7
            def local(x,z):
                xyz = (x/1000,y,z/1000)
                result = [sum(xyz[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
                assert abs(result[2])<1e-7
                return result[:2]

            contour = [(27.5,10), (27.5,10+major+1), (26.5,10+major),
                       (17.5,10+major), (17.5,10+pilot), (27.5-run,10+pilot),
                       (27.5-run-tip,10)]
            polys = [[local(sign*x,z) for x,z in contour] for sign in (-1,1)]
            old = sm.AddToDB
            sm.AddToDB = True
            try:
                loops = [[sketch_line(model,*p,*q) for p,q in zip(points,points[1:]+points[:1])]
                         for points in polys]
                axis = sm.CreateCenterLine(*polys[0][0],0,*polys[1][0],0)
            finally:
                sm.AddToDB = old
            for edges, points in zip(loops, polys):
                for cur, nxt in zip(edges,edges[1:]+edges[:1]):
                    relation(model,[g(cur,"GetEndPoint2"),g(nxt,"GetStartPoint2")],"sgCOINCIDENT")
                for edge,p,q in zip(edges,points,points[1:]+points[:1]):
                    if abs(p[0]-q[0])<1e-9:
                        relation(model,[edge],"sgVERTICAL")
                    elif abs(p[1]-q[1])<1e-9:
                        relation(model,[edge],"sgHORIZONTAL")
            for key, edges in zip(("GetStartPoint2","GetEndPoint2"),loops):
                relation(model,[g(axis,key),g(edges[0],"GetStartPoint2")],"sgCOINCIDENT")
            if abs(polys[0][0][0]-polys[1][0][0])<1e-9:
                relation(model,[axis],"sgVERTICAL")
            else:
                relation(model,[axis],"sgHORIZONTAL")
            dimension_existing_relations(model)
            for edges, points in zip(loops, polys):
                for e, p in zip(edges,points):
                    q = g(e,"GetStartPoint2")
                    assert math.dist((g(q,"X"),g(q,"Y")),p)<1e-7
        sn = name_sketch(model,ref,label+"两侧口剖面")
        _ensure_sketch_selected(model,sn)
        sd = g(model.SelectionManager,"CreateSelectData")
        sd.Mark = 16
        assert axis.Select4(True,sd)
        f = model.FeatureManager.FeatureRevolve2(True,True,False,True,False,False,0,0,
            math.tau,0,False,False,0,0,0,0,0,True,False,True)
        assert f is not None, label
        f.Name = label+"两Rp1/4侧口"
        checkpoint(f.Name)

    cut_circles(model, .024, [(x/1000,y/1000) for x,y in AIR_XY], .004, .014, -1, "止口面四Φ4连通孔")
    props = g(model.Extension,"CustomPropertyManager","")
    props.Add3("进出气口",30,"4×Rp1/4，有效入口段10；底孔全径深22；以名义圆柱简化，不建螺旋牙",2)
    checkpoint("四侧口与四Φ4连通")


def check_geometry(model, rows):
    result = {"外廓55×55×24": flat_box(model)==[-27.5,-27.5,0.0,27.5,27.5,24.0]}
    result["四角R1"] = all(cylinder_match(rows,"z",(x,y,0),1,0,20)
                          for x,y in ((i*26.5,j*26.5) for i in (-1,1) for j in (-1,1)))
    result["密封止口的1/2/1轴向及根径"] = all(cylinder_match(rows,"z",(0,0,0),r,lo,hi)
                                      for r,lo,hi in ((17.5,20,21),(16.25,21,23),(17.5,23,24)))
    result["中心Φ10盲窝深2"] = bool(cylinder_match(rows,"z",(0,0,0),5,22,24))
    result["四角孔Φ7贯穿，孔距44×44"] = all(cylinder_match(rows,"z",(x,y,0),3.5,.5,20) for x,y in CORNER_XY)
    result["Φ41.5分布圆上四Φ5及Φ8.5深3底面沉孔"] = all(
        cylinder_match(rows,"z",(x,y,0),2.5,3,20) and cylinder_match(rows,"z",(x,y,0),4.25,.5,3)
        for x,y in MOUNT_XY)

    positions = [(s,y) for s in (-1,1) for y in (-10.5,10.5)]
    result["四侧口位置及入口10和全孔22的阶梯"] = all(
        cylinder_match(rows,"x",(0,y,10),13.157/2,*sorted((s*17.5,s*26.5)))
        and cylinder_match(rows,"x",(0,y,10),5.9,*sorted((s*5.5,s*17.5)))
        for s,y in positions)
    tip = 5.5 - 5.9/math.tan(math.radians(59))
    result["四侧孔钻尖终点，中心保持隔墙"] = all(
        any("Cone" in r and abs(abs(r["Cone"][3])-1)<1e-8
            and abs(r["Cone"][1]*1000-y)<.001 and abs(r["Cone"][2]*1000-10)<.001
            and all(abs(r["box"][i]-v)<.001 for i,v in zip((0,3),sorted((s*tip,s*5.5))))
            for r in rows) for s,y in positions)
    result["Φ22分布圆的四Φ4向下连接对应侧口"] = all(
        any("Cylinder" in r and abs(r["Cylinder"][6]*1000-2)<.001
            and abs(r["Cylinder"][0]*1000-x)<.001 and abs(r["Cylinder"][1]*1000-y)<.001
            and abs(r["box"][5]-24)<.001 and 13<r["box"][2]<16
            for r in rows)
        for x,y in AIR_XY)
    return result


def check_sketches(quality):
    records = quality["sketches"]
    required = {"阀体外廓", "止口轮廓1", "止口轮廓2", "止口轮廓3",
                "四角Φ7贯穿轮廓", "Φ10中心盲窝轮廓", "甲排两侧口剖面",
                "乙排两侧口剖面", "止口面四Φ4连通孔轮廓"}
    names = {s["name"].split("<")[0] for s in records}
    return "设计草图全约束无Fix，向导内部系统轮廓单列", (
        required <= names and all(
            s["constrained_status"]==3 and not s["dangling"] and not s["errors"]
            and (not s["fix_count"] or (s["name"] not in required
                                      and s["name"]!="底面安装沉孔定位"
                                      and "底面安装沉孔母孔" in s["consumed_by"]))
            for s in records))


def main():
    out, pid, stage = Path(sys.argv[1]).resolve(), int(sys.argv[2]), sys.argv[3]
    target = out / "阀体_SDF-01.SLDPRT"
    assert out.is_dir()
    state = solidworks_processes()
    app = _attach_existing_with_pid_check(state, state[pid])
    model = g(app, "ActiveDoc")
    assert g(app, "GetProcessID") == pid and model is not None and g(model, "GetType") == 1
    if stage in ("review", "finish"):
        assert Path(g(model, "GetPathName")).resolve() == target
        if stage == "review":
            prepare_review(app,target,check_geometry,sketch_count=11,sketch_quality_check=check_sketches,
                           views=["isometric","front","back","right_wire"])
        else:
            finish_review(app,target,visual_notes=sys.argv[4])
        return

    def checkpoint(label):
        assert model.ForceRebuild3(False) and len(g(model, "GetBodies2", 0, False)) == 1
        assert save_document(model, str(target))
        model.ClearSelection2(True)
        model.ShowNamedView2("*Isometric", 7)
        model.ViewZoomtofit2()
        model.GraphicsRedraw2()
        with (out / "stages.jsonl").open("a", encoding="utf-8") as log:
            log.write(json.dumps({"time":datetime.now().astimezone().isoformat(),
                                  "stage":label, "file":target.name}, ensure_ascii=False)+"\n")
        print(label, flush=True)

    with quiet_dimensions(app):
        if stage == "body":
            assert not target.exists()
            body(model, checkpoint)
        else:
            assert Path(g(model, "GetPathName")).resolve() == target
            if stage in ("holes","holes-cut"):
                mount_holes(model,checkpoint,wizard=stage=="holes")
            elif stage == "ports":
                side_ports(model,checkpoint)
            else:
                raise ValueError(stage)


if __name__ == "__main__":
    main()
