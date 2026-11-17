"""王静LJT05.11；阶段执行，目标及进程显式核对。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "practice"))
from shared.cad import member as g, mm, sketch, sketch_line, new_document, save_document, dimension_existing_relations, offset_plane
from sw_part import extrude_midplane, extrude_cut, _ensure_sketch_selected, sketch_arc
from sw_review import save_preview
from sw_connect import connect_solidworks
sys.path.insert(0, str(ROOT / "practice"))
from shared.single_part import relation, circle, drive, anchor

out = Path(sys.argv[1]).resolve()
stage = sys.argv[2]
target = out / "支架_LJT05.11.SLDPRT"
assert out.is_dir()
sw, _ = connect_solidworks(version=2024, visible=True, wait_seconds=.1)
assert g(sw, "GetProcessID") == 42548
if stage == "body":
    assert not target.exists() and g(sw, "GetFirstDocument") is None
    m = new_document(sw, "part")
else:
    m = g(sw, "ActiveDoc")
    assert (stage == "tube-resume" and not g(m,"GetPathName") and m.FeatureByName("底座") is not None) or Path(g(m, "GetPathName")).resolve() == target
timer = time.perf_counter()


def poly(points):
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        edges = [sketch_line(m, mm(a[0]), mm(a[1]), mm(b[0]), mm(b[1]))
                 for a, b in zip(points, points[1:] + points[:1])]
    finally:
        sm.AddToDB = old
    for a, b in zip(edges, edges[1:] + edges[:1]):
        relation(m, [g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")
    for edge, a, b in zip(edges, points, points[1:] + points[:1]):
        if a[0] == b[0] or a[1] == b[1]:
            relation(m, [edge], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    return edges


def boss(points, thickness, name):
    with sketch(m, "Front Plane") as ref:
        poly(points)
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name = name + "截面"
    feature = extrude_midplane(m, name + "截面", mm(thickness))
    assert feature is not None, name
    feature.Name = name
    return feature


def checkpoint(name):
    assert m.ForceRebuild3(False)
    assert save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    (out / "checkpoint.json").write_text(json.dumps({"stage": name, "saved": datetime.now().astimezone().isoformat(),
        "target": target.name, "configuration": g(g(m.ConfigurationManager, "ActiveConfiguration"), "Name"),
        "next": "孔槽与过渡" if stage == "body" else "集中终检"}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[{name}] 保存完成，本段{time.perf_counter()-timer:.2f}秒", flush=True)


old_prompt = sw.GetUserPreferenceToggle(10)
sw.SetUserPreferenceToggle(10, False)
try:
    if stage in ["body","tube-resume"]:
        # 轴线Z，夹套中心为原点；底座背面X60、顶承面Y-80。
        if stage == "body":
            boss([(44,-120),(60,-120),(60,-80),(68,-80),(68,-70),(44,-70)], 82, "底座")
            with sketch(m, "Front Plane") as ref:
                for r in [17.5,10]:
                    c = circle(m, 0, 0, mm(r))
                    anchor(m,g(c,"GetCenterPoint2"),x=True,y=True)
                    drive(m, [c], (mm(r+5),mm(r+5),0), mm(2*r), "diameter")
                dimension_existing_relations(m)
        else:
            ref="草图2"
            m.ClearSelection2(True)
            assert m.FeatureByName(ref).Select2(False,0)
            m.EditSketch()
            for c in g(g(m.SketchManager,"ActiveSketch"),"GetSketchSegments"):
                anchor(m,g(c,"GetCenterPoint2"),x=True,y=True)
            dimension_existing_relations(m)
            m.SketchManager.InsertSketch(True)
        m.FeatureByName(str(ref)).Name = "套筒截面"
        f = extrude_midplane(m, "套筒截面", .050)
        assert f is not None
        f.Name = "夹紧套筒"
        # 翼板外侧为从(64,-70)向R17.5圆作的上侧切线；厚6、宽40。
        px,py,r = 64.,-70.,17.5
        l2 = px*px+py*py
        k = math.sqrt(l2-r*r)/l2
        nx,ny = r*px/l2-k*py, r*py/l2+k*px
        tx,ty = r*nx,r*ny
        points = [(tx,ty),(px,py),(px-6*nx,py-6*ny),(tx-6*nx,ty-6*ny)]
        boss(points,40,"斜臂翼板")
        checkpoint("主体阶段")
        # 腹板的下边由图示X=-7以及底座左侧距底5定义。
        with sketch(m, "Front Plane") as ref:
            old=m.SketchManager.AddToDB
            m.SketchManager.AddToDB=True
            try:
                line = sketch_line(m, -.007, -mm(math.sqrt(17.5**2-7**2)), .044, -.115)
            finally:m.SketchManager.AddToDB=old
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name = "腹板边界"
        before = []
        f = g(m, "FirstFeature")
        while f is not None:
            if g(f,"GetTypeName2") == "Rib": before.append(g(f,"Name"))
            f = g(f,"GetNextFeature")
        _ensure_sketch_selected(m,"腹板边界")
        # 本机类型库为十参数且返回VOID；不调用安装版九参数封装。
        m.FeatureManager.InsertRib(True,False,.008,0,True,False,False,0.,False,False)
        added=[]
        f=g(m,"FirstFeature")
        while f is not None:
            if g(f,"GetTypeName2")=="Rib" and g(f,"Name") not in before: added.append(f)
            f=g(f,"GetNextFeature")
        assert len(added)==1, "未新增唯一原生Rib；主体保存点已保留"
        added[0].Name="斜臂腹板"
        checkpoint("原生筋阶段")
        save_preview(m,str(out/"body.bmp"),"isometric")
    elif stage in ["holes","ear-resume","ear-rebuild","ear-done","remaining"]:
        rib=m.FeatureByName("筋2")
        if rib is not None: rib.Name="斜臂腹板"
        checkpoint("原生筋已确认")
        if stage != "remaining":
            if stage == "ear-rebuild":
                m.ClearSelection2(True)
                assert m.FeatureByName("草图5").Select2(False,0)
                assert m.Extension.DeleteSelection2(0)
            if stage in ["holes","ear-rebuild"]:
              with sketch(m,"Top Plane") as ref:
                sm=m.SketchManager
                old=sm.AddToDB
                sm.AddToDB=True
                try:
                    arc=sketch_arc(m,-.025,0,-.025,.013,-.025,-.013,1)
                    edges=[arc,sketch_line(m,-.025,-.013,0,-.013),sketch_line(m,0,-.013,0,.013),sketch_line(m,0,.013,-.025,.013)]
                finally: sm.AddToDB=old
                for a,b in zip(edges,edges[1:]+edges[:1]):
                    relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
                for e,k in zip(edges[1:],["sgHORIZONTAL","sgVERTICAL","sgHORIZONTAL"]): relation(m,[e],k)
                anchor(m,g(arc,"GetCenterPoint2"),y=True)
                anchor(m,g(edges[2],"GetStartPoint2"),x=True)
                relation(m,[g(arc,"GetStartPoint2"),g(arc,"GetEndPoint2")],"sgVERTICALPOINTS2D")
                relation(m,[g(arc,"GetCenterPoint2"),g(arc,"GetStartPoint2")],"sgVERTICALPOINTS2D")
                dimension_existing_relations(m)
            elif stage == "ear-done":
                ref="草图6"
            else:
                ref="草图5"
                m.ClearSelection2(True)
                assert m.FeatureByName(ref).Select2(False,0)
                m.EditSketch()
                arc=list(g(g(m.SketchManager,"ActiveSketch"),"GetSketchSegments"))[0]
                anchor(m,g(arc,"GetCenterPoint2"),y=True)
                dimension_existing_relations(m)
                m.SketchManager.InsertSketch(True)
            m.FeatureByName(str(ref)).Name="夹耳轮廓"
            f=extrude_midplane(m,"夹耳轮廓",.018)
            assert f is not None
            f.Name="夹紧耳"

        def cut_circles(plane,centers,radius,depth,direction_world,name):
            with sketch(m,plane) as ref:
                transform=g(g(m.SketchManager,"ActiveSketch"),"ModelToSketchTransform")
                a=list(g(transform,"ArrayData"))
                inv=list(g(g(transform,"Inverse"),"ArrayData"))
                normal=inv[6:9]
                for p in centers:
                    q=[sum(mm(p[j])*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
                    assert abs(q[2])<1e-8,(plane,p,q)
                    c=circle(m,q[0],q[1],mm(radius))
                    if abs(q[0])<1e-9 or abs(q[1])<1e-9:
                        anchor(m,g(c,"GetCenterPoint2"),x=abs(q[0])<1e-9,y=abs(q[1])<1e-9)
                    drive(m,[c],(q[0]+.012,q[1]+.012,0),mm(2*radius),"diameter")
                dimension_existing_relations(m)
            m.FeatureByName(str(ref)).Name=name+"草图"
            f=extrude_cut(m,name+"草图",mm(depth),direction=sum(normal[i]*direction_world[i] for i in range(3))>0)
            assert f is not None,name
            f.Name=name
            return f

        if stage != "remaining":
            # 夹耳加料后重切主孔，避免耳部侵入Φ20通孔。
            front=offset_plane(m,"Front Plane",25,"套筒端面")
            cut_circles(front,[(0,0,25)],10,50,(0,0,-1),"主孔贯通")
            with sketch(m,"Front Plane") as ref:
                edges=poly([(-40,-1.5),(0,-1.5),(0,1.5),(-40,1.5)])
                anchor(m,g(edges[1],"GetStartPoint2"),x=True)
                dimension_existing_relations(m)
        else:
            ref="草图8"
            m.ClearSelection2(True)
            assert m.FeatureByName(ref).Select2(False,0)
            m.EditSketch()
            segs=list(g(g(m.SketchManager,"ActiveSketch"),"GetSketchSegments"))
            anchor(m,g(segs[1],"GetStartPoint2"),x=True)
            dimension_existing_relations(m)
            m.SketchManager.InsertSketch(True)
        m.FeatureByName(str(ref)).Name="夹紧槽草图"
        f=extrude_cut(m,"夹紧槽草图",.026,both_directions=True)
        assert f is not None
        f.Name="夹紧槽"
        back=offset_plane(m,"Right Plane",60,"底座背面")
        cut_circles(back,[(60,-100,-20),(60,-100,20)],7.5,16,(-1,0,0),"两安装通孔")
        cut_circles(back,[(60,-100,-20),(60,-100,20)],14,3,(-1,0,0),"两安装沉孔")
        upper=offset_plane(m,"Top Plane",9,"夹耳上表面")
        cut_circles(upper,[(-25,9,0)],5.5,7.5,(0,-1,0),"夹耳间隙孔")
        cut_circles(upper,[(-25,9,0)],9,3,(0,-1,0),"夹耳沉孔")
        lower=offset_plane(m,"Top Plane",-9,"夹耳下表面")
        cut_circles(lower,[(-25,-9,0)],4.25,7.5,(0,1,0),"M10攻丝底孔")
        checkpoint("孔槽阶段")
        save_preview(m,str(out/"holes.bmp"),"isometric")
    else:
        raise ValueError(stage)
finally:
    sw.SetUserPreferenceToggle(10,old_prompt)
    m.GraphicsRedraw2()
