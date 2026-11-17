"""LJT05.19：弯曲变宽T形臂、斜安装面及GB孔向导。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "practice"))
from shared.cad import member as g, new_document, save_document, sketch, sketch_line, dimension_existing_relations
from shared.single_part import relation, anchor, drive, circle, quiet_dimensions
from sw_part import extrude_midplane, extrude_boss, extrude_cut, sketch_arc, _ensure_sketch_selected
from sw_review import save_preview

out = Path(sys.argv[1]).resolve()
stage = sys.argv[2]
target = out / "踏架_LJT05.19.SLDPRT"
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert g(sw, "GetProcessID") == 49384
m = g(sw, "ActiveDoc")
timer = time.perf_counter()


def poly(points):
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        edges = [sketch_line(m, a[0]/1000, a[1]/1000, b[0]/1000, b[1]/1000)
                 for a, b in zip(points, points[1:]+points[:1])]
    finally:
        sm.AddToDB = old
    join(edges)
    for e, a, b in zip(edges, points, points[1:]+points[:1]):
        if a[0] == b[0] or a[1] == b[1]:
            relation(m, [e], "sgVERTICAL" if a[0] == b[0] else "sgHORIZONTAL")
    return edges


def join(edges):
    for a, b in zip(edges, edges[1:]+edges[:1]):
        relation(m, [g(a, "GetEndPoint2"), g(b, "GetStartPoint2")], "sgCOINCIDENT")


def name_sketch(ref, name):
    m.FeatureByName(str(ref)).Name = name
    return name


def checkpoint(name):
    assert m.ForceRebuild3(False) and save_document(m, str(target))
    m.ClearSelection2(True)
    m.ShowNamedView2("*Isometric", 7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    (out/"checkpoint.json").write_text(json.dumps({"stage": name,
        "saved": datetime.now().astimezone().isoformat(), "target": target.name,
        "next": "按计划继续加料或孔槽；未做最终验收"}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[{name}] 已保存，本段{time.perf_counter()-timer:.2f}秒", flush=True)


def select_face(normal, point):
    faces=[]
    for body in g(m,"GetBodies2",0,False):
        for face in g(body,"GetFaces"):
            su=g(face,"GetSurface")
            if not g(su,"IsPlane"):continue
            p=list(g(su,"PlaneParams"));box=list(g(face,"GetBox"))
            if abs(abs(sum(p[i]*normal[i] for i in range(3)))-1)<1e-8 and abs(sum((point[i]-p[i+3])*p[i] for i in range(3)))<1e-8:
                if all(box[i]-1e-8<=point[i]<=box[i+3]+1e-8 for i in range(3)):faces.append(face)
    assert len(faces)==1,len(faces)
    m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData")
    sd.X,sd.Y,sd.Z=point
    assert faces[0].Select4(False,sd)


def locate_wizard(f,name,world):
    assert f is not None,name
    f.Name=name
    sf=g(f,"GetFirstSubFeature");sf.Name=name+"孔位"
    profile=g(sf,"GetNextSubFeature")
    if profile is not None:profile.Name=name+"原生轮廓"
    m.ClearSelection2(True);assert sf.Select2(False,0);m.EditSketch()
    sk=g(m.SketchManager,"ActiveSketch");pts=list(g(sk,"GetSketchPoints2"));assert len(pts)==1
    p=pts[0]
    if abs(g(p,"X"))<1e-9 or abs(g(p,"Y"))<1e-9:
        anchor(m,p,x=abs(g(p,"X"))<1e-9,y=abs(g(p,"Y"))<1e-9)
    dimension_existing_relations(m)
    a=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"));q=[g(p,"X"),g(p,"Y"),g(p,"Z")]
    actual=[sum(q[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
    assert math.dist(actual,world)<1e-7,(actual,world)
    m.SketchManager.InsertSketch(True)
    return g(f,"GetDefinition")


with quiet_dimensions(sw):
    if stage == "body":
        assert not target.exists() and g(sw, "GetDocumentCount") == 0
        m = new_document(sw, "part")
        # R9、厚5、圆心(28,24)，外侧与踏板左顶点(59,37)切线连接。
        slope = (806-math.sqrt(806**2+4*765*27))/(2*765)
        norm = math.sqrt(1+slope*slope)
        po = (28-14*slope/norm, 24+14/norm)
        pi = (28-9*slope/norm, 24+9/norm)
        yi = 37-5*norm
        with sketch(m, "Front Plane") as ref:
            sm = m.SketchManager
            old = sm.AddToDB
            sm.AddToDB = True
            try:
                es = [sketch_line(m, .014, 0, .019, 0), sketch_line(m, .019, 0, .019, .024),
                      sketch_arc(m, .028, .024, .019, .024, pi[0]/1000, pi[1]/1000, -1),
                      sketch_line(m, pi[0]/1000, pi[1]/1000, .059, yi/1000),
                      sketch_line(m, .059, yi/1000, .059, .037),
                      sketch_line(m, .059, .037, po[0]/1000, po[1]/1000),
                      sketch_arc(m, .028, .024, po[0]/1000, po[1]/1000, .014, .024, 1),
                      sketch_line(m, .014, .024, .014, 0)]
            finally:
                sm.AddToDB = old
            join(es)
            relation(m, [es[0]], "sgHORIZONTAL")
            for i in [1, 4, 7]: relation(m, [es[i]], "sgVERTICAL")
            for a, b in [(1, 2), (2, 3), (5, 6), (6, 7)]: relation(m, [es[a], es[b]], "sgTANGENT")
            relation(m, [es[2], es[6]], "sgCONCENTRIC")
            relation(m, [es[3], es[5]], "sgPARALLEL")
            anchor(m, g(es[0], "GetStartPoint2"), y=True)
            drive(m, [es[2]], (.025, .030, 0), .009)
            drive(m, [es[0]], (.015, -.005, 0), .005)
            dimension_existing_relations(m)
            for e, r in [(es[2], .009), (es[6], .014)]: assert abs(g(e, "GetRadius")-r)<1e-8
        f = extrude_midplane(m, name_sketch(ref, "弯臂翼缘截面"), .038)
        assert f is not None
        f.Name = "翼缘等宽毛坯"
        # 先在单独翼缘上收宽，再加轴套与踏板，避免切伤相邻主体。
        with sketch(m, "Top Plane") as ref:
            for sign in [1, -1]:
                poly([(14, sign*19), (59, sign*14), (65, sign*14), (65, sign*22), (14, sign*22)])
            dimension_existing_relations(m)
        f = extrude_cut(m, name_sketch(ref, "翼缘38至28收宽"), 0, both_directions=True)
        assert f is not None
        f.Name = "翼缘渐变宽度"
        with sketch(m, "Front Plane") as ref:
            c = circle(m, 0, 0, .019)
            anchor(m, g(c, "GetCenterPoint2"), x=True, y=True)
            drive(m, [c], (-.025, -.015, 0), .038, kind="diameter")
            sm = m.SketchManager
            old = sm.AddToDB
            sm.AddToDB = True
            try:
                vertical = sm.CreateCenterLine(0,0,0,0,.028,0)
                normal = sm.CreateCenterLine(0,0,0,-.014,.028*math.cos(math.pi/6),0)
            finally: sm.AddToDB = old
            for e in [vertical, normal]: anchor(m, g(e,"GetStartPoint2"), x=True,y=True)
            relation(m,[vertical],"sgVERTICAL")
            relation(m,[vertical,normal],"sgSAMELENGTH")
            drive(m,[normal],(-.020,.012,0),.028)
            drive(m,[vertical,normal],(-.010,.035,0),math.pi/6)
            dimension_existing_relations(m)
        f = extrude_midplane(m, name_sketch(ref,"轴套与斜面定位"), .058)
        assert f is not None
        f.Name = "轴套实心"
        with sketch(m,"Front Plane") as ref:
            poly([(59,30),(93,30),(93,37),(59,37)])
            dimension_existing_relations(m)
        f = extrude_midplane(m,name_sketch(ref,"踏板截面"),.050)
        assert f is not None
        f.Name="踏板"
        checkpoint("翼缘、轴套和踏板")
        # 垂直于30°定位线、经过28端点，唯一有结构用途的斜基准面。
        m.ClearSelection2(True)
        sd=g(m.SelectionManager,"CreateSelectData");sd.Mark=0
        assert normal.Select4(False,sd)
        sd=g(m.SelectionManager,"CreateSelectData");sd.Mark=1
        assert g(normal,"GetEndPoint2").Select4(True,sd)
        plane=m.FeatureManager.InsertRefPlane(2,0.,4,0.,0,0.)
        assert plane is not None,"斜面创建失败，已保留阶段文件"
        ft=g(m,"FirstFeature");last=None
        while ft is not None:
            if g(ft,"GetTypeName2")=="RefPlane":last=ft
            ft=g(ft,"GetNextFeature")
        last.Name="30度安装面"
        checkpoint("安装面定位完成")
    elif stage == "rib":
        assert m is not None and Path(g(m,"GetPathName")).resolve()==target
        # SolidWorks退出空草图后会丢弃空节点，重新在安装面上开始。
        from sw_connect import create_empty_dispatch_variant
        if g(m.SketchManager,"ActiveSketch") is None:
            m.ClearSelection2(True)
            assert m.Extension.SelectByID2("30度安装面","PLANE",0,0,0,False,0,create_empty_dispatch_variant(),0)
            m.SketchManager.InsertSketch(True)
        assert not (g(g(m.SketchManager,"ActiveSketch"),"GetSketchSegments") or [])
        ref=g(g(m.SketchManager,"ActiveSketch"),"Name")
        sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
        try:
            es=[sketch_line(m,.010,-.0125,.010,.0125),
                sketch_arc(m,0,.0125,.010,.0125,-.010,.0125,1),
                sketch_line(m,-.010,.0125,-.010,-.0125),
                sketch_arc(m,0,-.0125,-.010,-.0125,.010,-.0125,1)]
        finally:sm.AddToDB=old
        join(es)
        for i in [0,2]:relation(m,[es[i]],"sgVERTICAL")
        for a,b in [(0,1),(1,2),(2,3),(3,0)]:relation(m,[es[a],es[b]],"sgTANGENT")
        relation(m,[es[1],es[3]],"sgSAMELENGTH")
        anchor(m,g(es[1],"GetCenterPoint2"),x=True)
        drive(m,[es[1]],(.015,.020,0),.010)
        dimension_existing_relations(m)
        sm.InsertSketch(True)
        f=extrude_boss(m,name_sketch(ref,"安装凸台腰形"),.015,direction=True)
        assert f is not None
        f.Name="斜安装凸台"
        assert len(g(m,"GetBodies2",0,False))==1
        # R30与右踏板孔轴(76,37)的上切线，左侧竖直切线延入主体。
        # (48k-13)^2=30^2(1+k^2)，取负斜率支。
        k=(1248-math.sqrt(1248**2+4*1404*731))/(2*1404)
        norm=math.sqrt(1+k*k)
        p=(28-30*k/norm,24+30/norm)
        with sketch(m,"Front Plane") as ref:
            sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
            try:
                es=[sketch_line(m,-.002,.015,-.002,.024),
                    sketch_arc(m,.028,.024,-.002,.024,p[0]/1000,p[1]/1000,-1),
                    sketch_line(m,p[0]/1000,p[1]/1000,.076,.037)]
            finally:sm.AddToDB=old
            for a,b in zip(es,es[1:]):relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
            relation(m,[es[0]],"sgVERTICAL")
            relation(m,[es[0],es[1]],"sgTANGENT")
            relation(m,[es[1],es[2]],"sgTANGENT")
            drive(m,[es[1]],(.020,.057,0),.030)
            dimension_existing_relations(m)
        ref=name_sketch(ref,"腹板R30边界")
        _ensure_sketch_selected(m,ref)
        m.FeatureManager.InsertRib(True,False,.010,0,False,False,False,0.,False,False)
        f=g(m,"FirstFeature");ribs=[]
        while f is not None:
            if g(f,"GetTypeName2")=="Rib":ribs.append(f)
            f=g(f,"GetNextFeature")
        assert len(ribs)==1,"筋创建失败，需检查材料侧和边界"
        ribs[0].Name="T形腹板"
        checkpoint("全部主体加料")
        m.ClearSelection2(True)
        assert m.FeatureByName("30度安装面").Select2(False,0)
        m.BlankRefGeom()
    elif stage == "holes":
        assert Path(g(m,"GetPathName")).resolve()==target
        if m.FeatureByName("圆角1") is not None:m.FeatureByName("圆角1").Name="踏板四角R5_MCP"
        # MCP当前入口不支持双向贯穿切除；用原生COM补齐真实ThroughAll。
        if m.FeatureByName("主孔贯通") is None:
            with sketch(m,"Front Plane") as ref:
                c=circle(m,0,0,.012)
                anchor(m,g(c,"GetCenterPoint2"),x=True,y=True)
                drive(m,[c],(-.018,-.020,0),.024,kind="diameter")
            f=extrude_cut(m,name_sketch(ref,"主孔草图"),0,both_directions=True)
            assert f is not None
            f.Name="主孔贯通"
        f=m.FeatureByName("GB沉头母孔")
        if f is None:
            select_face((0,1,0),(.076,.037,-.016))
            f=m.FeatureManager.HoleWizard5(1,13,362,"M6",1,.007,.007,-1,
                .013,math.pi/2,0,-1,-1,0,0,-1,0,-1,-1,-1,
                "",False,False,True,False,False,False)
            d=locate_wizard(f,"GB沉头母孔",(.076,.037,-.016))
        else:d=g(f,"GetDefinition")
        assert g(d,"Standard2")==13
        assert abs(g(d,"ThruHoleDiameter")-.007)<1e-8
        assert abs(g(d,"CounterSinkDiameter")-.013)<1e-8
        assert abs(g(d,"CounterSinkAngle")-math.pi/2)<1e-8
        checkpoint("主孔与GB沉头母孔")
    elif stage == "tap":
        assert Path(g(m,"GetPathName")).resolve()==target
        n=(-.5,math.cos(math.pi/6),0)
        point=(n[0]*.028,n[1]*.028,-.0125)
        select_face(n,point)
        # 本机UpToNext实测会误穿远壁；20深在主孔内结束，只打开近壁。
        f=m.FeatureManager.HoleWizard5(4,13,359,"M10",0,.0085,.020,-1,
            .020,.010,math.pi/2,0,0,-1,1,0,-1,-1,-1,-1,
            "7H",False,False,True,False,False,False)
        d=locate_wizard(f,"GB_M10螺纹母孔",point)
        assert g(d,"Standard2")==13
        print('tap',g(d,"Type"),g(d,"HoleDiameter"),flush=True)
        checkpoint("GB螺纹母孔")
    else:
        assert m is not None and Path(g(m,"GetPathName")).resolve()==target
        raise ValueError(stage)
    save_preview(m,str(out/(stage+".bmp")),"isometric")
