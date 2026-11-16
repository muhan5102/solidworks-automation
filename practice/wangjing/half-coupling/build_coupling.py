"""LJT04.30：原图特殊阶梯锥孔采用单次回转切除加八实例阵列。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/"year08/connector"))
from build_connector import (member as g,mm,sketch,sketch_line,sketch_circle,
    new_document,save_document,dimension_existing_relations,offset_plane,chamfer_edges)
from sw_part import revolve_boss,extrude_cut,_ensure_sketch_selected,_get_plane_name_candidates,_select_first_candidate
from sw_connect import create_empty_dispatch_variant
from sw_sketch_quality import fully_define_sketch,inspect_model_sketches

out=Path(sys.argv[1]).resolve();target=out/"半联轴器_LJT04.30.SLDPRT"
resume=len(sys.argv)>2 and sys.argv[2]=="finish"
assert target.exists() if resume else not target.exists()
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
s.Visible=True;m=g(s,"ActiveDoc") if resume else new_document(s,"part");start=time.perf_counter()
if resume:assert Path(g(m,"GetPathName")).resolve()==target
def relation(items,kind):
    m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData")
    for i,item in enumerate(items):assert item.Select4(i>0,data)
    m.SketchAddConstraints(kind)
def origin(p):
    m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData");assert p.Select4(False,data)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
    m.SketchAddConstraints("sgCOINCIDENT")
def connected(edges):
    for a,b in zip(edges,edges[1:]+edges[:1]):
        relation([g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
def orthogonal(edges,points,skip=()):
    for i,(e,a,b) in enumerate(zip(edges,points,points[1:]+points[:1])):
        if i not in skip and (a[0]==b[0] or a[1]==b[1]):
            relation([e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
def checkpoint(stage):
    assert m.ForceRebuild3(False);assert save_document(m,str(target))
    m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
    (out/"checkpoint.json").write_text(json.dumps({"stage":stage,"saved":datetime.now().astimezone().isoformat(),
        "file":target.name},ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"[{stage}] 已保存，执行{time.perf_counter()-start:.2f}秒",flush=True)

if not resume:
    points=[(21,0),(50,0),(50,19.2),(47,19.2),(47,20.8),(50,20.8),(50,34),
            (54,38),(100,38),(100,80),(21,80)]
    with sketch(m,"Front Plane") as ref:
        sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
        try:
            axis=sm.CreateCenterLine(0,0,0,0,.080,0);edges=[]
            for i,(a,b) in enumerate(zip(points,points[1:]+points[:1])):
                e=sm.CreateArc(.054,.034,0,mm(a[0]),mm(a[1]),0,mm(b[0]),mm(b[1]),0,-1) if i==6 \
                    else sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1]))
                assert e is not None;edges.append(e)
        finally:sm.AddToDB=old
        connected(edges);orthogonal(edges,points,[6])
        origin(g(axis,"GetStartPoint2"));relation([axis],"sgVERTICAL")
        relation([g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([g(axis,"GetEndPoint2"),g(edges[9],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([edges[5],edges[6]],"sgTANGENT");relation([edges[6],edges[7]],"sgTANGENT")
        dimension_existing_relations(m)
        for e,p in zip(edges,points):
            q=g(e,"GetStartPoint2");assert math.dist([g(q,"X")*1000,g(q,"Y")*1000],p)<1e-5
    m.FeatureByName(str(ref)).Name="主体轴向截面"
    f=revolve_boss(m,"主体轴向截面",2*math.pi);assert f is not None;f.Name="主体回转"
    chamfer_edges(m,[([0,0,0],21),([0,80,0],21),([0,0,0],50),
                     ([0,38,0],100),([0,80,0],100)],2,"图示C2")
    checkpoint("主体槽和R4")

    # 八孔母孔轴线平行Y，世界中心X75,Z0，入口在法兰后侧Y38。
    # 三级截面：D38深16，D26深0.5，1:10锥孔右端D18。
    points=[(75,38),(94,38),(94,54),(88,54),(88,54.5),(82.725,54.5),(84,80),(75,80)]
    with sketch(m,"Front Plane") as ref:
        sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
        try:
            axis=sm.CreateCenterLine(.075,.038,0,.075,.080,0)
            edges=[sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1])) for a,b in zip(points,points[1:]+points[:1])]
        finally:sm.AddToDB=old
        connected(edges);orthogonal(edges,points)
        relation([g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgCOINCIDENT")
        relation([g(axis,"GetEndPoint2"),g(edges[-1],"GetStartPoint2")],"sgCOINCIDENT")
        relation([axis],"sgVERTICAL")
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name="阶梯锥孔截面"
    _ensure_sketch_selected(m,"阶梯锥孔截面")
    data=g(m.SelectionManager,"CreateSelectData");data.Mark=16;assert axis.Select4(True,data)
    f=m.FeatureManager.FeatureRevolve2(True,True,False,True,False,False,0,0,2*math.pi,0,
                                      False,False,0,0,0,0,0,True,False,True)
    assert f is not None;f.Name="阶梯锥孔"
    m.ClearSelection2(True)
    assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Front Plane"),"PLANE")
    assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Right Plane"),"PLANE",append=True)
    assert m.InsertAxis2(True)
    f=g(m,"FirstFeature")
    while f is not None:
        if g(f,"GetTypeName2")=="RefAxis":f.Name="回转轴"
        f=g(f,"GetNextFeature")
    m.ClearSelection2(True);assert m.FeatureByName("阶梯锥孔").Select2(False,4)
    assert m.Extension.SelectByID2("回转轴","AXIS",0,0,0,True,1,create_empty_dispatch_variant(),0)
    f=m.FeatureManager.FeatureCircularPattern4(8,2*math.pi,False,"",False,True,False)
    assert f is not None;f.Name="八孔阵列"
    checkpoint("八个阶梯锥孔")

# 键槽在+Z侧。12JS9宽，45.3是从D42孔底Z=-21到键槽顶面的距离。
if resume:
    f=m.FeatureByName("草图3")
    assert f is not None
    m.ClearSelection2(True);assert f.Select2(False,0);assert m.Extension.DeleteSelection2(0)
with sketch(m,"Top Plane") as ref:
    sk=g(m.SketchManager,"ActiveSketch");t=g(sk,"ModelToSketchTransform");a=list(g(t,"ArrayData"));inv=list(g(g(t,"Inverse"),"ArrayData"))
    def local(x,z):
        p=[mm(x),0,mm(z)]
        return [sum(p[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
    world=[(-6,0),(6,0),(6,24.3),(-6,24.3)];pts=[local(*p) for p in world]
    sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
    try:
        edges=[sketch_line(m,*aa[:2],*bb[:2]) for aa,bb in zip(pts,pts[1:]+pts[:1])]
        base=local(0,-21);top=local(0,24.3);datum=sm.CreateCenterLine(*base,*top)
    finally:sm.AddToDB=old
    connected(edges);orthogonal(edges,pts)
    relation([datum],"sgVERTICAL")
    relation([g(datum,"GetEndPoint2"),edges[2]],"sgCOINCIDENT")
    relation([g(edges[0],"GetStartPoint2"),g(edges[0],"GetEndPoint2"),datum],"sgSYMMETRIC")
    for point,kind in [(g(edges[0],"GetStartPoint2"),"sgHORIZONTALPOINTS2D"),
                       (g(datum,"GetStartPoint2"),"sgVERTICALPOINTS2D")]:
        m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData");assert point.Select4(False,data)
        assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
        m.SketchAddConstraints(kind)
    old_prompt=s.GetUserPreferenceToggle(10);s.SetUserPreferenceToggle(10,False)
    try:
        m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData");assert edges[0].Select4(False,data)
        dd=m.AddDimension2(0,0,-.010);assert dd is not None
        assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-.012)<1e-8
        m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData");assert datum.Select4(False,data)
        dd=m.AddDimension2(.012,0,0);assert dd is not None
        assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-.0453)<1e-8
    finally:s.SetUserPreferenceToggle(10,old_prompt)
    dimension_existing_relations(m)
f=extrude_cut(m,ref,.080,direction=inv[7]>0);assert f is not None;f.Name="贯穿键槽"
m.FeatureByName(str(ref)).Name="键槽轮廓"

# 径向M12位于Y20、-Z侧，钻透单壁，不接触+Z键槽。
plane=offset_plane(m,"Front Plane",-50,"径向孔入口")
with sketch(m,plane) as ref:
    sk=g(m.SketchManager,"ActiveSketch");t=g(sk,"ModelToSketchTransform");a=list(g(t,"ArrayData"));inv=list(g(g(t,"Inverse"),"ArrayData"))
    p=[0,.020,-.050];q=[sum(p[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
    assert abs(q[2])<1e-8;assert sketch_circle(m,q[0],q[1],.0051) is not None
    fully_define_sketch(m)
f=extrude_cut(m,ref,.030,direction=inv[8]>0);assert f is not None;f.Name="M12径向底孔"
m.FeatureByName(str(ref)).Name="径向孔位"
holes=[]
for b in g(m,"GetBodies2",0,False):
    for face in g(b,"GetFaces"):
        su=g(face,"GetSurface")
        if g(su,"IsCylinder"):
            p=g(su,"CylinderParams")
            if abs(p[6]-.0051)<1e-8 and abs(p[1]-.020)<1e-8:holes.append(face)
assert len(holes)==1
edges=list(g(holes[0],"GetEdges"))
edge=min(edges,key=lambda e:min(list(g(g(e,"GetCurveParams3"),"StartPoint"))[2],
                               list(g(g(e,"GetCurveParams3"),"EndPoint"))[2]))
m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData");assert edge.Select4(False,data)
thread=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M12x1.75",.012,0,.029,"M12-7H，径向通单壁")
assert thread is not None;thread.Name="M12装饰螺纹"
for name in ["回转轴","径向孔入口"]:
    m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,0);m.BlankRefGeom()
checkpoint("键槽和径向螺纹")
(out/"build.json").write_text(json.dumps({"finished":datetime.now().astimezone().isoformat(),
    "execution_seconds":time.perf_counter()-start,"quality":inspect_model_sketches(m)},ensure_ascii=False,indent=2),encoding="utf-8")
