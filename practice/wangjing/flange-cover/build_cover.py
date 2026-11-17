"""LJT04.18：按本机session.json串行执行，回转由正式MCP完成。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from shared.cad import (member as g,mm,sketch,sketch_line,sketch_circle,
    new_document,save_document,dimension_existing_relations,offset_plane,
    circle_cut,chamfer_edges)
from sw_connect import create_empty_dispatch_variant
from sw_part import extrude_cut,_get_plane_name_candidates,_select_first_candidate
from sw_sketch_quality import fully_define_sketch,inspect_model_sketches
import win32com.client

stage=sys.argv[1];out=Path(sys.argv[2]).resolve()
target=out/"法兰盘_LJT04.18.SLDPRT"
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
s.Visible=True
if stage=="profile":
    assert not target.exists();m=new_document(s,"part")
else:
    m=g(s,"ActiveDoc");assert Path(g(m,"GetPathName")).resolve()==target

def relation(items,kind):
    m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData")
    for i,item in enumerate(items):assert item.Select4(i>0,data)
    m.SketchAddConstraints(kind)

def origin(point):
    m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData")
    assert point.Select4(False,data)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
    m.SketchAddConstraints("sgCOINCIDENT")

def transform():
    t=g(g(m.SketchManager,"ActiveSketch"),"ModelToSketchTransform")
    return list(g(t,"ArrayData")),list(g(g(t,"Inverse"),"ArrayData"))

def local(p,matrix):
    p=[mm(v) for v in p]
    q=[sum(p[j]*matrix[j*3+i] for j in range(3))*matrix[12]+matrix[9+i] for i in range(3)]
    assert abs(q[2])<1e-8,q
    return q

def name_consumed_sketch(feature,name):
    child=g(feature,"GetFirstSubFeature")
    while child is not None:
        if g(child,"GetTypeName2")=="ProfileFeature":
            child.Name=name
            return
        child=g(child,"GetNextSubFeature")
    raise RuntimeError("无法回读特征消耗的草图")

if stage=="profile":
    # 左端Y0，法兰Y3～15，右端Y42。退刀槽两向1、3及R1。
    points=[(21,0),(35,0),(35,3),(65,3),(65,15),(30.5,15),(29.5,14),
            (27.5,14),(26.5,15),(26.5,17),(27.5,18),(27.5,42),(21,42)]
    with sketch(m,"Front Plane") as ref:
        sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
        try:
            axis=sm.CreateCenterLine(0,0,0,0,mm(42),0)
            edges=[]
            for i,(a,b) in enumerate(zip(points,points[1:]+points[:1])):
                if i==7:
                    e=sm.CreateArc(mm(27.5),mm(15),0,mm(a[0]),mm(a[1]),0,mm(b[0]),mm(b[1]),0,-1)
                else:e=sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1]))
                assert e is not None;edges.append(e)
        finally:sm.AddToDB=old
        for a,b in zip(edges,edges[1:]+edges[:1]):
            relation([g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
        for i,(e,a,b) in enumerate(zip(edges,points,points[1:]+points[:1])):
            if i!=7 and (a[0]==b[0] or a[1]==b[1]):
                relation([e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
        relation([edges[6],edges[7]],"sgTANGENT")
        relation([edges[7],edges[8]],"sgTANGENT")
        relation([axis],"sgVERTICAL");origin(g(axis,"GetStartPoint2"))
        relation([g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([g(axis,"GetEndPoint2"),g(edges[11],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([g(edges[7],"GetCenterPoint2"),g(edges[4],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([g(edges[7],"GetCenterPoint2"),g(edges[10],"GetStartPoint2")],"sgVERTICALPOINTS2D")
        dimension_existing_relations(m)
        for e,p in zip(edges,points):
            point=g(e,"GetStartPoint2")
            assert math.dist([g(point,"X")*1000,g(point,"Y")*1000],p)<1e-5
    m.FeatureByName(str(ref)).Name="轴向截面"
elif stage=="holes":
    m.FeatureByName("旋转1").Name="主体回转"
    # 相同C2合为一个原生倒角，避免把45°斜线拆成两个独立坐标尺寸。
    chamfer_edges(m,[([0,0,0],21),([0,42,0],21),([0,42,0],27.5),([0,15,0],65)],2,"图示C2")
    plane=offset_plane(m,"Top Plane",15,"法兰后端面")
    center=42.5/math.sqrt(2)
    with sketch(m,plane) as ref:
        matrix,inv=transform();p=local((center,15,center),matrix)
        old=m.SketchManager.AddToDB;m.SketchManager.AddToDB=True
        try:
            circle=sketch_circle(m,p[0],p[1],mm(3.5))
            radial=m.SketchManager.CreateCenterLine(0,0,0,p[0],p[1],0)
            horizontal=m.SketchManager.CreateCenterLine(0,0,0,p[0],0,0)
        finally:m.SketchManager.AddToDB=old
        origin(g(radial,"GetStartPoint2"));origin(g(horizontal,"GetStartPoint2"))
        relation([g(radial,"GetEndPoint2"),g(circle,"GetCenterPoint2")],"sgCOINCIDENT")
        relation([horizontal],"sgHORIZONTAL")
        relation([g(horizontal,"GetEndPoint2"),g(circle,"GetCenterPoint2")],"sgVERTICALPOINTS2D")
        old_prompt=s.GetUserPreferenceToggle(10);s.SetUserPreferenceToggle(10,False)
        try:
            data=g(m.SelectionManager,"CreateSelectData")
            m.ClearSelection2(True);assert radial.Select4(False,data)
            dd=m.AddDimension2(mm(36),mm(15),mm(20));assert dd is not None
            dim=g(dd,"GetDimension2",0);assert abs(g(dim,"SystemValue")-mm(42.5))<1e-7
            m.ClearSelection2(True);assert horizontal.Select4(False,data) and radial.Select4(True,data)
            dd=m.AddDimension2(mm(20),mm(15),mm(8));assert dd is not None
            dim=g(dd,"GetDimension2",0);assert abs(g(dim,"SystemValue")-math.pi/4)<1e-7
            m.ClearSelection2(True);assert circle.Select4(False,data)
            dd=m.AddDimension2(mm(43),mm(15),mm(35));assert dd is not None
            assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-mm(7))<1e-7
        finally:s.SetUserPreferenceToggle(10,old_prompt)
        assert g(g(m.SketchManager,"ActiveSketch"),"GetConstrainedStatus")==3
        mother_center=g(circle,"GetCenterPoint2")
    f=extrude_cut(m,ref,mm(12),direction=inv[7]*(-1)>0);assert f is not None
    f.Name="安装通孔";name_consumed_sketch(f,"安装孔位")
    with sketch(m,plane) as ref:
        matrix,inv=transform();p=local((center,15,center),matrix)
        e=sketch_circle(m,p[0],p[1],mm(6));assert e is not None
        relation([g(e,"GetCenterPoint2"),mother_center],"sgCOINCIDENT")
        dimension_existing_relations(m)
    f=extrude_cut(m,ref,mm(5),direction=inv[7]*(-1)>0);assert f is not None
    f.Name="安装沉孔";name_consumed_sketch(f,"沉孔轮廓")
    m.ClearSelection2(True)
    assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Front Plane"),"PLANE")
    assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Right Plane"),"PLANE",append=True)
    assert m.InsertAxis2(True)
    f=g(m,"FirstFeature")
    while f is not None:
        if g(f,"GetTypeName2")=="RefAxis":f.Name="回转轴"
        f=g(f,"GetNextFeature")
    m.ClearSelection2(True);assert m.FeatureByName("安装通孔").Select2(False,4)
    assert m.FeatureByName("安装沉孔").Select2(True,4)
    assert m.Extension.SelectByID2("回转轴","AXIS",0,0,0,True,1,create_empty_dispatch_variant(),0)
    f=m.FeatureManager.FeatureCircularPattern4(4,2*math.pi,False,"",False,True,False)
    assert f is not None;f.Name="四组安装孔阵列"
    circle_cut(m,s,plane,[(0,15,57),(0,15,-57)],2.5,12,-1,"两定位孔")
elif stage=="finish":
    plane="法兰后端面"
    with sketch(m,plane) as ref:
        matrix,inv=transform();old=m.SketchManager.AddToDB;m.SketchManager.AddToDB=True
        try:
            groups=[]
            for sign in [-1,1]:
                world=[(50*sign,15,-80),(80*sign,15,-80),(80*sign,15,80),(50*sign,15,80)]
                points=[local(p,matrix) for p in world]
                edges=[sketch_line(m,*a[:2],*b[:2]) for a,b in zip(points,points[1:]+points[:1])]
                groups.append(edges)
        finally:m.SketchManager.AddToDB=old
        for edges in groups:
            for a,b in zip(edges,edges[1:]+edges[:1]):
                relation([g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
            for i,e in enumerate(edges):
                relation([e],"sgHORIZONTAL" if i%2==0 else "sgVERTICAL")
        dimension_existing_relations(m)
    f=extrude_cut(m,ref,mm(12),direction=inv[7]*(-1)>0);assert f is not None
    f.Name="两侧削平";name_consumed_sketch(f,"削平轮廓")
    # 径向M6位于距后端13.5处，Z正侧；只钻透上壁，不触及对侧。
    plane=offset_plane(m,"Front Plane",27.5,"径向孔入口")
    with sketch(m,plane) as ref:
        matrix,inv=transform();p=local((0,28.5,27.5),matrix)
        e=sketch_circle(m,p[0],p[1],mm(2.5));assert e is not None
        fully_define_sketch(m)
    f=extrude_cut(m,ref,mm(7),direction=inv[8]*(-1)>0);assert f is not None
    f.Name="M6径向底孔";name_consumed_sketch(f,"径向孔位")
    # 曲面相交入口的原生装饰螺纹单独核实，不能将底孔冒称完整M6。
    faces=[]
    for body in g(m,"GetBodies2",0,False):
        for face in g(body,"GetFaces"):
            surface=g(face,"GetSurface")
            if g(surface,"IsCylinder"):
                params=g(surface,"CylinderParams")
                if abs(params[6]-.0025)<1e-8 and abs(params[1]-.0285)<1e-8:faces.append(face)
    assert len(faces)==1
    m.ClearSelection2(True)
    data=g(m.SelectionManager,"CreateSelectData")
    assert faces[0].Select4(False,data)
    thread=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M6x1",mm(6),0,mm(6.5),"M6，径向通上壁")
    attempts=[{"selection":"cylindrical_face","created":thread is not None}]
    if thread is None:
        edges=list(g(faces[0],"GetEdges"))
        edge=max(edges,key=lambda edge:max(g(edge,"GetCurveParams3").EndPoint))
        m.ClearSelection2(True);assert edge.Select4(False,data)
        thread=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M6x1",mm(6),0,mm(6.5),"M6，径向通上壁")
        attempts.append({"selection":"upper_intersection_edge","created":thread is not None})
    if thread is not None:thread.Name="M6装饰螺纹"
    (out/"thread-attempts.json").write_text(json.dumps(attempts,ensure_ascii=False,indent=2),encoding="utf-8")
    for name in ["法兰后端面","回转轴","径向孔入口"]:
        m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,0);m.BlankRefGeom()
else:raise ValueError(stage)
assert m.ForceRebuild3(False)
assert save_document(m,str(target))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
quality=inspect_model_sketches(m)
(out/f"{stage}.json").write_text(json.dumps({"finished":datetime.now().astimezone().isoformat(),"quality":quality},ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"stage":stage,"sketch_count":len(quality["sketches"]),"quality":quality["status"]},ensure_ascii=False))
