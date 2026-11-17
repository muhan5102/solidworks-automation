"""偏心盘LJT04.26：分阶段执行，轴套回转由正式MCP完成。"""
from datetime import datetime
import ctypes
import json
import math
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from shared.cad import (member as g, mm, sketch, sketch_line, sketch_circle,
    new_document, save_document, dimension_existing_relations, offset_plane,
    circle_cut, round_edges, chamfer_edges)
from sw_connect import create_empty_dispatch_variant
from sw_part import extrude_boss, extrude_cut
from sw_sketch_quality import fully_define_sketch, inspect_model_sketches
from sw_review import save_preview
import win32com.client

stage=sys.argv[1]
out=Path(sys.argv[2]).resolve()
dest=out/"偏心盘_LJT04.26.SLDPRT"
assert out.is_dir()
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==34092
s.Visible=True
assert ctypes.windll.user32.IsWindowVisible(69222)
if ctypes.windll.user32.IsIconic(69222):ctypes.windll.user32.ShowWindow(69222,9)
if stage=="profile":
    assert not dest.exists()
    m=new_document(s,"part")
else:
    m=g(s,"ActiveDoc")
    assert Path(g(m,"GetPathName")).resolve()==dest

def relation(items,kind):
    m.ClearSelection2(True)
    sd=g(m.SelectionManager,"CreateSelectData")
    for i,item in enumerate(items):assert item.Select4(i>0,sd)
    m.SketchAddConstraints(kind)

def closed_lines(points):
    old=m.SketchManager.AddToDB;m.SketchManager.AddToDB=True
    try:edges=[sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1]))
               for a,b in zip(points,points[1:]+points[:1])]
    finally:m.SketchManager.AddToDB=old
    for a,b in zip(edges,edges[1:]+edges[:1]):
        relation([g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
    for e,a,b in zip(edges,points,points[1:]+points[:1]):
        relation([e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
    return edges

def transform():
    t=g(g(m.SketchManager,"ActiveSketch"),"ModelToSketchTransform")
    return list(g(t,"ArrayData")),list(g(g(t,"Inverse"),"ArrayData"))

def local(point,matrix):
    p=[mm(v) for v in point]
    q=[sum(p[j]*matrix[j*3+i] for j in range(3))*matrix[12]+matrix[9+i] for i in range(3)]
    assert abs(q[2])<1e-8, q
    return q

if stage=="profile":
    with sketch(m,"Front Plane") as ref:
        # 轴套止于后盘端面70；总长76由后端6mm条形凸台取得，不能整圈延长。
        axis=m.SketchManager.CreateCenterLine(0,0,0,0,mm(70),0)
        edges=closed_lines([(18,0),(30,0),(30,70),(18,70)])
        relation([axis],"sgVERTICAL")
        m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData")
        assert g(axis,"GetStartPoint2").Select4(False,sd)
        assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
        m.SketchAddConstraints("sgCOINCIDENT")
        relation([g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([g(axis,"GetEndPoint2"),g(edges[2],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name="轴套截面"
elif stage=="body":
    m.FeatureByName("旋转1").Name="轴套回转"
    for y,radius,name in [(10,54,"前偏心盘"),(56,46,"后偏心盘")]:
        plane=offset_plane(m,"Top Plane",y,f"端面{y}")
        with sketch(m,plane) as ref:
            matrix,inv=transform()
            p=local((0,y,10),matrix);hole=local((0,y,0),matrix)
            outer=sketch_circle(m,p[0],p[1],mm(radius))
            inner=sketch_circle(m,hole[0],hole[1],mm(18))
            assert outer is not None and inner is not None
            fully_define_sketch(m)
            for segment,expected,r in [(outer,p,radius),(inner,hole,18)]:
                center=g(segment,"GetCenterPoint2")
                assert math.dist([g(center,"X"),g(center,"Y")],expected[:2])<1e-8
                assert abs(g(segment,"GetRadius")-mm(r))<1e-8
        feature=extrude_boss(m,ref,mm(14),direction=inv[7]<0)
        assert feature is not None;feature.Name=name
        m.FeatureByName(str(ref)).Name=name+"轮廓"
        print(f"[主体] {name}：偏心10，轴向{y}至{y+14}",flush=True)
    plane=offset_plane(m,"Top Plane",70,"后端凸台面")
    a=math.sqrt(76**2-14**2)
    with sketch(m,plane) as ref:
        matrix,inv=transform()
        points=[(-14,70,-28+a),(14,70,-28+a),(14,70,46-a),(-14,70,46-a)]
        q=[local(p,matrix) for p in points]
        centers=[local((0,70,-28),matrix),local((0,70,46),matrix)]
        old=m.SketchManager.AddToDB;m.SketchManager.AddToDB=True
        try:
            top=m.SketchManager.CreateArc(*centers[0],*q[0],*q[1],1)
            right=sketch_line(m,*q[1][:2],*q[2][:2])
            bottom=m.SketchManager.CreateArc(*centers[1],*q[2],*q[3],1)
            left=sketch_line(m,*q[3][:2],*q[0][:2])
        finally:m.SketchManager.AddToDB=old
        edges=[top,right,bottom,left];assert all(e is not None for e in edges)
        for ea,eb in zip(edges,edges[1:]+edges[:1]):
            relation([g(ea,"GetEndPoint2"),g(eb,"GetStartPoint2")],"sgCOINCIDENT")
        relation([right],"sgVERTICAL");relation([left],"sgVERTICAL")
        relation([top,bottom],"sgEQUAL")
        relation([g(top,"GetCenterPoint2"),g(bottom,"GetCenterPoint2")],"sgVERTICALPOINTS2D")
        m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData")
        assert g(top,"GetCenterPoint2").Select4(False,sd)
        assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
        m.SketchAddConstraints("sgVERTICALPOINTS2D")
        dimension_existing_relations(m)
        for e,point in zip(edges,q):
            p=g(e,"GetStartPoint2");assert math.dist([g(p,"X"),g(p,"Y")],point[:2])<1e-8
    f=extrude_boss(m,ref,mm(6),direction=inv[7]<0);assert f is not None
    f.Name="后端弧形凸台";m.FeatureByName(str(ref)).Name="R76凸台轮廓"
    circle_cut(m,s,plane,[(0,70,0)],18,6,1,"凸台内孔")
elif stage=="details":
    # 孔按两端实际材料范围处理，不使用一条Through All误切另一片盘。
    front="端面10"
    circle_cut(m,s,front,[(0,10,50)],6,14,1,"前盘D12孔")
    rear=offset_plane(m,"Top Plane",76,"后端面")
    circle_cut(m,s,rear,[(0,76,40)],3.4,20,-1,"M8底孔")
    with sketch(m,"Top Plane") as ref:
        matrix,inv=transform()
        corners=[(-5,0,0),(5,0,0),(5,0,21.3),(-5,0,21.3)]
        points=[[v*1000 for v in local(p,matrix)[:2]] for p in corners]
        closed_lines(points)
        fully_define_sketch(m)
    f=extrude_cut(m,ref,mm(76),direction=inv[7]>0);assert f is not None
    f.Name="贯通键槽";m.FeatureByName(str(ref)).Name="键槽轮廓"
    # 10N9、39.3由孔底(-18)至槽顶21.3；C2只加题面可见的左孔口。
    chamfer_edges(m,[([0,0,0],18)],2,"前孔口C2")
    mouths=[e for e,p,r in round_edges(m) if math.dist(p,[0,76,40])<1e-5 and abs(r-3.4)<1e-5]
    assert len(mouths)==1;m.ClearSelection2(True);assert mouths[0].Select2(False,0)
    f=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M8x1.25",mm(8),0,mm(20),"M8-6H")
    assert f is not None;f.Name="M8内螺纹"
    feature=g(m,"FirstFeature")
    while feature is not None:
        if g(feature,"GetTypeName2")=="RefPlane" and g(feature,"Name") in ["端面10","端面56","后端凸台面","后端面"]:
            m.ClearSelection2(True);feature.Select2(False,0);m.BlankRefGeom()
        feature=g(feature,"GetNextFeature")
else:raise ValueError(stage)
assert m.ForceRebuild3(False)
assert save_document(m,str(dest))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
quality=inspect_model_sketches(m)
(out/f"{stage}.json").write_text(json.dumps({"stage":stage,"finished":datetime.now().astimezone().isoformat(),"quality":quality},ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"stage":stage,"sketches":len(quality["sketches"]),"quality":quality["status"]},ensure_ascii=False))
