"""法兰盘LJT04.17；MCP回转，API合并阵列及原图角度驱动。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/"year08/connector"))
from build_connector import (member as g, mm, sketch, sketch_line, new_document,
    dimension_existing_relations, save_document, offset_plane, circle_cut)
from sw_part import _get_plane_name_candidates, _select_first_candidate
from sw_connect import create_empty_dispatch_variant
from sw_sketch_quality import inspect_model_sketches, inspect_sketch
import win32com.client

stage=sys.argv[1];out=Path(sys.argv[2]).resolve()
dest=out/"法兰盘_LJT04.17.SLDPRT"
assert out.is_dir()
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==34092
s.Visible=True

def define_notch_by_angle(model):
    """将完全定义圆的横纵定位替换为80mm连线长度和15°角度驱动。"""
    feature=model.FeatureByName("孔位草图_R30偏置缺口")
    assert feature is not None and feature.Select2(False,0)
    model.EditSketch()
    active=g(model.SketchManager,"ActiveSketch")
    sd=g(model.SelectionManager,"CreateSelectData")
    old_prompt=s.GetUserPreferenceToggle(10)  # 本机类型库swInputDimValOnCreate
    s.SetUserPreferenceToggle(10,False)
    try:
        for name in ["D2@孔位草图_R30偏置缺口","D3@孔位草图_R30偏置缺口"]:
            model.ClearSelection2(True)
            assert model.Extension.SelectByID2(name,"DIMENSION",0,0,0,False,0,create_empty_dispatch_variant(),0)
            assert model.Extension.DeleteSelection2(0)
        circle=next(e for e in g(active,"GetSketchSegments") if abs(g(e,"GetRadius")-.03)<1e-8)
        center=g(circle,"GetCenterPoint2");x=g(center,"X");y=g(center,"Y")
        old_db=model.SketchManager.AddToDB
        model.SketchManager.AddToDB=True
        try:
            polar=model.SketchManager.CreateCenterLine(0,0,0,x,y,0)
            vertical=model.SketchManager.CreateCenterLine(0,0,0,0,y,0)
        finally:model.SketchManager.AddToDB=old_db
        def relation(items,kind):
            model.ClearSelection2(True)
            for i,item in enumerate(items):assert item.Select4(i>0,sd)
            model.SketchAddConstraints(kind)
        for line in [polar,vertical]:
            model.ClearSelection2(True)
            assert g(line,"GetStartPoint2").Select4(False,sd)
            assert model.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
            model.SketchAddConstraints("sgCOINCIDENT")
        relation([g(polar,"GetEndPoint2"),center],"sgCOINCIDENT")
        relation([vertical],"sgVERTICAL")
        relation([g(vertical,"GetEndPoint2"),center],"sgHORIZONTALPOINTS2D")
        model.ClearSelection2(True);assert polar.Select4(False,sd)
        dim=model.AddDimension2(.035,.018,-.032)
        assert dim is not None
        length=g(dim,"GetDimension2",0)
        assert abs(g(length,"SystemValue")-.08)<1e-7
        length.SystemValue=.08
        model.ClearSelection2(True)
        assert vertical.Select4(False,sd) and polar.Select4(True,sd)
        # AddDimension2的文字位置使用模型坐标；不能把局部Y直接当轴向Y。
        dim=model.AddDimension2(.012,.018,-.05)
        assert dim is not None
        angle=g(dim,"GetDimension2",0)
        assert abs(g(angle,"SystemValue")-math.pi/12)<1e-6
        angle.SystemValue=math.pi/12
        assert inspect_sketch(active)["status"]=="pass"
        assert math.dist([g(center,"X"),g(center,"Y")],[x,y])<1e-8
    finally:
        s.SetUserPreferenceToggle(10,old_prompt)
        if g(model.SketchManager,"ActiveSketch") is not None:
            model.SketchManager.InsertSketch(True)

if stage=="profile":
    assert not dest.exists();m=new_document(s,"part")
    points=[(30,0),(35,0),(35,18),(60,18),(60,38),(34.5,38),(34.5,40),
            (35,40),(35,44),(34,45),(15,45),(15,25),(31,25),(31,21),(30,21)]
    with sketch(m,"Front Plane") as ref:
        old=m.SketchManager.AddToDB;m.SketchManager.AddToDB=True
        try:
            axis=m.SketchManager.CreateCenterLine(0,0,0,0,mm(45),0)
            edges=[sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1]))
                   for a,b in zip(points,points[1:]+points[:1])]
        finally:m.SketchManager.AddToDB=old
        assert axis is not None and all(e is not None for e in edges)
        sd=g(m.SelectionManager,"CreateSelectData")
        def relation(items,kind):
            m.ClearSelection2(True)
            for i,e in enumerate(items):assert e.Select4(i>0,sd)
            m.SketchAddConstraints(kind)
        for a,b in zip(edges,edges[1:]+edges[:1]):
            relation([g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
        for e,a,b in zip(edges,points,points[1:]+points[:1]):
            if a[0]==b[0] or a[1]==b[1]:
                relation([e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
        relation([axis],"sgVERTICAL")
        m.ClearSelection2(True);assert g(axis,"GetStartPoint2").Select4(False,sd)
        assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
        m.SketchAddConstraints("sgCOINCIDENT")
        relation([g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([g(axis,"GetEndPoint2"),g(edges[9],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
        relation([edges[1],edges[7]],"sgCOLINEAR")
        dimension_existing_relations(m)
        for e,p in zip(edges,points):
            q=g(e,"GetStartPoint2")
            assert math.dist([g(q,"X")*1000,g(q,"Y")*1000],p)<1e-5
        contours=g(g(m.SketchManager,"ActiveSketch"),"GetSketchContours")
        assert len(contours)==1 and g(contours[0],"IsClosed")
    m.FeatureByName(str(ref)).Name="轴向截面"
elif stage=="features":
    m=g(s,"ActiveDoc");assert Path(g(m,"GetPathName")).resolve()==dest
    m.FeatureByName("旋转1").Name="主体回转"
    plane=offset_plane(m,"Top Plane",18,"法兰前端面")
    circle_cut(m,s,plane,[(0,18,47.5)],5.5,20,1,"安装通孔")
    circle_cut(m,s,plane,[(0,18,47.5)],9,10,1,"安装沉孔")
    # 15°朝图面右侧偏离下方径线；R30圆弧最近点距轴50，故圆心半径80。
    circle_cut(m,s,plane,[(80*math.sin(math.pi/12),18,-80*math.cos(math.pi/12))],
               30,20,1,"R30偏置缺口")
    define_notch_by_angle(m)
    m.ClearSelection2(True)
    assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Front Plane"),"PLANE")
    assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Right Plane"),"PLANE",append=True)
    assert m.InsertAxis2(True)
    f=g(m,"FirstFeature")
    while f is not None:
        if g(f,"GetTypeName2")=="RefAxis":f.Name="回转轴"
        f=g(f,"GetNextFeature")
    for name in ["法兰前端面","回转轴"]:
        m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,0);m.BlankRefGeom()
elif stage=="pattern":
    m=g(s,"ActiveDoc");assert Path(g(m,"GetPathName")).resolve()==dest
    assert m.FeatureByName("三组安装孔阵列") is None
    m.ClearSelection2(True)
    assert m.FeatureByName("安装通孔").Select2(False,4)
    assert m.FeatureByName("安装沉孔").Select2(True,4)
    assert m.Extension.SelectByID2("回转轴","AXIS",0,0,0,True,1,create_empty_dispatch_variant(),0)
    f=m.FeatureManager.FeatureCircularPattern4(3,2*math.pi,False,"",False,True,False)
    assert f is not None;f.Name="三组安装孔阵列"
else:raise ValueError(stage)
assert m.ForceRebuild3(False)
assert save_document(m,str(dest))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
q=inspect_model_sketches(m)
(out/f"{stage}.json").write_text(json.dumps({"stage":stage,"finished":datetime.now().astimezone().isoformat(),"quality":q},ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"stage":stage,"sketches":len(q["sketches"]),"quality":q["status"]},ensure_ascii=False))
