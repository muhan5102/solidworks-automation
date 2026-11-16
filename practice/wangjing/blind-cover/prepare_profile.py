"""仅准备闷盖参数化截面；实体回转交由正式MCP执行。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]/"year08/connector"))
from build_connector import (member, mm, sketch, sketch_line, new_document,
                             dimension_existing_relations, save_document)
from sw_connect import create_empty_dispatch_variant
from sw_sketch_quality import inspect_model_sketches
import win32com.client

out = Path(sys.argv[1]).resolve()
name = sys.argv[2] if len(sys.argv)>2 else "闷盖_LJT04.32.SLDPRT"
assert Path(name).name==name and Path(name).suffix.lower()==".sldprt"
dest = out/name
assert out.is_dir() and not dest.exists()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == 34092
sw.Visible = True
m = new_document(sw, "part")
center_r = 26-.05*14-2*math.sqrt(1+.05**2)
tangent_r = center_r+2/math.sqrt(1+.05**2)
tangent_y = 14+.1/math.sqrt(1+.05**2)
points = [(26,0),(29.5,0),(31,1.5),(31,16),(50,16),(50,28),
          (31,28),(29,26),(0,26),(0,16),(center_r,16),(tangent_r,tangent_y)]
with sketch(m, "Front Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0,0,0,0,mm(28),0)
        edges = []
        for i,(a,b) in enumerate(zip(points,points[1:]+points[:1])):
            if i in (6,10):
                c = (29,28) if i == 6 else (center_r,14)
                edges.append(sm.CreateArc(mm(c[0]),mm(c[1]),0,
                    mm(a[0]),mm(a[1]),0,mm(b[0]),mm(b[1]),0,-1))
            else:
                edges.append(sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1])))
    finally:
        sm.AddToDB = old
    assert axis is not None and all(e is not None for e in edges)
    sd = member(m.SelectionManager,"CreateSelectData")
    for a,b in zip(edges,edges[1:]+edges[:1]):
        m.ClearSelection2(True)
        assert member(a,"GetEndPoint2").Select4(False,sd)
        assert member(b,"GetStartPoint2").Select4(True,sd)
        m.SketchAddConstraints("sgCOINCIDENT")
    for i,(e,a,b) in enumerate(zip(edges,points,points[1:]+points[:1])):
        if i not in (6,10) and (a[0]==b[0] or a[1]==b[1]):
            m.ClearSelection2(True)
            assert e.Select4(False,sd)
            m.SketchAddConstraints("sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
    for arc,line in [(edges[6],edges[7]),(edges[10],edges[9]),(edges[10],edges[11])]:
        m.ClearSelection2(True)
        assert arc.Select4(False,sd) and line.Select4(True,sd)
        m.SketchAddConstraints("sgTANGENT")
    m.ClearSelection2(True)
    assert member(edges[6],"GetCenterPoint2").Select4(False,sd)
    assert member(edges[6],"GetStartPoint2").Select4(True,sd)
    m.SketchAddConstraints("sgHORIZONTALPOINTS2D")
    m.ClearSelection2(True)
    assert axis.Select4(False,sd)
    m.SketchAddConstraints("sgVERTICAL")
    m.ClearSelection2(True)
    assert member(axis,"GetStartPoint2").Select4(False,sd)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,
                                  create_empty_dispatch_variant(),0)
    m.SketchAddConstraints("sgCOINCIDENT")
    m.ClearSelection2(True)
    assert axis.Select4(False,sd) and edges[8].Select4(True,sd)
    m.SketchAddConstraints("sgCOLINEAR")
    m.ClearSelection2(True)
    assert member(edges[0],"GetStartPoint2").Select4(False,sd)
    assert member(axis,"GetStartPoint2").Select4(True,sd)
    m.SketchAddConstraints("sgHORIZONTALPOINTS2D")
    dimension_existing_relations(m)
    for e,p in zip(edges,points):
        actual = member(e,"GetStartPoint2")
        assert math.dist([member(actual,"X")*1000,member(actual,"Y")*1000],p)<1e-5
    contours = list(member(member(sm,"ActiveSketch"),"GetSketchContours") or [])
    assert len(contours)==1 and member(contours[0],"IsClosed")
sketch_feature=m.FeatureByName(str(ref))
sketch_feature.Name="轴向截面"
assert save_document(m,str(dest))
quality=inspect_model_sketches(m)
# 此阶段草图尚未被MCP回转消耗，整体审查会按孤立草图标review_required；
# 这里只验约束，回转后再验无孤立草图。
assert len(quality["sketches"])==1
assert all(s["constrained_status"]==3 and s["fix_count"]==0 for s in quality["sketches"])
m.ShowNamedView2("*Front",1)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
result={"path":str(dest),"sketch_name":"轴向截面","axis_sketch_name":"轴向截面",
        "profile_points_mm":points,"sketch_quality":quality,
        "finished":datetime.now().astimezone().isoformat()}
(out/"profile.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"path":str(dest),"quality":quality["status"]},ensure_ascii=False))
