"""LJT04.22：球面台阶回转＋原生锥形沉头孔＋孔群阵列。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/"year08/connector"))
from build_connector import (member as g,mm,sketch,sketch_line,new_document,
    save_document,dimension_existing_relations,chamfer_edges,circle_cut)
from sw_connect import create_empty_dispatch_variant
from sw_part import revolve_boss,_get_plane_name_candidates,_select_first_candidate

out=Path(sys.argv[1]).resolve();target=out/(sys.argv[2] if len(sys.argv)>2 else "调节盘_LJT04.22.SLDPRT")
assert not target.exists()
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
s.Visible=True;m=new_document(s,"part");start=time.perf_counter()
def relation(items,kind):
    m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData")
    for i,item in enumerate(items):assert item.Select4(i>0,data)
    m.SketchAddConstraints(kind)
def origin(point):
    m.ClearSelection2(True);data=g(m.SelectionManager,"CreateSelectData")
    assert point.Select4(False,data)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
    m.SketchAddConstraints("sgCOINCIDENT")
def progress(stage):
    assert m.ForceRebuild3(False)
    assert save_document(m,str(target))
    m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
    print(f"[{stage}] 已保存，用时{time.perf_counter()-start:.2f}秒",flush=True)
    (out/"checkpoint.json").write_text(json.dumps({"stage":stage,"saved":datetime.now().astimezone().isoformat(),
        "file":target.name,"next":"继续孔群或最终核验"},ensure_ascii=False,indent=2),encoding="utf-8")

rs0=math.sqrt(30**2-22**2);rs18=math.sqrt(30**2-4**2)
points=[(16,0),(rs0,0),(rs18,18),(29,18),(29,22),(36,22),(36,31),
        (41,36),(93.5,36),(93.5,46),(27,46),(27,29),(23,29),(23,14),(16,14)]
with sketch(m,"Front Plane") as ref:
    sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
    try:
        axis=sm.CreateCenterLine(0,0,0,0,mm(46),0);edges=[]
        for i,(a,b) in enumerate(zip(points,points[1:]+points[:1])):
            if i==1:e=sm.CreateArc(0,mm(22),0,mm(a[0]),mm(a[1]),0,mm(b[0]),mm(b[1]),0,1)
            elif i==6:e=sm.CreateArc(mm(41),mm(31),0,mm(a[0]),mm(a[1]),0,mm(b[0]),mm(b[1]),0,-1)
            else:e=sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1]))
            assert e is not None;edges.append(e)
    finally:sm.AddToDB=old
    for a,b in zip(edges,edges[1:]+edges[:1]):
        relation([g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
    for i,(e,a,b) in enumerate(zip(edges,points,points[1:]+points[:1])):
        if i not in [1,6] and (a[0]==b[0] or a[1]==b[1]):
            relation([e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
    origin(g(axis,"GetStartPoint2"));relation([axis],"sgVERTICAL")
    relation([g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
    relation([g(axis,"GetEndPoint2"),g(edges[9],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
    relation([g(edges[1],"GetCenterPoint2"),axis],"sgCOINCIDENT")
    relation([g(edges[1],"GetCenterPoint2"),g(edges[4],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
    relation([edges[5],edges[6]],"sgTANGENT");relation([edges[6],edges[7]],"sgTANGENT")
    dimension_existing_relations(m)
    for e,p in zip(edges,points):
        q=g(e,"GetStartPoint2");assert math.dist([g(q,"X")*1000,g(q,"Y")*1000],p)<1e-5
m.FeatureByName(str(ref)).Name="球面轴向截面"
f=revolve_boss(m,"球面轴向截面",2*math.pi);assert f is not None;f.Name="主体回转"
chamfer_edges(m,[([0,0,0],16),([0,46,0],27),([0,36,0],93.5),([0,46,0],93.5)],1,"图示C1")
progress("主体含球面环槽和R5")

# 轴向Y，右端Y46是沉头入口；位置选择按已验证射线法，孔向负Y。
c=.085/math.sqrt(2);m.ClearSelection2(True)
assert m.Extension.SelectByRay(c,.056,c,0,-1,0,.0001,2,False,0,0)
# 本机HoleWizard5帮助：沉头Value1=直径、Value2=角度、Value3=头余量；
# Value6/7为远侧沉头，显式零。不套用M6库的默认通孔和沉头尺寸。
f=m.FeatureManager.HoleWizard5(1,8,141,"M6",1,.007,.010,-1,
    .012,math.pi/2,0,-1,-1,0,0,-1,0,-1,-1,-1,
    "",False,False,True,False,False,False)
assert f is not None;f.Name="四组沉头孔"
d=g(f,"GetDefinition")
assert abs(g(d,"CounterSinkDiameter")-.012)<1e-8
assert abs(g(d,"CounterSinkAngle")-math.pi/2)<1e-8
sf=g(f,"GetFirstSubFeature");sf.Name="沉头孔位"
position=g(sf,"GetSpecificFeature2");point=g(position,"GetSketchPoints2")[0]
g(sf,"GetNextSubFeature").Name="沉头孔向导轮廓"
m.ClearSelection2(True);assert sf.Select2(False,0);m.EditSketch()
sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
try:
    x,y=g(point,"X"),g(point,"Y")
    radial=sm.CreateCenterLine(0,0,0,x,y,0);horizontal=sm.CreateCenterLine(0,0,0,x,0,0)
finally:sm.AddToDB=old
origin(g(radial,"GetStartPoint2"));origin(g(horizontal,"GetStartPoint2"))
relation([g(radial,"GetEndPoint2"),point],"sgCOINCIDENT")
relation([horizontal],"sgHORIZONTAL");relation([g(horizontal,"GetEndPoint2"),point],"sgVERTICALPOINTS2D")
old_prompt=s.GetUserPreferenceToggle(10);s.SetUserPreferenceToggle(10,False)
try:
    data=g(m.SelectionManager,"CreateSelectData")
    m.ClearSelection2(True);assert radial.Select4(False,data)
    dd=m.AddDimension2(.075,.046,.04);assert dd is not None
    assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-.085)<1e-8
    m.ClearSelection2(True);assert horizontal.Select4(False,data) and radial.Select4(True,data)
    dd=m.AddDimension2(.04,.046,.016);assert dd is not None
    assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-math.pi/4)<1e-8
    assert g(g(sm,"ActiveSketch"),"GetConstrainedStatus")==3
finally:s.SetUserPreferenceToggle(10,old_prompt);sm.InsertSketch(True)
m.ClearSelection2(True)
assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Front Plane"),"PLANE")
assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Right Plane"),"PLANE",append=True)
assert m.InsertAxis2(True)
axis=g(m,"FirstFeature")
while axis is not None:
    if g(axis,"GetTypeName2")=="RefAxis":axis.Name="回转轴"
    axis=g(axis,"GetNextFeature")
def pattern(name,count):
    m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,4)
    assert m.Extension.SelectByID2("回转轴","AXIS",0,0,0,True,1,create_empty_dispatch_variant(),0)
    p=m.FeatureManager.FeatureCircularPattern4(count,2*math.pi,False,"",False,True,False)
    assert p is not None;p.Name=name+"阵列"
pattern("四组沉头孔",4)
# 六个D6分布圆D110，60度等分，母孔位于正Z轴；四个D11为D170主轴位置。
from build_connector import offset_plane
plane=offset_plane(m,"Top Plane",46,"右端孔位面")
circle_cut(m,s,plane,[(0,46,55)],3,10,-1,"六孔母孔")
pattern("六孔母孔",6)
circle_cut(m,s,plane,[(85,46,0),(-85,46,0),(0,46,85),(0,46,-85)],5.5,10,-1,"四定位通孔")
for name in ["回转轴","右端孔位面"]:
    m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,0);m.BlankRefGeom()
progress("全部孔群")
(out/"build.json").write_text(json.dumps({"finished":datetime.now().astimezone().isoformat(),
    "execution_seconds":time.perf_counter()-start,
    "quality":"阶段断言完成；综合草图和几何检查留给重开后统一执行"},ensure_ascii=False,indent=2),encoding="utf-8")
