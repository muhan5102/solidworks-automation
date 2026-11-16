"""补孔向导位置驱动并阵列；内部轮廓由原生孔参数管理，不手改其生成关系。"""
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g,create_empty_dispatch_variant
from sw_sketch_quality import inspect_model_sketches

out=Path(sys.argv[1]).resolve();start=time.perf_counter()
s=win32com.client.Dispatch("SldWorks.Application.32");m=g(s,"ActiveDoc")
assert g(s,"GetProcessID")==12356
assert Path(g(m,"GetPathName")).resolve()==out/"法兰盘_LJT04.18_孔向导.SLDPRT"
wizard=m.FeatureByName("安装沉孔向导")
sf=g(wizard,"GetFirstSubFeature");sk=g(sf,"GetSpecificFeature2")
assert len(g(sk,"GetSketchSegments") or [])==0
point=g(sk,"GetSketchPoints2")[0]
sf.Name="安装孔位"
profile=g(sf,"GetNextSubFeature");profile.Name="沉孔向导轮廓"
m.ClearSelection2(True);assert sf.Select2(False,0);m.EditSketch()
sm=m.SketchManager;data=g(m.SelectionManager,"CreateSelectData")
def relation(items,kind):
    m.ClearSelection2(True)
    for i,item in enumerate(items):assert item.Select4(i>0,data)
    m.SketchAddConstraints(kind)
def origin(p):
    m.ClearSelection2(True);assert p.Select4(False,data)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
    m.SketchAddConstraints("sgCOINCIDENT")
old=sm.AddToDB;sm.AddToDB=True
try:
    x,y=g(point,"X"),g(point,"Y")
    radial=sm.CreateCenterLine(0,0,0,x,y,0)
    horizontal=sm.CreateCenterLine(0,0,0,x,0,0)
finally:sm.AddToDB=old
origin(g(radial,"GetStartPoint2"));origin(g(horizontal,"GetStartPoint2"))
relation([g(radial,"GetEndPoint2"),point],"sgCOINCIDENT")
relation([horizontal],"sgHORIZONTAL")
relation([g(horizontal,"GetEndPoint2"),point],"sgVERTICALPOINTS2D")
old_prompt=s.GetUserPreferenceToggle(10);s.SetUserPreferenceToggle(10,False)
try:
    m.ClearSelection2(True);assert radial.Select4(False,data)
    dd=m.AddDimension2(.036,.015,.02);assert dd is not None
    assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-.0425)<1e-8
    m.ClearSelection2(True);assert horizontal.Select4(False,data) and radial.Select4(True,data)
    dd=m.AddDimension2(.020,.015,.008);assert dd is not None
    assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-math.pi/4)<1e-8
    assert g(g(sm,"ActiveSketch"),"GetConstrainedStatus")==3
finally:
    s.SetUserPreferenceToggle(10,old_prompt)
    sm.InsertSketch(True)
m.ClearSelection2(True);assert wizard.Select2(False,4)
assert m.Extension.SelectByID2("回转轴","AXIS",0,0,0,True,1,create_empty_dispatch_variant(),0)
f=m.FeatureManager.FeatureCircularPattern4(4,2*math.pi,False,"",False,True,False)
assert f is not None;f.Name="四组安装孔阵列"
assert m.ForceRebuild3(False)
err=win32com.client.VARIANT(16387,0);warn=win32com.client.VARIANT(16387,0)
assert m.Save3(1,err,warn);assert err.value==0
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
report={"seconds":time.perf_counter()-start,"quality":inspect_model_sketches(m)}
(out/"wizard-finish.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(report,ensure_ascii=False))
