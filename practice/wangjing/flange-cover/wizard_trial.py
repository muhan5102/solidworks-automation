"""LJT04.18孔向导替换实验，仅操作独立副本；API参数据本机CHM核对。"""
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g,create_empty_dispatch_variant

out=Path(sys.argv[1]).resolve()
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
m=g(s,"ActiveDoc")
assert Path(g(m,"GetPathName")).resolve()==out/"法兰盘_LJT04.18_孔向导.SLDPRT"
start=time.perf_counter()
for name in ["四组安装孔阵列","安装沉孔","沉孔轮廓","安装通孔","安装孔位"]:
    f=m.FeatureByName(name)
    if f is not None:
        m.ClearSelection2(True)
        assert f.Select2(False,0)
        assert m.Extension.DeleteSelection2(0),name
assert m.ForceRebuild3(False)
c=.0425/math.sqrt(2)
m.ClearSelection2(True)
assert m.Extension.SelectByRay(c,.025,c,0,-1,0,.0001,2,False,0,0)
# Counterbore、ISO、SocketHeadCap、M6；通孔D7，沉孔D12深5，头部余量0。
# Value1/2为沉孔径/深，Value6..11为额外锥沉孔，均显式置零。
f=m.FeatureManager.HoleWizard5(
    0,8,139,"M6",1,.007,.012,-1,
    .012,.005,0,-1,-1,0,0,0,0,0,0,-1,
    "",False,False,True,False,False,False)
assert f is not None,"HoleWizard5未创建"
f.Name="安装沉孔向导"
d=g(f,"GetDefinition")
properties={name:g(d,name) for name in [
    "Type","Standard2","FastenerSize","Diameter","HoleDiameter","CounterBoreDiameter",
    "CounterBoreDepth","EndCondition","HeadClearance","NearSideCounterSink","FarSideCounterSink"]}
children=[]
sf=g(f,"GetFirstSubFeature")
while sf is not None:
    row={"name":g(sf,"Name"),"type":g(sf,"GetTypeName2")}
    if row["type"] in ["ProfileFeature","3DProfileFeature"]:
        sk=g(sf,"GetSpecificFeature2")
        row["points"]=[{"x":g(p,"X"),"y":g(p,"Y"),"z":g(p,"Z")} for p in (g(sk,"GetSketchPoints2") or [])]
        row["segments"]=len(g(sk,"GetSketchSegments") or [])
    children.append(row)
    sf=g(sf,"GetNextSubFeature")
assert m.ForceRebuild3(False)
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
report={"stage_seconds":time.perf_counter()-start,"properties":properties,"children":children}
(out/"wizard-create.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(report,ensure_ascii=False))
