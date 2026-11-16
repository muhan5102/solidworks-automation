"""独立回读BRep与草图，按原图核对定位套。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_sketch_quality import inspect_model_sketches
from sw_review import collect_model_summary, save_preview
import win32com.client

out=Path(sys.argv[1]).resolve()
sw=win32com.client.Dispatch("SldWorks.Application.32")
assert g(sw,"GetProcessID")==34092
m=g(sw,"ActiveDoc")
assert Path(g(m,"GetPathName")).resolve()==out/"定位套_LJT03.28.SLDPRT"
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m)
health=collect_model_summary(m)
faces=[]
for body in g(m,"GetBodies2",0,False):
    for face in g(body,"GetFaces"):
        surf=g(face,"GetSurface")
        row={"box_mm":[x*1000 for x in g(face,"GetBox")],"area_mm2":g(face,"GetArea")*1e6}
        for name,test,params in [("cylinder","IsCylinder","CylinderParams"),
                                ("cone","IsCone","ConeParams"),("torus","IsTorus","TorusParams"),
                                ("plane","IsPlane","PlaneParams")]:
            if g(surf,test):row[name]=list(g(surf,params))
        faces.append(row)
def near(a,b):return abs(a-b)<1e-5
def cyl(radius,x,z,lo,hi):
    return [f for f in faces if "cylinder" in f and
            near(f["cylinder"][6]*1000,radius) and near(f["cylinder"][0]*1000,x)
            and near(f["cylinder"][2]*1000,z) and near(f["box_mm"][1],lo)
            and near(f["box_mm"][4],hi)]
cones=[f for f in faces if "cone" in f]
tori=[f for f in faces if "torus" in f]
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "全部草图完全定义无Fix无孤立草图":quality["status"]=="pass",
    "无特征错误警告":not health["faulty_features"] and not health["warning_features"],
    "左端D100定位圆柱Y2至15":len(cyl(50,0,0,2,15))==1,
    "法兰D190厚33位于Y15至48":len(cyl(95,0,0,15,48))==1,
    "外槽D100宽22位于Y93至115":len(cyl(50,0,0,93,115))==1,
    "右端D100圆柱Y150至188":len(cyl(50,0,0,150,188))==1,
    "左主孔D74至内槽Y160":len(cyl(37,0,0,0,160))==1,
    "内槽D82宽8位于Y160至168":len(cyl(41,0,0,160,168))==1,
    "右螺纹区起点Y168及C2起点Y188":len(cyl((80-1.0825317547305484*2)/2,0,0,168,188))==1,
    "法兰削平距轴70且只切厚33":any("plane" in f and near(f["box_mm"][2],-70)
        and near(f["box_mm"][5],-70) and near(f["box_mm"][1],15)
        and near(f["box_mm"][4],48) for f in faces),
    "锥肩对轴30度终点Y150":any(near(math.tan(f["cone"][7]),math.tan(math.pi/6))
        and near(f["box_mm"][1],150-5/math.tan(math.pi/6))
        and near(f["box_mm"][4],150) for f in cones),
    "退刀槽底D108_8位于Y48_4至51_4":len(cyl(54.4,0,0,48.4,51.4))==1,
    "退刀槽R1且位置正确":len(tori)==1 and near(tori[0]["torus"][7]*1000,1)
        and near(tori[0]["torus"][6]*1000,55.4)
        and near(tori[0]["box_mm"][1],47.4) and near(tori[0]["box_mm"][4],48.4),
    "退刀槽垂直侧Y47_4":any("plane" in f and near(f["box_mm"][1],47.4)
        and near(f["box_mm"][4],47.4) for f in faces),
    "两侧退刀斜面45度":sum(near(math.tan(f["cone"][7]),1) and
        (near(f["box_mm"][1],47.4) or near(f["box_mm"][1],51.4)) for f in cones)==2,
    "左外端C2":any(near(f["box_mm"][1],0) and near(f["box_mm"][4],2)
        and near(math.tan(f["cone"][7]),1) for f in cones),
    "右端内外两处C2":sum(near(f["box_mm"][1],188) and near(f["box_mm"][4],190)
        and near(math.tan(f["cone"][7]),1) for f in cones)==2,
    "原生M80装饰螺纹存在":m.FeatureByName("M80内螺纹") is not None,
}
for index,(x,z) in enumerate([(0,78),(-78*math.sqrt(3)/2,-39),(78*math.sqrt(3)/2,-39)],1):
    checks[f"第{index}孔D12_5从沉孔底Y23贯穿至48"]=len(cyl(6.25,x,z,23,48))==1
    checks[f"第{index}沉孔D20入口Y15深8"]=len(cyl(10,x,z,15,23))==1
report={"checks":checks,"passed":all(checks.values()),"faces":faces,"sketches":quality,"health":health,
        "finished":datetime.now().astimezone().isoformat(),
        "limits":"用户提示确认螺纹段30-8=22；M80为基本小径圆柱+原生装饰螺纹，无实体牙型。未生成工程图、未检验STEP重导入。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
for name,view in [("定位套_等轴测.bmp","isometric"),("定位套_端视.bmp","top"),("定位套_侧视.bmp","front")]:
    save_preview(m,str(out/name),view)
m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(json.dumps({"passed":report["passed"],"checks":checks},ensure_ascii=False))
assert report["passed"]
