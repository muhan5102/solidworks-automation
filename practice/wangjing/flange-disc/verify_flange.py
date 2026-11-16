"""法兰盘的轴向材料范围、端面面积、三沉孔及单侧缺口检查。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_sketch_quality import inspect_model_sketches
from sw_review import collect_model_summary,save_preview
import win32com.client

out=Path(sys.argv[1]).resolve();s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==34092
m=g(s,"ActiveDoc");assert Path(g(m,"GetPathName")).resolve()==out/"法兰盘_LJT04.17.SLDPRT"
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m);health=collect_model_summary(m)
faces=[]
for b in g(m,"GetBodies2",0,False):
    for f in g(b,"GetFaces"):
        surf=g(f,"GetSurface")
        row={"box_mm":[v*1000 for v in g(f,"GetBox")],"area_mm2":g(f,"GetArea")*1e6}
        for name,test,params in [("cylinder","IsCylinder","CylinderParams"),("cone","IsCone","ConeParams"),("plane","IsPlane","PlaneParams")]:
            if g(surf,test):row[name]=list(g(surf,params))
        faces.append(row)
def near(a,b):return abs(a-b)<1e-5
def cyl(r,x,z,lo,hi):
    return [f for f in faces if "cylinder" in f and near(f["cylinder"][6]*1000,r)
        and near(f["cylinder"][0]*1000,x) and near(f["cylinder"][2]*1000,z)
        and near(f["box_mm"][1],lo) and near(f["box_mm"][4],hi)]
def area(y):
    return sum(f["area_mm2"] for f in faces if "plane" in f
               and near(f["box_mm"][1],y) and near(f["box_mm"][4],y))
cut_x=80*math.sin(math.pi/12);cut_z=-80*math.cos(math.pi/12)
# 两圆交集是单侧R30缺口去除的端面面积；两端分别扣除沉孔/通孔。
R,r,d=60,30,80
overlap=(R*R*math.acos((d*d+R*R-r*r)/(2*d*R))
         +r*r*math.acos((d*d+r*r-R*R)/(2*d*r))
         -.5*math.sqrt((-d+R+r)*(d+R-r)*(d-R+r)*(d+R+r)))
cones=[f for f in faces if "cone" in f]
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "四草图完全定义无Fix无孤立":quality["status"]=="pass" and len(quality["sketches"])==4,
    "无特征错误警告":not health["faulty_features"] and not health["warning_features"],
    "左D70外圆仅Y0至18":len(cyl(35,0,0,0,18))==1,
    "法兰D120仅Y18至38":len(cyl(60,0,0,18,38))==1,
    "右D70外圆Y40至44":len(cyl(35,0,0,40,44))==1,
    "外退刀槽2乘0_5位置38至40":len(cyl(34.5,0,0,38,40))==1,
    "左D60孔深21":len(cyl(30,0,0,0,21))==1,
    "内槽4乘1位置21至25":len(cyl(31,0,0,21,25))==1,
    "右D30孔仅Y25至45":len(cyl(15,0,0,25,45))==1,
    "R30缺口仅一个且位于右下15度":len(cyl(30,cut_x,cut_z,18,38))==1,
    "左端材料环面积正确":near(area(0),math.pi*(35**2-30**2)),
    "右端材料环面积正确无多余材料":near(area(45),math.pi*(34**2-15**2)),
    "法兰入口面含缺口及三沉孔材料面积正确":near(area(18),math.pi*(60**2-35**2)-3*math.pi*9**2-overlap),
    "法兰背面含缺口及三通孔材料面积正确":near(area(38),math.pi*(60**2-34.5**2)-3*math.pi*5.5**2-overlap),
    "仅图示右外端C1":len(cones)==1 and near(cones[0]["box_mm"][1],44)
        and near(cones[0]["box_mm"][4],45) and near(math.tan(cones[0]["cone"][7]),1),
    "同一三组孔阵列存在且旧双阵列已移除":m.FeatureByName("三组安装孔阵列") is not None
        and m.FeatureByName("阵列(圆周)1") is None and m.FeatureByName("阵列(圆周)2") is None,
}
for i,(x,z) in enumerate([(0,47.5),(-47.5*math.sqrt(3)/2,-23.75),(47.5*math.sqrt(3)/2,-23.75)],1):
    checks[f"第{i}孔D11位于Y28至38"]=len(cyl(5.5,x,z,28,38))==1
    checks[f"第{i}沉孔D18入口Y18深10"]=len(cyl(9,x,z,18,28))==1
report={"checks":checks,"passed":all(checks.values()),"quality":quality,"health":health,"faces":faces,
        "end_face_areas_mm2":{str(y):area(y) for y in [0,18,38,45]},
        "finished":datetime.now().astimezone().isoformat(),
        "limits":"经用户提示将通孔沉孔合并为一个阵列；未出工程图，未重导入STEP。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
save_preview(m,str(out/"法兰盘_等轴测.bmp"),"isometric")
save_preview(m,str(out/"法兰盘_侧视.bmp"),"front")
m.ShowNamedView2("*Bottom",6);m.ViewZoomtofit2();m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"法兰盘_前端.bmp"),1600,1000)
m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(json.dumps({"passed":report["passed"],"checks":checks},ensure_ascii=False))
assert report["passed"]
