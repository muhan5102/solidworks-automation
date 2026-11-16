"""按LJT04.22原图独立核对球面、材料区、孔群与参数驱动。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_review import collect_model_summary,save_preview
from sw_sketch_quality import inspect_model_sketches

out=Path(sys.argv[1]).resolve();s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
m=g(s,"ActiveDoc");target=out/(sys.argv[2] if len(sys.argv)>2 else "调节盘_LJT04.22.SLDPRT")
assert Path(g(m,"GetPathName")).resolve()==target
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m);health=collect_model_summary(m)
faces=[]
for b in g(m,"GetBodies2",0,False):
    for f in g(b,"GetFaces"):
        su=g(f,"GetSurface");row={"box":[v*1000 for v in g(f,"GetBox")],"area":g(f,"GetArea")*1e6}
        for kind in ["Cylinder","Cone","Sphere","Torus","Plane"]:
            if g(su,"Is"+kind):row[kind]=list(g(su,kind+"Params"))
        faces.append(row)
def near(a,b):return abs(a-b)<1e-5
def cyl(r,x,z,y0,y1):
    return [f for f in faces if "Cylinder" in f and near(f["Cylinder"][6]*1000,r)
        and near(f["Cylinder"][0]*1000,x) and near(f["Cylinder"][2]*1000,z)
        and near(f["box"][1],y0) and near(f["box"][4],y1)]
def area(y):
    return sum(f["area"] for f in faces if "Plane" in f and near(f["box"][1],y) and near(f["box"][4],y))
def dimensions(name):
    f=m.FeatureByName(name);d=g(f,"GetFirstDisplayDimension");result=[]
    while d is not None:
        value=g(d,"GetDimension2",0);result.append({"name":g(value,"FullName"),
            "value":g(value,"SystemValue"),"driven":g(value,"DrivenState")})
        d=g(f,"GetNextDisplayDimension",d)
    return result
spheres=[f for f in faces if "Sphere" in f];tori=[f for f in faces if "Torus" in f]
cones=[f for f in faces if "Cone" in f]
generated=[sk for sk in quality["sketches"] if sk["name"]=="沉头孔向导轮廓"]
manual=[sk for sk in quality["sketches"] if sk["name"]!="沉头孔向导轮廓"]
d=g(m.FeatureByName("四组沉头孔"),"GetDefinition")
wizard={k:g(d,k) for k in ["ThruHoleDiameter","CounterSinkDiameter","CounterSinkAngle","EndCondition","FarCounterSinkDiameter"]}
hole_dims=dimensions("沉头孔位");profile_dims=dimensions("球面轴向截面")
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "无特征错误警告":not health["faulty_features"] and not health["warning_features"],
    "四张手工草图完全定义无Fix无孤立":len(manual)==4 and all(sk["status"]=="pass" for sk in manual),
    "孔向导内部截面单列核验":len(generated)==1 and generated[0]["constrained_status"]==3
        and generated[0]["dimension_count"]==4 and generated[0]["fix_count"]==1 and not generated[0]["dangling"],
    "SR30球心Y22表面限Y0至18":len(spheres)==1 and all(near(a,b) for a,b in zip(spheres[0]["Sphere"],[0,.022,0,.030]))
        and near(spheres[0]["box"][1],0) and near(spheres[0]["box"][4],18),
    "球面真实面积符合SR30乘轴向跨度18":len(spheres)==1 and near(spheres[0]["area"],2*math.pi*30*18),
    "环槽D58宽4位置Y18至22":len(cyl(29,0,0,18,22))==1,
    "轮毂D72从B面Y22至R5起点31":len(cyl(36,0,0,22,31))==1,
    "R5根部圆角位置及半径":len(tori)==1 and near(tori[0]["Torus"][1],.031)
        and near(tori[0]["Torus"][6],.041) and near(tori[0]["Torus"][7],.005),
    "法兰D187厚10两侧图示C1":len(cyl(93.5,0,0,37,45))==1,
    "D32前孔止于Y14含入口C1":len(cyl(16,0,0,1,14))==1,
    "D46中孔Y14至29深15":len(cyl(23,0,0,14,29))==1,
    "D54后孔Y29至45含出口C1":len(cyl(27,0,0,29,45))==1,
    "Y14内台阶环面积正确":near(area(14),math.pi*(23**2-16**2)),
    "Y29内台阶环面积正确":near(area(29),math.pi*(27**2-23**2)),
    "左端材料面积无多余环":near(area(0),math.pi*((30**2-22**2)-17**2)),
    "右端材料面积含全部14孔及C1":near(area(46),math.pi*(92.5**2-28**2)-4*math.pi*6**2-4*math.pi*5.5**2-6*math.pi*3**2),
    "法兰左面材料面积含R5及全部14孔":near(area(36),math.pi*(92.5**2-41**2)-4*math.pi*3.5**2-4*math.pi*5.5**2-6*math.pi*3**2),
    "锥面总数仅四沉头加四图示C1":len(cones)==8,
    "孔向导D7沉头D12全角90贯穿无远侧沉头":near(wizard["ThruHoleDiameter"],.007)
        and near(wizard["CounterSinkDiameter"],.012) and near(wizard["CounterSinkAngle"],math.pi/2)
        and wizard["EndCondition"]==1 and near(wizard["FarCounterSinkDiameter"],0),
    "沉头孔半径85及45度为驱动":all(any(near(v["value"],expected) and v["driven"]==2 for v in hole_dims) for expected in [.085,math.pi/4]),
    "两个原生孔群阵列":sum(f["type"]=="CirPattern" for f in health["features"])==2,
}
for i,(x,z) in enumerate([(a*85/math.sqrt(2),b*85/math.sqrt(2)) for a,b in [(1,1),(-1,1),(-1,-1),(1,-1)]],1):
    checks[f"沉头{i}通孔D7中心及Y36至43_5"]=len(cyl(3.5,x,z,36,43.5))==1
    rows=[f for f in cones if near(f["Cone"][0]*1000,x) and near(f["Cone"][2]*1000,z)]
    checks[f"沉头{i}位于右侧Y43_5至46全角90"]=len(rows)==1 and near(rows[0]["box"][1],43.5) \
        and near(rows[0]["box"][4],46) and near(abs(rows[0]["Cone"][7]),math.pi/4)
for i in range(6):
    angle=math.pi/2+i*math.pi/3
    checks[f"D6第{i+1}孔分布圆110贯穿10"]=len(cyl(3,55*math.cos(angle),55*math.sin(angle),36,46))==1
for i,(x,z) in enumerate([(85,0),(-85,0),(0,85),(0,-85)],1):
    checks[f"D11第{i}孔分布圆170贯穿10"]=len(cyl(5.5,x,z,36,46))==1
report={"passed":all(checks.values()),"checks":checks,"faces":faces,"quality":quality,"health":health,
        "hole_dimensions":hole_dims,"profile_dimensions":profile_dims,"wizard":wizard,
        "finished":datetime.now().astimezone().isoformat(),"limits":"原生几何练习；未建工程图、未赋完整制造标注、STEP未重导入。孔向导自动截面内部Fix保留，通用质量原结果为review_required，已单列复核。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
save_preview(m,str(out/"调节盘_等轴测.bmp"),"isometric")
save_preview(m,str(out/"调节盘_侧视.bmp"),"front")
m.ShowNamedView2("*Top",5);m.ViewZoomtofit2();m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"调节盘_球面端.bmp"),1600,1000)
m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(json.dumps({"passed":report["passed"],"count":len(checks),"failed":[k for k,v in checks.items() if not v]},ensure_ascii=False))
assert report["passed"]
