"""LJT04.30核验：孔群台阶/锥度、键槽、径向单壁及应留空材料。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_sketch_quality import inspect_model_sketches
from sw_review import collect_model_summary,save_preview
out=Path(sys.argv[1]).resolve();s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
m=g(s,"ActiveDoc");assert Path(g(m,"GetPathName")).resolve()==out/"半联轴器_LJT04.30.SLDPRT"
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m);health=collect_model_summary(m);faces=[];radial_neighbors=[]
for body in g(m,"GetBodies2",0,False):
    for face in g(body,"GetFaces"):
        su=g(face,"GetSurface");row={"box":[v*1000 for v in g(face,"GetBox")],"area":g(face,"GetArea")*1e6}
        for kind in ["Cylinder","Cone","Torus","Plane"]:
            if g(su,"Is"+kind):row[kind]=list(g(su,kind+"Params"))
        if "Cylinder" in row and abs(row["Cylinder"][6]-.0051)<1e-8:
            for edge in g(face,"GetEdges"):
                for adj in g(edge,"GetTwoAdjacentFaces2"):
                    ns=g(adj,"GetSurface")
                    if g(ns,"IsCylinder"):radial_neighbors.append(list(g(ns,"CylinderParams")))
        faces.append(row)
def near(a,b):return abs(a-b)<1e-5
def cyl(r,x,z,y0,y1):
    return [f for f in faces if "Cylinder" in f and near(f["Cylinder"][6]*1000,r)
        and near(f["Cylinder"][0]*1000,x) and near(f["Cylinder"][2]*1000,z)
        and near(f["box"][1],y0) and near(f["box"][4],y1)]
def area(y):
    return sum(f["area"] for f in faces if "Plane" in f and near(f["box"][1],y) and near(f["box"][4],y))
def key_extra(r):
    return 12*24.3-(6*math.sqrt(r*r-36)+r*r*math.asin(6/r))
def dimensions(name):
    f=m.FeatureByName(name);dd=g(f,"GetFirstDisplayDimension");values=[]
    while dd is not None:
        d=g(dd,"GetDimension2",0);values.append({"name":g(d,"FullName"),"value":g(d,"SystemValue"),"driven":g(d,"DrivenState")})
        dd=g(f,"GetNextDisplayDimension",dd)
    return values
cones=[f for f in faces if "Cone" in f];tori=[f for f in faces if "Torus" in f]
radial=[f for f in faces if "Cylinder" in f and near(f["Cylinder"][6],.0051)]
thread=m.FeatureByName("M12装饰螺纹");td=g(thread,"GetDefinition") if thread is not None else None
thread_data={k:g(td,k) for k in ["Diameter","BlindDepth","ThreadCallout"]} if td is not None else None
key_dims=dimensions("键槽轮廓")
floor=[f for f in faces if "Plane" in f and near(f["box"][2],24.3) and near(f["box"][5],24.3)]
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "四草图完全定义无Fix无孤立":quality["status"]=="pass" and len(quality["sketches"])==4,
    "特征无错误警告":not health["faulty_features"] and not health["warning_features"],
    "法兰D200限Y40至78":len(cyl(100,0,0,40,78))==1,
    "轴套D100前段Y2至19_2":len(cyl(50,0,0,2,19.2))>=1,
    "轴套D100后段Y20_8至34":len(cyl(50,0,0,20.8,34))>=1,
    "环槽底D94且宽1_6":len(cyl(47,0,0,19.2,20.8))>=1,
    "主体内孔D42限Y2至78":len(cyl(21,0,0,2,78))>=1,
    "根部R4位置Y34至38":len(tori)==1 and near(tori[0]["Torus"][1],.034)
        and near(tori[0]["Torus"][6],.054) and near(tori[0]["Torus"][7],.004),
    "键槽顶Z24_3宽12贯穿80":len(floor)==1 and near(floor[0]["box"][0],-6)
        and near(floor[0]["box"][3],6) and near(floor[0]["box"][1],0)
        and near(floor[0]["box"][4],80) and near(floor[0]["area"],960),
    "键宽12及孔底到槽顶45_3为驱动":all(any(near(d["value"],v) and d["driven"]==2 for d in key_dims) for v in [.012,.0453]),
    "右端材料含八锥孔键槽及C2":near(area(80),math.pi*(98**2-23**2)-8*math.pi*9**2-key_extra(23)),
    "左端材料含键槽及内外C2":near(area(0),math.pi*(48**2-23**2)-key_extra(23)),
    "法兰后端材料含八沉孔及R4":near(area(38),math.pi*(98**2-54**2)-8*math.pi*19**2),
    "仅一个八实例阵列":sum(f["type"]=="CirPattern" for f in health["features"])==1,
    "径向D10_2仅一孔且位于Y20负Z":len(radial)==1 and near(radial[0]["Cylinder"][1],.020)
        and radial[0]["box"][5]<0,
    "径向孔邻接D42内壁及轴套外壁":all(any(near(p[6],r) for p in radial_neighbors) for r in [.021,.050]),
    "M12原生装饰螺纹保留":thread is not None and g(thread,"GetTypeName2")=="CosmeticThread"
        and near(thread_data["Diameter"],.012) and "M12-7H" in thread_data["ThreadCallout"],
}
for i in range(8):
    theta=i*math.pi/4;x,z=75*math.cos(theta),75*math.sin(theta)
    checks[f"第{i+1}孔D38沉孔Y38至54"]=len(cyl(19,x,z,38,54))==1
    checks[f"第{i+1}孔D26小台阶Y54至54_5"]=len(cyl(13,x,z,54,54.5))==1
    cs=[f for f in cones if near(f["Cone"][0]*1000,x) and near(f["Cone"][2]*1000,z)]
    checks[f"第{i+1}锥孔1比10右口D18"]=len(cs)==1 and near(math.tan(cs[0]["Cone"][7]),.05) \
        and near(cs[0]["box"][1],54.5) and near(cs[0]["box"][4],80) \
        and near(cs[0]["Cone"][6]*1000,7.725) and near(7.725+(80-54.5)*math.tan(cs[0]["Cone"][7]),9)
report={"passed":all(checks.values()),"checks":checks,"quality":quality,"health":health,"faces":faces,
    "key_dimensions":key_dims,"thread":thread_data,"radial_neighbors":radial_neighbors,
    "finished":datetime.now().astimezone().isoformat(),
    "limits":"出口两个1经用户提示按塞规检验位置理解，未建额外凸台或倒角；M12为底孔加原生装饰螺纹，无实体牙型；未出工程图，STEP未重导入。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
save_preview(m,str(out/"半联轴器_等轴测.bmp"),"isometric")
m.ShowNamedView2("*Bottom",6);m.ViewZoomtofit2();m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"半联轴器_轮毂端.bmp"),1600,1000)
m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(json.dumps({"passed":report["passed"],"count":len(checks),"failed":[k for k,v in checks.items() if not v]},ensure_ascii=False))
assert report["passed"]
