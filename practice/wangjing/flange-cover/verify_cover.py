"""由题图独立核对各段材料、端面面积、孔群、径向孔和草图驱动。"""
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
assert g(s,"GetProcessID")==json.loads((out/"session.json").read_text(encoding="utf-8"))["process_id"]
m=g(s,"ActiveDoc");target=out/(sys.argv[2] if len(sys.argv)>2 else "法兰盘_LJT04.18.SLDPRT")
assert Path(g(m,"GetPathName")).resolve()==target
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m);health=collect_model_summary(m)
faces=[]
radial_neighbors=[]
for body in g(m,"GetBodies2",0,False):
    for face in g(body,"GetFaces"):
        surf=g(face,"GetSurface")
        row={"box_mm":[v*1000 for v in g(face,"GetBox")],"area_mm2":g(face,"GetArea")*1e6}
        for name,test,params in [("cylinder","IsCylinder","CylinderParams"),("cone","IsCone","ConeParams"),
                                 ("torus","IsTorus","TorusParams"),("plane","IsPlane","PlaneParams")]:
            if g(surf,test):row[name]=list(g(surf,params))
        if "cylinder" in row and abs(row["cylinder"][6]-.0025)<1e-8 and abs(row["cylinder"][1]-.0285)<1e-8:
            for edge in g(face,"GetEdges"):
                for neighbor in g(edge,"GetTwoAdjacentFaces2"):
                    ns=g(neighbor,"GetSurface")
                    if g(ns,"IsCylinder"):
                        radial_neighbors.append(list(g(ns,"CylinderParams")))
        faces.append(row)
def near(a,b):return abs(a-b)<1e-5
def cyl(r,x,z,lo,hi):
    return [f for f in faces if "cylinder" in f and near(f["cylinder"][6]*1000,r)
        and near(f["cylinder"][0]*1000,x) and near(f["cylinder"][2]*1000,z)
        and near(f["box_mm"][1],lo) and near(f["box_mm"][4],hi)]
def area(y):
    return sum(f["area_mm2"] for f in faces if "plane" in f and near(f["box_mm"][1],y)
               and near(f["box_mm"][4],y))
def strip_area(r,a):return 2*(a*math.sqrt(r*r-a*a)+r*r*math.asin(a/r))
def dimensions(name):
    feature=m.FeatureByName(name);dd=g(feature,"GetFirstDisplayDimension");result=[]
    while dd is not None:
        d=g(dd,"GetDimension2",0)
        result.append({"name":g(d,"FullName"),"value":g(d,"SystemValue"),"driven_state":g(d,"DrivenState")})
        dd=g(feature,"GetNextDisplayDimension",dd)
    return result
hole_dims=dimensions("安装孔位")
cones=[f for f in faces if "cone" in f];tori=[f for f in faces if "torus" in f]
radial=[f for f in faces if "cylinder" in f and near(f["cylinder"][6]*1000,2.5)
        and near(f["cylinder"][0]*1000,0) and near(f["cylinder"][1]*1000,28.5)
        and near(abs(f["cylinder"][5]),1)]
thread=m.FeatureByName("M6装饰螺纹")
thread_data=None
if thread is not None:
    data=g(thread,"GetDefinition")
    thread_data={"size":g(data,"Size"),"diameter_m":g(data,"Diameter"),
                 "depth_m":g(data,"BlindDepth"),"callout":g(data,"ThreadCallout")}
wizard=m.FeatureByName("安装沉孔向导")
wizard_data=None
if wizard is not None:
    data=g(wizard,"GetDefinition")
    wizard_data={k:g(data,k) for k in ["ThruHoleDiameter","CounterBoreDiameter","CounterBoreDepth",
        "EndCondition","NearCounterSinkDiameter","MidCounterSinkDiameter","FarCounterSinkDiameter"]}
    generated=[sk for sk in quality["sketches"] if sk["name"]=="沉孔向导轮廓"]
    user_sketches=[sk for sk in quality["sketches"] if sk["name"]!="沉孔向导轮廓"]
    # 孔向导生成的截面含内部原点固定关系，保留通用审查的review_required原结果，
    # 本案例单列复核其原生参数、四个轮廓驱动尺寸及实际几何，不把它当手工草图Fix。
    sketch_ok=all(sk["status"]=="pass" for sk in user_sketches) and len(generated)==1 \
        and generated[0]["constrained_status"]==3 and generated[0]["dimension_count"]==4 \
        and generated[0]["fix_count"]==1 and not generated[0]["dangling"]
else:
    sketch_ok=quality["status"]=="pass"
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "手工草图完全定义无Fix且原生孔截面已单列复核":sketch_ok,
    "无特征错误警告":not health["faulty_features"] and not health["warning_features"],
    "左定位D70仅Y0至3":len(cyl(35,0,0,0,3))==1,
    "法兰D130外弧仅Y3至13":len(cyl(65,0,0,3,13))==2,
    "右轴套D55在Y18至40":len(cyl(27.5,0,0,18,40))==1,
    "D42孔两端C2间Y2至40":len(cyl(21,0,0,2,40))==1,
    "退刀槽底D53位于Y15至17":len(cyl(26.5,0,0,15,17))==1,
    "退刀槽R1位置Y14至15":len(tori)==1 and near(tori[0]["torus"][6]*1000,27.5)
        and near(tori[0]["torus"][7]*1000,1) and near(tori[0]["box_mm"][1],14)
        and near(tori[0]["box_mm"][4],15),
    "退刀槽两个45度斜面位置正确":sum(near(math.tan(f["cone"][7]),1)
        and ((near(f["box_mm"][1],14) and near(f["box_mm"][4],15))
             or (near(f["box_mm"][1],17) and near(f["box_mm"][4],18))) for f in cones)==2,
    "左端环形材料面积正确":near(area(0),math.pi*(35**2-23**2)),
    "右端环形材料面积正确无多余材料":near(area(42),math.pi*(25.5**2-23**2)),
    "法兰前端面积含双削平及六孔":near(area(3),strip_area(65,50)-math.pi*35**2-4*math.pi*3.5**2-2*math.pi*2.5**2),
    "法兰后端面积含双削平退刀及沉孔":near(area(15),strip_area(63,50)-math.pi*30.5**2-4*math.pi*6**2-2*math.pi*2.5**2),
    "两侧削平距轴各50且限法兰厚12":sum("plane" in f and near(abs(f["box_mm"][0]),50)
        and near(f["box_mm"][0],f["box_mm"][3]) and near(f["box_mm"][1],3)
        and near(f["box_mm"][4],15) for f in faces)==2,
    "孔群具有原生45度驱动":any(near(d["value"],math.pi/4) and d["driven_state"]==2 for d in hole_dims),
    "孔群具有42_5半径驱动":any(near(d["value"],.0425) and d["driven_state"]==2 for d in hole_dims),
    "仅一个四组孔阵列":sum(f["type"]=="CirPattern" for f in health["features"])==1,
    "左内口C2":any(near(f["box_mm"][1],0) and near(f["box_mm"][4],2) for f in cones),
    "右端内外两处C2":sum(near(f["box_mm"][1],40) and near(f["box_mm"][4],42) for f in cones)==2,
    "法兰后侧外弧C2":sum(near(f["box_mm"][1],13) and near(f["box_mm"][4],15) for f in cones)==2,
    "径向D5底孔仅一处且轴向位置28_5":len(radial)==1,
    "径向孔只切正Z侧上壁":len(radial)==1 and near(radial[0]["box_mm"][5],27.5)
        and radial[0]["box_mm"][2]>0
        and all(any(near(p[6],r) and near(abs(p[4]),1) for p in radial_neighbors) for r in [.021,.0275]),
    "M6原生装饰螺纹保留":thread is not None and g(thread,"GetTypeName2")=="CosmeticThread"
        and near(thread_data["diameter_m"],.006) and near(thread_data["depth_m"],.0065)
        and "M6" in thread_data["callout"],
}
if wizard_data is not None:
    checks["原生孔向导D7_D12深5贯穿无附加锥沉孔"]=near(wizard_data["ThruHoleDiameter"],.007) \
        and near(wizard_data["CounterBoreDiameter"],.012) and near(wizard_data["CounterBoreDepth"],.005) \
        and wizard_data["EndCondition"]==1 \
        and all(near(wizard_data[k],0) for k in ["NearCounterSinkDiameter","MidCounterSinkDiameter","FarCounterSinkDiameter"])
for i,(x,z) in enumerate([(a*42.5/math.sqrt(2),b*42.5/math.sqrt(2)) for a,b in [(1,1),(-1,1),(-1,-1),(1,-1)]],1):
    checks[f"第{i}D7通孔中心及Y3至10"]=len(cyl(3.5,x,z,3,10))==1
    checks[f"第{i}D12沉孔入口Y15深5"]=len(cyl(6,x,z,10,15))==1
for z in [-57,57]:
    checks[f"D5定位孔Z{z}贯穿法兰"]=len(cyl(2.5,0,z,3,15))==1
report={"passed":all(checks.values()),"checks":checks,"faces":faces,"quality":quality,"health":health,
        "hole_dimensions":hole_dims,"thread":thread_data,"finished":datetime.now().astimezone().isoformat(),
        "radial_adjacent_cylinders":radial_neighbors,
        "wizard":wizard_data,
        "limits":"M6为D5底孔加自定义原生装饰螺纹，无实体牙型；未生成工程图，未重导入STEP。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
save_preview(m,str(out/"法兰盘_等轴测.bmp"),"isometric")
save_preview(m,str(out/"法兰盘_侧视.bmp"),"front")
m.ShowNamedView2("*Bottom",6);m.ViewZoomtofit2();m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"法兰盘_前端.bmp"),1600,1000)
m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(json.dumps({"passed":report["passed"],"checks":checks},ensure_ascii=False))
assert report["passed"]
