"""按原卷分段孔、观察侧及留空范围检查；不把特征成功等同于符合原图。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import win32com.client
from PIL import Image,ImageDraw

sys.path.insert(0,str(Path(__file__).resolve().parents[3]/"practice"))
from shared.cad import member as g
from shared.single_part import surfaces,dimensions,save
from sw_review import collect_model_summary,save_preview
from sw_sketch_quality import inspect_model_sketches

out=Path(sys.argv[1]).resolve();pid=int(sys.argv[2]);target=out/"阀上盖_SDF-04.SLDPRT"
sw=win32com.client.Dispatch("SldWorks.Application.32");m=g(sw,"ActiveDoc")
assert g(sw,"GetProcessID")==pid and Path(g(m,"GetPathName")).resolve()==target
assert target.with_suffix(".step").stat().st_size>0
save(m,target);sw.CloseDoc(g(m,"GetTitle"))
err=win32com.client.VARIANT(16387,0);warn=win32com.client.VARIANT(16387,0)
m=sw.OpenDoc6(str(target),1,1,"",err,warn)
assert m is not None and err.value==0 and m.ForceRebuild3(False)
health=collect_model_summary(m);quality=inspect_model_sketches(m);rows=surfaces(m)
def near(a,b):return abs(a-b)<1e-7
def cyl(r):return [f for f in rows if "Cylinder" in f and near(f["Cylinder"][6],r)]
def yrange(row,a,b):return near(row["box"][1],a) and near(row["box"][4],b)
slot=cyl(.004);ends=cyl(.0025);spring=cyl(.003);tap=cyl(.0034)
axis_angle=math.radians(22.6);n=[-math.cos(axis_angle),math.sin(axis_angle),0]
p0=[-.0284,.018,0]
assert len(tap)==1
delta=[tap[0]["Cylinder"][i]-p0[i] for i in range(3)]
dot=sum(delta[i]*n[i] for i in range(3));axis_error=math.sqrt(sum((delta[i]-dot*n[i])**2 for i in range(3)))
blind_data=g(m.FeatureByName("GB钢球弹簧母孔"),"GetDefinition")
q=quality["sketches"];builtin={"弹簧孔原生轮廓"}
arc_area=[]
for body in g(m,"GetBodies2",0,False):
    for face in g(body,"GetFaces"):
        su=g(face,"GetSurface")
        if g(su,"IsCylinder") and near(g(su,"CylinderParams")[6],.0165):arc_area.append(g(face,"GetArea"))
profile_dims=[d["value"] for d in dimensions(m,"轴向截面")]
arc_dims=[d["value"] for d in dimensions(m,"90度弧槽轮廓")]
cones=[r for r in rows if "Cone" in r]
entry_distance=(.0284-.0165)/(math.cos(axis_angle)+5*math.sin(axis_angle)/12)
entry=[p0[i]-entry_distance*n[i] for i in range(3)]
apex=[entry[i]-(.007+.0034/math.tan(math.radians(59)))*n[i] for i in range(3)]
thread=m.FeatureByName("M8装饰螺纹")
thread_data=g(thread,"GetDefinition") if thread is not None else None
checks={
    "保存重开及重建":err.value==0,
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "特征健康完整且无错误警告":not health.get("feature_error") and not health["faulty_features"] and not health["warning_features"] and not health["unknown_feature_health"],
    "六张草图完全定义且被消费":len(q)==6 and all(s["constrained_status"]==3 and not s["dangling"] and s["consumed_by"] for s in q),
    "无手工Fix且仅孔向导内置Fix":all(s["fix_count"]==(1 if s["name"] in builtin else 0) for s in q),
    "外圆Φ40直段至9点6":len(cyl(.020))==1 and yrange(cyl(.020)[0],0,9.6),
    "外锥段9点6至18及末端Φ33":any(yrange(c,9.6,18) and near(c["Cone"][6],.020) and near(c["Cone"][7],math.atan(5/12)) for c in cones),
    "杯口Φ34深2":len(cyl(.017))==1 and yrange(cyl(.017)[0],0,2),
    "外端Φ29浅凹0点5":len(cyl(.0145))==1 and yrange(cyl(.0145)[0],17.5,18),
    "中央Φ7配合段6点5至9":len(cyl(.0035))==1 and yrange(cyl(.0035)[0],6.5,9),
    "中央Φ5点2连接段9至11":len(cyl(.0026))==1 and yrange(cyl(.0026)[0],9,11),
    "图示C1为连接段11至12锥口":any(yrange(c,11,12) and near(c["Cone"][7],math.pi/4) and near(c["Cone"][6],.0036) for c in cones),
    "按B-B修正的Φ9沉孔12至17点5":len(cyl(.0045))==1 and yrange(cyl(.0045)[0],12,17.5),
    "长圆槽宽8长17深4点5":len(slot)==2 and all(yrange(c,2,6.5) and near(c["box"][2],-4) and near(c["box"][5],4) for c in slot) and sorted(round(c["Cylinder"][0]*1000,5) for c in slot)==[-4.5,4.5],
    "弧槽两端R2点5及中心圆Φ28":len(ends)==2 and all(near(math.hypot(c["Cylinder"][0],c["Cylinder"][2]),.014) and yrange(c,2,5.5) for c in ends),
    "弧槽位于AA正侧且关于AA对称":len(ends)==2 and all(c["Cylinder"][0]>0 for c in ends) and near(ends[0]["Cylinder"][0],ends[1]["Cylinder"][0]) and near(ends[0]["Cylinder"][2],-ends[1]["Cylinder"][2]),
    "弧槽外内R16点5和R11点5":len(cyl(.0165))==1 and len(cyl(.0115))==1,
    "弧槽90度实际圆柱面积及角度驱动":len(arc_area)==1 and abs(arc_area[0]-.0165*.0035*math.pi/2)<1e-9 and any(near(v,math.pi/2) for v in arc_dims),
    "两Φ6弹簧孔位于BB两侧Φ28分布圆":len(spring)==2 and all(near(c["Cylinder"][0],0) for c in spring) and sorted(round(c["Cylinder"][2]*1000,4) for c in spring)==[-14,14],
    "两弹簧孔从杯底2至基准深13且平底":len(spring)==2 and all(yrange(c,2,13) for c in spring) and near(g(blind_data,"DrillAngle"),math.pi),
    "M8轴线通过原图28点4定位点":axis_error<1e-8,
    "M8轴线22点6度且位于AA负侧":math.dist(tap[0]["Cylinder"][3:6],n)<1e-7 and tap[0]["Cylinder"][0]<0 and near(tap[0]["Cylinder"][2],0),
    "M8圆柱深7及118度钻尖":any(math.dist(c["Cone"][:3],apex)<1e-7 and near(c["Cone"][6],0) and near(c["Cone"][7],math.radians(59)) for c in cones),
    "M8自定义原生装饰螺纹已持久化":thread is not None and g(thread,"GetTypeName2")=="CosmeticThread" and near(g(thread_data,"Diameter"),.008) and near(g(thread_data,"BlindDepth"),.007),
    "斜孔角度及28点4为驱动值":any(near(v,axis_angle) for v in profile_dims) and any(near(v,.0284) for v in profile_dims),
    "Φ6孔使用GB库":g(blind_data,"Standard2")==13 and g(blind_data,"FastenerSize")=="Ø6",
}
interface={"内芯Φ7末端长2_对应孔段长2点5_轴向余量_mm":.5,
    "内芯宽8扁位_对应8H9槽_公称宽差_mm":0,
    "M5杆与Φ5点2孔_直径间隙_mm":.2,
    "M5头Φ8点5与Φ9沉孔_直径间隙_mm":.5,
    "M5头高5_沉孔座12至浅凹面17点5_余量_mm":.5,
    "status":"原图接口尺寸比对，未进行真实Mate试装或干涉分析"}
images=[]
for view in ["bottom","top","front"]:
    images.append((view,save_preview(m,str(out/f"review_{view}.bmp"),view)))
m.ShowNamedView2("*Bottom",6);g(m,"ActiveView").RotateAboutCenter(.45,.45);m.ViewZoomtofit2();m.GraphicsRedraw2()
iso=out/"review_cup-isometric.bmp";assert m.SaveBMP(str(iso),1600,1000);images.insert(0,("cup-isometric",str(iso)))
grid=Image.new("RGB",(1600,1050),"white");draw=ImageDraw.Draw(grid)
for i,(view,path) in enumerate(images):
    with Image.open(path) as im:grid.paste(im.convert("RGB").resize((800,500)),((i%2)*800,(i//2)*525+25))
    draw.text(((i%2)*800+10,(i//2)*525+5),view,fill="black")
grid.save(out/"review.png")
with Image.open(iso) as im:im.save(out/"isometric.png")
report={"target":target.name,"process_id":pid,"checked":datetime.now().astimezone().isoformat(),
    "passed":all(checks.values()),"checks":checks,"quality":quality,"health":health,"surfaces":rows,
    "interfaces":interface,"m8_axis_offset_m":axis_error,
    "status":"awaiting_visual_review" if all(checks.values()) else "failed",
    "limits":"未实际装配、出工程图或重导入STEP；M8为自定义装饰螺纹，Standard=-2，未与GB孔向导库关联；钻尖118°为普通钻尖表达假设，原图未单独标角度；材料按零件图45记录，首页LY12冲突待整卷核实。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"checks":len(checks),"failed":[k for k,v in checks.items() if not v],"status":report["status"]},ensure_ascii=False))
