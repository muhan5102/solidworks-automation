"""偏心盘：回读原图所需的偏心、轴向范围、孔位、键槽和后端圆弧。"""
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
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==34092
m=g(s,"ActiveDoc")
assert Path(g(m,"GetPathName")).resolve()==out/"偏心盘_LJT04.26.SLDPRT"
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m);health=collect_model_summary(m)
faces=[]
for b in g(m,"GetBodies2",0,False):
    for f in g(b,"GetFaces"):
        surf=g(f,"GetSurface")
        row={"box_mm":[v*1000 for v in g(f,"GetBox")],"area_mm2":g(f,"GetArea")*1e6}
        for name,test,params in [("cylinder","IsCylinder","CylinderParams"),
                                ("cone","IsCone","ConeParams"),("plane","IsPlane","PlaneParams")]:
            if g(surf,test):row[name]=list(g(surf,params))
        faces.append(row)
def near(a,b):return abs(a-b)<1e-5
def cyl(r,x,z,lo,hi):
    return [f for f in faces if "cylinder" in f and near(f["cylinder"][6]*1000,r)
            and near(f["cylinder"][0]*1000,x) and near(f["cylinder"][2]*1000,z)
            and near(f["box_mm"][1],lo) and near(f["box_mm"][4],hi)]
def flat(coord,value,lo,hi):
    return [f for f in faces if "plane" in f and near(f["box_mm"][coord],value)
            and near(f["box_mm"][coord+3],value) and near(f["box_mm"][1],lo)
            and near(f["box_mm"][4],hi)]
cones=[f for f in faces if "cone" in f]
end_faces=[f for f in faces if "plane" in f and near(f["box_mm"][1],76) and near(f["box_mm"][4],76)]
def chord_integral(radius,half_width):
    return half_width*math.sqrt(radius**2-half_width**2)+radius**2*math.asin(half_width/radius)
# 从宽28、两条R76弧、D36孔、键槽和D6.8孔独立算端面材料面积。
expected_end_area=(2*chord_integral(76,14)-74*28
                   -2*chord_integral(18,14)
                   -(10*21.3-chord_integral(18,5))-math.pi*3.4**2)
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "八草图完全定义无Fix无孤立":quality["status"]=="pass" and len(quality["sketches"])==8,
    "无特征错误警告":not health["faulty_features"] and not health["warning_features"],
    "前盘D108偏心向上10且Y10至24":len(cyl(54,0,10,10,24))==1,
    "后盘D92偏心向上10且Y56至70":len(cyl(46,0,10,56,70))==1,
    "轴套D60前段Y0至10":len(cyl(30,0,0,0,10))==1,
    "轴套D60中段Y24至56":len(cyl(30,0,0,24,56))==1,
    "D36孔除前倒角外贯通至Y76":len(cyl(18,0,0,2,76))==1,
    "前D12孔高50且仅切前盘":len(cyl(6,0,50,10,24))==1,
    "后M8底孔高40且Y56至76":len(cyl(3.4,0,40,56,76))==1,
    "两端孔中心相差10未复制错位":len(cyl(6,0,50,10,24))==1 and len(cyl(3.4,0,40,56,76))==1,
    "后凸台宽28左侧X负14":bool(flat(0,-14,70,76)),
    "后凸台宽28右侧X正14":bool(flat(0,14,70,76)),
    "上弧R76中心Z负28最高48":len(cyl(76,0,-28,70,76))==1 and
        near(cyl(76,0,-28,70,76)[0]["box_mm"][5],48),
    "下弧R76中心Z46最低负30":bool(cyl(76,0,46,70,76)) and
        near(min(f["box_mm"][2] for f in cyl(76,0,46,70,76)),-30),
    "键槽左壁X负5贯通":bool(flat(0,-5,0,76)),
    "键槽右壁X正5贯通":bool(flat(0,5,0,76)),
    "键槽顶Z21_3贯通且距孔底39_3":bool(flat(2,21.3,0,76)) and near(21.3-(-18),39.3),
    "仅前孔口C2且轴向0至2":len(cones)==1 and near(cones[0]["box_mm"][1],0)
        and near(cones[0]["box_mm"][4],2) and near(math.tan(cones[0]["cone"][7]),1),
    "原生M8装饰螺纹":m.FeatureByName("M8内螺纹") is not None,
    "右端面仅有宽28范围内的材料":bool(end_faces) and all(
        f["box_mm"][0]>=-14-1e-5 and f["box_mm"][3]<=14+1e-5 for f in end_faces),
    "右端是两块凸台端面而非整圈":len(end_faces)==2,
    "右端材料面积与原图截面独立计算一致":near(
        sum(f["area_mm2"] for f in end_faces),expected_end_area),
    "D60轴套未延伸至最后6毫米":not cyl(30,0,0,70,76),
}
report={"checks":checks,"passed":all(checks.values()),"faces":faces,"quality":quality,"health":health,
        "finished":datetime.now().astimezone().isoformat(),
        "right_end_area_expected_mm2":expected_end_area,
        "limits":"右端整圈多余材料由用户指出后修正，非独立识图通过。M8为D6.8底孔+装饰螺纹，无实体牙型；公差及材料按属性表达；未出工程图；未重导入STEP。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
save_preview(m,str(out/"偏心盘_等轴测.bmp"),"isometric")
save_preview(m,str(out/"偏心盘_轴向侧视.bmp"),"front")
m.ShowNamedView2("*Bottom",6);m.ViewZoomtofit2();m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"偏心盘_前端.bmp"),1600,1000)
m.ShowNamedView2("*Top",5);m.ViewZoomtofit2();m.GraphicsRedraw2()
assert m.SaveBMP(str(out/"偏心盘_后端.bmp"),1600,1000)
m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(json.dumps({"passed":report["passed"],"checks":checks},ensure_ascii=False))
assert report["passed"]
