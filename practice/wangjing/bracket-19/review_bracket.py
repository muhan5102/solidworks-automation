"""以题面尺寸、空区和近侧壁范围核验本轮模型；STEP已由MCP导出。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client
from PIL import Image,ImageDraw

sys.path.insert(0,str(Path(__file__).resolve().parents[3]/"practice"))
from shared.cad import member as g
from shared.single_part import surfaces,dimensions,save
from sw_review import collect_model_summary,save_preview
from sw_sketch_quality import inspect_model_sketches
from sw_connect import create_empty_dispatch_variant
from fillet_check import check_fillet_outline

out=Path(sys.argv[1]).resolve();target=out/(sys.argv[2] if len(sys.argv)>2 else "踏架_LJT05.19.SLDPRT")
sw=win32com.client.Dispatch("SldWorks.Application.32");m=g(sw,"ActiveDoc")
assert g(sw,"GetProcessID")==49384 and Path(g(m,"GetPathName")).resolve()==target
assert target.with_suffix(".step").stat().st_size>0
save(m,target);sw.CloseDoc(g(m,"GetTitle"))
e=win32com.client.VARIANT(16387,0);w=win32com.client.VARIANT(16387,0)
m=sw.OpenDoc6(str(target),1,1,"",e,w)
assert m is not None and e.value==0 and m.ForceRebuild3(False)
edge_counts={}
for name in ["图示根部R3","翼缘边缘R2"]:
    fd=g(m.FeatureByName(name),"GetDefinition")
    assert fd.AccessSelections(m,create_empty_dispatch_variant())
    edge_counts[name]=len(g(fd,"Edges"))
    fd.ReleaseSelectionAccess()
quality=inspect_model_sketches(m);health=collect_model_summary(m);rows=surfaces(m)
tol=1e-7
def close(a,b):return abs(a-b)<tol
def cylinders(radius):return [r for r in rows if "Cylinder" in r and close(r["Cylinder"][6],radius)]
def one_radius(radius,center):
    return any(math.dist(r["Cylinder"][:2],center)<tol for r in cylinders(radius))
q=quality["sketches"]
builtin={"GB沉头母孔原生轮廓","GB_M10螺纹母孔原生轮廓"}
seat=cylinders(.0035);tap=cylinders(.00425);main=cylinders(.012);boss=cylinders(.010)
rib=cylinders(.030)
d=g(m.FeatureByName("GB沉头母孔"),"GetDefinition")
t=g(m.FeatureByName("GB_M10螺纹母孔"),"GetDefinition")
cone=[r for r in rows if "Cone" in r]
checks={
    "重开无错误并重建":e.value==0,
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "特征健康完整无警告":not health.get("feature_error") and not health["faulty_features"] and not health["warning_features"] and not health["unknown_feature_health"],
    "11张草图完全定义且被消费":len(q)==11 and all(r["constrained_status"]==3 and not r["dangling"] and r["consumed_by"] for r in q),
    "无手工Fix且仅两个孔向导轮廓内置Fix":all(r["fix_count"]==(1 if r["name"] in builtin else 0) for r in q),
    "主孔Φ24轴向58且中心正确":len(main)==1 and math.dist(main[0]["Cylinder"][:2],[0,0])<tol and close(main[0]["box"][2],-29) and close(main[0]["box"][5],29),
    "外轴套Φ38长58":any(close(r["box"][2],-29) and close(r["box"][5],29) for r in cylinders(.019)),
    "弯臂内R9圆心28和24":one_radius(.009,[.028,.024]),
    "腹板R30圆心28和24及厚10":len(rib)==1 and math.dist(rib[0]["Cylinder"][:2],[.028,.024])<tol and close(rib[0]["box"][2],-5) and close(rib[0]["box"][5],5),
    "安装凸台两端R10轴向总长45":len(boss)==2 and close(min(r["box"][2] for r in boss),-22.5) and close(max(r["box"][5] for r in boss),22.5),
    "安装面30度和28为原图驱动":all(close(a,b) for a,b in zip([r["value"] for r in dimensions(m,"轴套与斜面定位")],[.038,.028,math.pi/6])),
    "踏板34乘50厚7外范围":any("Plane" in r and close(r["box"][1],30) and close(r["box"][4],30) and close(r["box"][0],59) and close(r["box"][3],93) and close(r["box"][2],-25) and close(r["box"][5],25) for r in rows),
    "踏板四角R5":len(cylinders(.005))==4,
    "两Φ7孔X76孔距32":len(seat)==2 and all(close(r["Cylinder"][0],.076) for r in seat) and all(close(a,b) for a,b in zip(sorted(r["Cylinder"][2] for r in seat),[-.016,.016])),
    "沉头孔穿透厚7且无远端沉头":all(close(r["box"][1],30) and close(r["box"][4],34) for r in seat) and len([r for r in cone if r["box"][0]>60])==2,
    "GB沉头Φ13和90度及Φ7":g(d,"Standard2")==13 and close(g(d,"ThruHoleDiameter"),.007) and close(g(d,"CounterSinkDiameter"),.013) and close(g(d,"CounterSinkAngle"),math.pi/2),
    "两M10位置25间距及30度径向":len(tap)==2 and all(math.dist(r["Cylinder"][3:6],[-.5,math.cos(math.pi/6),0])<tol for r in tap) and all(close(a,b) for a,b in zip(sorted(r["Cylinder"][2] for r in tap),[-.0125,.0125])),
    "M10孔仅近壁远壁完整":len(tap)==2 and all(r["box"][1]>0 for r in tap),
    "GB_M10_7H及底孔8点5":g(t,"Standard2")==13 and g(t,"FastenerSize")=="M10" and g(t,"ThreadClass")=="7H" and close(g(t,"TapDrillDiameter"),.0085),
}
checks.update(check_fillet_outline(m))
images=[]
for view in ["isometric","front","top","bottom"]:
    path=save_preview(m,str(out/f"review_{view}.bmp"),view);images.append((view,path))
grid=Image.new("RGB",(1600,1050),"white");draw=ImageDraw.Draw(grid)
for i,(view,path) in enumerate(images):
    with Image.open(path) as im:grid.paste(im.convert("RGB").resize((800,500)),((i%2)*800,(i//2)*525+25))
    draw.text(((i%2)*800+10,(i//2)*525+5),view,fill="black")
grid.save(out/"review.png")
with Image.open(images[0][1]) as im:im.save(out/"isometric.png")
m.ShowNamedView2("*Isometric",7);m.GraphicsRedraw2()
report={"target":target.name,"process_id":49384,"checked":datetime.now().astimezone().isoformat(),
    "passed":all(checks.values()),"checks":checks,"quality":quality,"health":health,"surfaces":rows,"fillet_seed_counts":edge_counts,
    "status":"awaiting_visual_review" if all(checks.values()) else "failed",
    "limits":"螺纹为孔向导装饰表达；STEP未重导入；材料与工艺作为属性，未作实体材料/制造验收；未出工程图。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"checks":len(checks),"failed":[k for k,v in checks.items() if not v],"status":report["status"]},ensure_ascii=False))
