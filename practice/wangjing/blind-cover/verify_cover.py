"""MCP建模后补孔草图设计约束，独立核验原图所要求的材料范围。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from sw_connect import get_com_member as g
from sw_sketch_quality import fully_define_sketch, inspect_model_sketches
from sw_review import collect_model_summary, save_preview
import win32com.client

out=Path(sys.argv[1]).resolve()
dest=out/"闷盖_LJT04.32.SLDPRT"
sw=win32com.client.Dispatch("SldWorks.Application.32")
assert int(g(sw,"GetProcessID"))==34092
m=g(sw,"ActiveDoc")
assert Path(g(m,"GetPathName")).resolve()==dest
for i in ([] if "--read-only" in sys.argv else [1]):
    hole=m.FeatureByName(f"安装孔{i}")
    f=g(hole,"GetFirstSubFeature")
    while f is not None and g(f,"GetTypeName2")!="ProfileFeature":
        f=g(f,"GetNextSubFeature")
    assert f is not None
    m.ClearSelection2(True)
    assert f.Select2(False,0)
    m.EditSketch()
    try:
        fully_define_sketch(m)
    finally:
        m.SketchManager.InsertSketch(True)
    f.Name=f"孔位{i}"
assert m.ForceRebuild3(False)
quality=inspect_model_sketches(m)
health=collect_model_summary(m)
faces=[]
for body in g(m,"GetBodies2",0,False):
    for face in g(body,"GetFaces"):
        s=g(face,"GetSurface")
        row={"box_m":list(g(face,"GetBox")),"area_m2":g(face,"GetArea")}
        for name,test,params in [("cylinder","IsCylinder","CylinderParams"),
                                  ("cone","IsCone","ConeParams"),
                                  ("torus","IsTorus","TorusParams"),
                                  ("plane","IsPlane","PlaneParams")]:
            if g(s,test):
                row[name]=list(g(s,params))
        faces.append(row)

def near(a,b):return abs(a-b)<1e-5
def cyl(radius,x,z,lo,hi):
    return [f for f in faces if "cylinder" in f and
            near(f["cylinder"][6]*1000,radius) and
            near(f["cylinder"][0]*1000,x) and near(f["cylinder"][2]*1000,z) and
            near(f["box_m"][1]*1000,lo) and near(f["box_m"][4]*1000,hi)]
def plane(y,radius):
    return any("plane" in f and near(f["box_m"][1]*1000,y) and
               near(f["box_m"][4]*1000,y) and
               near(f["area_m2"]*1e6,math.pi*radius**2) for f in faces)
torus=[f for f in faces if "torus" in f]
cone=[f for f in faces if "cone" in f]
checks={
    "单实体":len(g(m,"GetBodies2",0,False))==1,
    "两草图全约束无Fix":quality["status"]=="pass" and len(quality["sketches"])==2,
    "无特征错误警告":not health["faulty_features"] and not health["warning_features"],
    "法兰D100厚12位置16至28":len(cyl(50,0,0,16,28))==1,
    "定位圆柱D62含端C1_5":len(cyl(31,0,0,1.5,16))==1,
    "两处R2回转过渡":len(torus)==2 and all(near(f["torus"][7]*1000,2) for f in torus),
    "锥腔1比10":any(near(math.tan(f["cone"][7]),.05) for f in cone),
    "仅锥腔和C1_5两个锥面":len(cone)==2,
    "外端C1_5位置0至1_5":any(near(f["box_m"][1]*1000,0) and near(f["box_m"][4]*1000,1.5)
                                   and near(math.tan(f["cone"][7]),1) for f in cone),
    "后盲腔底Y16及锥圆角切点半径":plane(16,26-.05*14-2*math.sqrt(1+.05**2)),
    "前浅凹底Y26半径29":plane(26,29),
    "轴心隔壁10毫米":plane(26,29) and plane(16,26-.05*14-2*math.sqrt(1+.05**2)),
    "浅凹R2从Y26至28且口D62":any(near(f["box_m"][1]*1000,26) and
        near(f["box_m"][4]*1000,28) and near(f["torus"][6]*1000,29) for f in torus),
}
for x,z in [(41,0),(0,41),(-41,0),(0,-41)]:
    checks[f"D9孔中心{x}_{z}贯穿法兰"]=len(cyl(4.5,x,z,16,28))==1
report={"checks":checks,"passed":all(checks.values()),"faces":faces,
        "sketches":quality,"health":health,"finished":datetime.now().astimezone().isoformat(),
        "limits":"材料、公差与粗糙度仅题面属性；未出工程图、装配或STEP重导入。"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
for filename,view in [("闷盖_前侧.bmp","isometric"),("闷盖_前端.bmp","top"),
                       ("闷盖_侧面.bmp","front")]:
    save_preview(m,str(out/filename),view)
print(json.dumps({"passed":report["passed"],"checks":checks,
                  "torus":[f["torus"] for f in torus]},ensure_ascii=False))
assert report["passed"]
