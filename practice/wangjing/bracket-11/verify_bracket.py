"""LJT05.11原图参数与实体/草图交叉核对，统一收尾。"""
from pathlib import Path
import sys,math,json
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
sys.path.insert(0,str(ROOT/"practice"))
from sw_connect import connect_solidworks,get_com_member as g
from shared.single_part import prepare_review,dimensions,finish_review
out=Path(sys.argv[1]).resolve()
optimized="r02" in out.name
target=out/("支架_LJT05.11_指令优化.SLDPRT" if optimized else "支架_LJT05.11.SLDPRT")
sw,m=connect_solidworks(2024,.1)
assert g(sw,"GetProcessID")==42548


def near(a,b):return abs(a-b)<1e-7


def check(m,rows):
    def cylinders(radius,axis,center):
        return [f for f in rows if "Cylinder" in f and near(f["Cylinder"][6],radius/1000)
                and abs(f["Cylinder"][3+axis])>.999999
                and all(near(f["Cylinder"][i],center[i]/1000) for i in range(3) if i!=axis)]
    def ranges(faces,axis,a,b):
        return any(abs(f["box"][axis]-a)<1e-4 and abs(f["box"][axis+3]-b)<1e-4 for f in faces)
    def depth(name):return g(g(m.FeatureByName(name),"GetDefinition"),"GetDepth",True)
    rib=m.FeatureByName("斜臂腹板")
    ribsk=g(m.FeatureByName("腹板边界"),"GetSpecificFeature2")
    seg=list(g(ribsk,"GetSketchSegments"))[0]
    p,q=g(seg,"GetStartPoint2"),g(seg,"GetEndPoint2")
    rd=g(rib,"GetDefinition")
    slot=list(g(g(m.FeatureByName("夹紧槽草图"),"GetSpecificFeature2"),"GetSketchSegments"))
    points=[g(e,"GetStartPoint2") for e in slot]
    basepoints=[g(e,"GetStartPoint2") for e in g(g(m.FeatureByName("底座截面"),"GetSpecificFeature2"),"GetSketchSegments")]
    basex=[g(p,"X") for p in basepoints]
    basey=[g(p,"Y") for p in basepoints]
    # 图中直接值：Φ35/Φ20×50，底座82/50/16，孔距40、2Φ15/Φ28深3；
    # 腹板8、翼板6×40；夹耳R13、孔轴距25、18高、槽3、Φ11/Φ18深3、M10。
    checks={
        "套筒外径35及内孔20沿Z长50":ranges(cylinders(17.5,2,(0,0,0)),2,-25,25)
            and ranges(cylinders(10,2,(0,0,0)),2,-25,25),
        "套筒轴向拉伸50及底座宽82":near(depth("套筒实心毛坯" if optimized else "夹紧套筒"),.050) and near(depth("底座"),.082),
        "翼板中面宽40":near(depth("斜臂翼板"),.040),
        "翼板实际两平行面间距6":False,
        "原生Rib及厚8":g(rib,"GetTypeName2")=="Rib" and near(g(rd,"Thickness"),.008),
        "腹板上起点图示左偏7且位于套筒外圆":near(g(p,"X"),-.007)
            and near(g(p,"Y"),-math.sqrt(.0175**2-.007**2)),
        "腹板下端X44距底面120为5":near(g(q,"X"),.044) and near(g(q,"Y"),-.115),
        "底座孔各自直径15从X60贯至44":all(ranges(cylinders(7.5,0,(0,-100,z)),0,44,57) for z in [-20,20]),
        "底座两沉孔直径28深3且孔距40":all(ranges(cylinders(14,0,(0,-100,z)),0,57,60) for z in [-20,20]),
        "底座背面相对套筒中心60":any("Plane" in f and abs(abs(f["Plane"][0])-1)<1e-7
            and abs(f["Plane"][3]-.060)<1e-7 for f in rows),
        "底座截面50高与顶层10、外伸24":near(max(basey)-min(basey),.050)
            and all(any(near(v,y) for y in basey) for v in [-.070,-.080]) and near(max(basex)-min(basex),.024),
        "夹耳R13、18高及中心距25":ranges(cylinders(13,1,(-25,0,0)),1,-9,-1.5)
            and ranges(cylinders(13,1,(-25,0,0)),1,1.5,9),
        "夹紧槽实际草图上下1点5且切至主孔":all(any(near(g(p,"Y"),v) for p in points) for v in [-.0015,.0015])
            and near(depth("夹紧槽"),.026),
        "夹耳上间隙孔11及沉孔18深3":ranges(cylinders(5.5,1,(-25,0,0)),1,1.5,6)
            and ranges(cylinders(9,1,(-25,0,0)),1,6,9),
        "M10底孔8点5通至槽":ranges(cylinders(4.25,1,(-25,0,0)),1,-9,-1.5),
        "M10装饰螺纹保存重开保留":m.FeatureByName("M10装饰螺纹") is not None,
        "图示各处R2过渡保存重开保留":all(m.FeatureByName(n) is not None for n in
            ["底座四角R2","腹板翼板过渡R2","腹板底座过渡R2","翼板底座过渡R2","夹耳套筒过渡R2"]),
    }
    # 从本图的斜翼板两面直接计算平面间距，排除以宽度误作法向厚度。
    planes=[f["Plane"] for f in rows if "Plane" in f and abs(f["Plane"][0])>.5 and abs(f["Plane"][1])>.3 and abs(f["Plane"][2])<1e-7]
    checks["翼板实际两平行面间距6"]=any(abs(abs(sum(a[i]*(b[3+i]-a[3+i]) for i in range(3)))-.006)<1e-7
        for a in planes for b in planes if abs(sum(a[i]*b[i] for i in range(3)))>.999999)
    if optimized:
        checks["底座孔向导及40毫米两实例原生阵列"]=m.FeatureByName("底座沉孔母孔") is not None \
            and g(m.FeatureByName("两组沉孔40阵列"),"GetTypeName2")=="LPattern" \
            and any(near(d["value"],.040) for d in dimensions(m,"两组沉孔40阵列"))
        checks["夹耳孔向导盲孔终止及沉孔深3"]=all(near(g(g(m.FeatureByName("夹耳沉孔向导"),"GetDefinition"),k),v)
            for k,v in [("EndCondition",0),("CounterBoreDepth",.003),("CounterBoreDiameter",.018)])
        checks["零额外重合基准面"]=all(m.FeatureByName(n) is None for n in ["套筒端面","底座背面","夹耳上表面","夹耳下表面"])
        sk=g(m.FeatureByName("套筒截面"),"GetSpecificFeature2")
        checks["套筒毛坯草图仅外圆后续统一切主孔"]=len(g(sk,"GetSketchSegments"))==1
    return checks


if len(sys.argv)>2 and sys.argv[2]=="finish":
    result=finish_review(sw,target,visual_notes="已对照PDF120印刷114：套筒开口及夹耳上下孔、斜臂T形截面、底座两组沉孔的入口侧与位置相符；腹板7和距底5已回读修正；按图示设置R2过渡。螺纹为装饰表达，无实体牙型；STEP未重导入。")
else:
    # 原生Rib在SW2024额外暴露一个同源派生草图节点；12张手工草图共13条检查记录。
    report=prepare_review(sw,target,check,sketch_count=13,views=["isometric","front_wire","right_wire","top_wire"])
    if optimized:
        quality=report["quality"]
        internal={"底座沉孔母孔原生轮廓","夹耳沉孔向导原生轮廓"}
        allowed=all(s["constrained_status"]==3 and not s["dangling"] and not s["errors"]
            and (not s["fix_count"] or s["name"] in internal) for s in quality["sketches"])
        assert internal.issubset({s["name"] for s in quality["sketches"]})
        report["checks"].pop("手工草图数量与完全定义无Fix无孤立")
        report["checks"]["手工草图完全定义无Fix；两张孔向导内置轮廓Fix单列例外"]=allowed
        report["passed"]=all(report["checks"].values())
        report["status"]="awaiting_visual_review" if report["passed"] else "failed"
        report["limits"]+=" 孔向导两张内置轮廓含原生Fix，已与实体孔径/孔深核对，不冒称全部无Fix。"
        (out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
        print(json.dumps({"final_passed":report["passed"],"failed":[k for k,v in report["checks"].items() if not v]},ensure_ascii=False))
