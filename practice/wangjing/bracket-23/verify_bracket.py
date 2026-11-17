"""LJT05.23集中保存、导出、重开；按原图独立列明实体与空腔检查。"""
from pathlib import Path
import sys,json,math
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3];sys.path.insert(0,str(ROOT/"practice"));sys.path.insert(0,str(Path.home()/".codex/skills/solidworks-automation/scripts"))
from shared.single_part import prepare_review,finish_review,dimensions
from sw_connect import connect_solidworks,get_com_member as g
out=Path(sys.argv[1]).resolve();target=out/"支架_LJT05.23.SLDPRT"
sw,m=connect_solidworks(2024,.1);assert g(sw,"GetProcessID")==42548


def near(a,b):return abs(a-b)<1e-7


def check(m,rows):
    def cylinders(radius,axis,center):
        return [f for f in rows if "Cylinder" in f and near(f["Cylinder"][6],radius/1000)
            and abs(f["Cylinder"][3+axis])>.999999
            and all(near(f["Cylinder"][i],center[i]/1000) for i in range(3) if i!=axis)]
    def ranges(faces,axis,a,b):return any(abs(f["box"][axis]-a)<1e-4 and abs(f["box"][axis+3]-b)<1e-4 for f in faces)
    def depth(name):return g(g(m.FeatureByName(name),"GetDefinition"),"GetDepth",True)
    minor=42-1.0825317547305484*2
    phi27=cylinders(13.5,1,(0,0,0));thread=cylinders(minor/2,1,(0,0,0));groove=cylinders(23,1,(0,0,0))
    neck=cylinders(22.5,1,(0,0,0));body=cylinders(30,1,(0,0,0))
    holes=[cylinders(2.5,0,(0,y,z)) for y in [85,135] for z in [-30,30]]
    radial=cylinders(2.5,1,(-45,0,0))
    radial_neighbors=set()
    for b in g(m,"GetBodies2",0,False):
        for face in g(b,"GetFaces"):
            s=g(face,"GetSurface")
            if not g(s,"IsCylinder"):continue
            p=list(g(s,"CylinderParams"))
            if not (near(p[6],.0025) and abs(p[4])>.999999 and near(p[0],-.045)):continue
            for edge in g(face,"GetEdges"):
                for adjacent in g(edge,"GetTwoAdjacentFaces2"):
                    if adjacent is None:continue
                    su=g(adjacent,"GetSurface")
                    if g(su,"IsCylinder"):
                        cp=list(g(su,"CylinderParams"))
                        if near(cp[6],.0075) and abs(cp[3])>.999999:radial_neighbors.add("主孔15")
                    elif g(su,"IsCone"):radial_neighbors.add("沉头锥面")
    plate=g(g(m.FeatureByName("安装板截面"),"GetSpecificFeature2"),"GetSketchSegments")
    points=[g(e,"GetStartPoint2") for e in plate]
    xs=[g(p,"X") for p in points];ys=[g(p,"Y") for p in points]
    sketch=g(m.FeatureByName("背面35高避空带轮廓"),"GetSpecificFeature2")
    tr=list(g(g(g(sketch,"ModelToSketchTransform"),"Inverse"),"ArrayData"));wp=[]
    for e in g(sketch,"GetSketchSegments"):
        if g(e,"GetType")!=0:continue
        p=g(e,"GetStartPoint2");q=[g(p,"X"),g(p,"Y"),0]
        wp.append([sum(q[j]*tr[j*3+i] for j in range(3))*tr[12]+tr[9+i] for i in range(3)])
    relief_y=[p[1] for p in wp]
    torus_outer=[f for f in rows if "Torus" in f and near(f["Torus"][1],.037)
        and near(f["Torus"][6],.027) and near(f["Torus"][7],.003)]
    tips=[f for f in rows if "Torus" in f and near(f["Torus"][7],.003)]
    d=g(m.FeatureByName("径向沉头孔"),"GetDefinition")
    ribsk=g(m.FeatureByName("弯臂腹板边界"),"GetSpecificFeature2")
    ribarc=next(e for e in g(ribsk,"GetSketchSegments") if g(e,"GetType")==1)
    rc=g(ribarc,"GetCenterPoint2")
    upper_sk=g(m.FeatureByName("上套筒外形截面"),"GetSpecificFeature2")
    upper_pts=[g(e,"GetStartPoint2") for e in g(upper_sk,"GetSketchSegments") if not g(e,"ConstructionGeometry")]
    ux=[g(p,"X") for p in upper_pts]
    wing_sk=g(m.FeatureByName("弯臂翼缘截面"),"GetSpecificFeature2")
    wing_arcs=[e for e in g(wing_sk,"GetSketchSegments") if g(e,"GetType")==1]
    outside=next(e for e in wing_arcs if near(g(e,"GetRadius"),.038))
    wc=g(outside,"GetCenterPoint2")
    checks={
        "安装板80宽70高12厚与孔中心高110":near(depth("安装板"),.080)
            and near(max(xs)-min(xs),.012) and near(max(ys)-min(ys),.070) and near((max(ys)+min(ys))/2,.110),
        "下座外圆60及顶颈45高57":ranges(body,1,0,37) and any(abs(f["box"][4]-57)<1e-4 for f in neck),
        "图示40高外侧肩部实际R3":bool(torus_outer),
        "下座27H7上孔从33至56及C1孔口":ranges(phi27,1,33,56),
        "M42简化小径从2至30入口C2":ranges(thread,1,2,30),
        "内槽实际直径46轴向30至33":ranges(groove,1,30,33),
        "上15H7主孔只保留两端C1轴向负72至负36":ranges(cylinders(7.5,0,(0,110,0)),0,-72,-36),
        "上套筒外径30及左凸台28":bool(cylinders(15,0,(0,110,0))) and bool(cylinders(14,0,(0,110,0))),
        "上套筒长38左伸5及右端距下座轴35":near(max(ux)-min(ux),.038)
            and near(min(ux),min(xs)-.005) and near(max(ux),-.035),
        "腹板相对上套筒右端后退3":near(g(rc,"X")-.030,max(ux)-.003),
        "翼缘相对安装板背面偏4且外弧理论接下座高度12":near(g(wc,"X")-.038,min(xs)+.004)
            and near(g(wc,"Y")-math.sqrt(.038**2-(-.030-g(wc,"X"))**2),.012),
        "四M6底孔直径5按60乘50定位且贯通12":all(ranges(faces,0,-68,-56) for faces in holes),
        "四孔为一个原生二维线性阵列":g(m.FeatureByName("四个M6孔二维阵列"),"GetTypeName2")=="LPattern"
            and all(any(near(d["value"],v) for d in dimensions(m,"四个M6孔二维阵列")) for v in [.050,.060]),
        # GetBox为近似包围盒；用相对原图主轴110的材料侧和实际相邻面确认单壁，
        # 不以解析交线117.071等作为近似包围盒的下限。
        "径向5孔位于距右端10且只穿上壁":len(radial)==1 and radial[0]["box"][1]>110
            and abs(radial[0]["box"][4]-123.5)<1e-4 and radial_neighbors=={"主孔15","沉头锥面"},
        "径向孔为原生孔向导8乘90度沉头":near(g(d,"CounterSinkDiameter"),.008) and near(g(d,"CounterSinkAngle"),math.pi/2),
        "弯臂翼缘38宽实际内R30外R38":near(depth("弯臂翼缘"),.038)
            and any("Cylinder" in f and near(f["Cylinder"][0],-.026) and near(f["Cylinder"][6],.030) for f in rows)
            and any("Cylinder" in f and near(f["Cylinder"][0],-.026) and near(f["Cylinder"][6],.038) for f in rows),
        "原生腹板厚6且边界R30":g(m.FeatureByName("弯臂腹板"),"GetTypeName2")=="Rib"
            and near(g(g(m.FeatureByName("弯臂腹板"),"GetDefinition"),"Thickness"),.006)
            and any(near(d["value"],.030) for d in dimensions(m,"弯臂腹板边界")),
        "R30与45外壁理论交点低于57顶面5且5为驱动":near(g(rc,"Y")-math.sqrt(.030**2-(-.0225-g(rc,"X"))**2),.052)
            and any(near(d["value"],.005) and d["driven"]==2 for d in dimensions(m,"弯臂腹板边界")),
        "T形断面翼缘与腹板自由端已收圆":all(m.FeatureByName(n) is not None for n in ["腹板自由端R3","翼缘自由端R4"]) and bool(tips),
        "背面35高避空带相对110对称":near(max(relief_y)-min(relief_y),.035)
            and near((max(relief_y)+min(relief_y))/2,.110),
        "避空深4且实体中存在负64槽底平面":near(depth("背面避空保留28凸台"),.004)
            and any("Plane" in f and near(abs(f["Plane"][0]),1) and near(f["Plane"][3],-.064)
                and f["box"][1]>90 and f["box"][4]>120 for f in rows),
        "避空带保留28凸台实体":ranges(cylinders(14,0,(0,110,0)),0,-73,-65.5),
        "用户允许自定槽底R1点5且已保留":m.FeatureByName("避空槽底R1点5_自定") is not None
            and any(near(d["value"],.0015) for d in dimensions(m,"避空槽底R1点5_自定")),
        "四角R10及主体18边R3三口C1保留":all(m.FeatureByName(n) is not None for n in ["安装板四角R10","同值铸造过渡R3","三处孔口C1"]),
        "M6与M42装饰螺纹持久保留":all(m.FeatureByName(n) is not None for n in ["M6装饰螺纹","M42装饰螺纹"]),
    }
    return checks


if len(sys.argv)>2 and sys.argv[2]=="finish":
    assert g(m.SketchManager,"ActiveSketch") is None
    g(m,"ActiveView").DisplayMode=5
    finish_review(sw,target,visual_notes="已对照PDF132印刷126与Front/Left/Top投影：上安装板及背面35高避空、28凸台、四孔、弯臂T断面、下座内腔和环槽、40高外侧R3及孔入口方向均逐项核对；R30与45外壁理论交点低于57顶面5已直接驱动，并复核下游过渡。槽底R1.5是用户允许的自定参数；M6与M42为装饰表达，STEP未重导入。")
else:
    g(m,"ActiveView").DisplayMode=5
    expected={"下座外形截面","安装板截面","上套筒外形截面","弯臂翼缘截面","弯臂腹板边界","下座内腔截面",
              "上套筒通孔草图","M6安装底孔草图","径向沉头孔孔位","径向沉头孔原生轮廓","背面35高避空带轮廓"}
    def sketch_quality(quality):
        records=quality["sketches"]
        names={s["name"].split("<")[0] for s in records}
        allowed=names==expected and all(s["constrained_status"]==3 and not s["dangling"] and not s["errors"]
            and (not s["fix_count"] or s["name"]=="径向沉头孔原生轮廓") for s in records)
        return "设计草图完全定义无Fix无孤立；孔向导内部Fix单列例外",allowed
    report=prepare_review(sw,target,check,sketch_count=11,views=["isometric","front_wire","left_wire","top_wire"],sketch_quality_check=sketch_quality)
    records=report["quality"]["sketches"]
    names={s["name"].split("<")[0] for s in records}
    allowed=names==expected and all(s["constrained_status"]==3 and not s["dangling"] and not s["errors"]
        and (not s["fix_count"] or s["name"]=="径向沉头孔原生轮廓") for s in records)
    assert report["checks"]["设计草图完全定义无Fix无孤立；孔向导内部Fix单列例外"]==allowed
    report["passed"]=all(report["checks"].values());report["status"]="awaiting_visual_review" if report["passed"] else "failed"
    report["limits"]+=" 槽底R1.5经用户允许自定；孔向导内部轮廓含原生Fix；M6孔向导入口两次空返回，采用母孔切除加二维阵列。"
    (out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps({"passed":report["passed"],"failed":[k for k,v in report["checks"].items() if not v],"sketch_names":sorted(names)},ensure_ascii=False))
