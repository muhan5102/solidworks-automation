"""分段装配第8届手动换向阀；只修改本轮装配和清单内的零件副本。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g
from shared.single_part import save
from sw_assembly import (add_component, add_mate5_checked, collect_mate_feature_summary,
                         create_math_transform, get_assembly_entity, get_component_feature_entity)
from sw_connect import _attach_existing_with_pid_check, solidworks_processes
from sw_review import save_preview

OUT = Path(sys.argv[1]).resolve()
PID, STAGE = int(sys.argv[2]), sys.argv[3]
TARGET = OUT / "手动换向阀_第8届.SLDASM"
PARTS = {k: (OUT / v["local"]).resolve()
         for k, v in json.loads((OUT / "inputs.json").read_text(encoding="utf-8"))["parts"].items()}
assert all(p.parent == OUT / "parts" and p.is_file() for p in PARTS.values())
state = solidworks_processes()
sw = _attach_existing_with_pid_check(state, state[PID])
asm = g(sw, "ActiveDoc")
assert g(sw, "GetProcessID") == PID and g(asm, "GetType") == 2
assert Path(g(asm, "GetPathName")).resolve() == TARGET
TITLE = g(asm, "GetTitle")
X, Y, Z = (1,0,0), (0,1,0), (0,0,1)
NX, NY, NZ = (-1,0,0), (0,-1,0), (0,0,-1)
UP, DOWN, ID = (X,Z,NY), (X,NZ,Y), (X,Y,Z)


def now():
    return datetime.now().astimezone().isoformat()


def checkpoint(label):
    assert asm.ForceRebuild3(False)
    save(asm, TARGET)
    asm.ShowNamedView2("*Isometric",7)
    asm.ViewZoomtofit2()
    asm.GraphicsRedraw2()
    with (OUT/"assembly-stages.jsonl").open("a", encoding="utf-8") as f:
        f.write(json.dumps({"time":now(), "stage":label,
                            "components":g(asm,"GetComponentCount",True)}, ensure_ascii=False)+"\n")
    print(label, flush=True)


def parts(key):
    return [c for c in (g(asm,"GetComponents",True) or [])
            if Path(g(c,"GetPathName")).resolve() == PARTS[key]]


def part(key):
    c, = parts(key)
    return c


def insert(key, xyz, basis=ID):
    c = add_component(asm, str(PARTS[key]), sw=sw)
    position(c,xyz,basis)
    return c


def position(c, xyz, basis, *, solve=False):
    transform = create_math_transform(sw, [v for vec in basis for v in vec]
                                    + [v/1000 for v in xyz] + [1,0,0,0])
    if solve:
        assert c.SetTransformAndSolve2(transform), g(c,"Name2")
    else:
        c.Transform2 = transform
    actual = g(g(c, "Transform2"), "ArrayData")
    assert math.dist(actual[9:12], [v/1000 for v in xyz]) < 1e-7, (g(c,"Name2"), actual)


def faces(c):
    model = g(c, "GetModelDoc2")
    assert model is not None, g(c, "Name2")
    for solid in g(model, "GetBodies2", 0, False):
        yield from g(solid, "GetFaces")


def surface(c, kind, predicate):
    candidates = []
    for f in faces(c):
        s = g(f,"GetSurface")
        if g(s, "Is"+kind) and predicate(g(s,kind+"Params")):
            candidates.append(f)
    if not candidates:
        raise RuntimeError(f"{g(c,'Name2')}找不到{kind}配合面")
    return get_assembly_entity(c,max(candidates,key=lambda f:g(f,"GetArea")))


def cyl(c, axis, radius, center=(0,0,0)):
    return surface(c,"Cylinder",lambda p: abs(abs(p[3+axis])-1)<1e-7
                   and abs(p[6]*1000-radius)<.015
                   and all(abs(p[i]*1000-center[i])<.02 for i in range(3) if i!=axis))


def plane_face(c, axis, level):
    return surface(c,"Plane",lambda p: abs(abs(p[axis])-1)<1e-7
                   and abs(p[axis+3]*1000-level)<.015)


def datum(c, name):
    return get_component_feature_entity(c, name, resolve=False)


def mate(a,b,kind,name,*,rotation=False,distance=0):
    asm.ClearSelection2(True)
    for index, obj in enumerate((a,b)):
        try:
            ok = obj.Select2(bool(index),1)
        except (AttributeError, TypeError):
            sd = g(asm.SelectionManager,"CreateSelectData")
            sd.Mark = 1
            ok = obj.Select4(bool(index),sd)
        assert ok,name
    assert asm.SelectionManager.GetSelectedObjectCount2(-1)==2
    add_mate5_checked(asm,kind,align=2,distance=distance/1000,lock_rotation=rotation,name=name)


def lock_pose(a,b,name):
    asm.ClearSelection2(True)
    sd = g(asm.SelectionManager,"CreateSelectData")
    sd.Mark = 1
    assert a.Select4(False,sd,False) and b.Select4(True,sd,False)
    add_mate5_checked(asm,16,align=2,name=name)


def remove_mates(names=None):
    feature = g(asm,"FirstFeature")
    locks = []
    while feature is not None:
        if g(feature,"GetTypeName2")=="MateGroup":
            f = g(feature,"GetFirstSubFeature")
            while f is not None:
                if (names is None and g(f,"GetTypeName2")=="MateLock") or (
                        names is not None and g(f,"Name") in names):
                    locks.append(f)
                f = g(f,"GetNextSubFeature")
        feature = g(feature,"GetNextFeature")
    assert len(locks)==(7 if names is None else len(names))
    asm.ClearSelection2(True)
    for i,f in enumerate(locks):
        assert f.Select2(bool(i),0)
    assert asm.Extension.DeleteSelection2(0)


def open_copy(key):
    assert not parts(key), f"{key} 已装入，禁止直接编辑当前装配引用"
    errors, warnings = (win32com.client.VARIANT(16387,0) for _ in range(2))
    d = sw.OpenDoc6(str(PARTS[key]),1,1,"",errors,warnings)
    assert d is not None and errors.value==0 and Path(g(d,"GetPathName")).resolve()==PARTS[key]
    return d


def correct_dimension(doc, feature, old_mm, new_mm):
    f = doc.FeatureByName(feature)
    display = g(f,"GetFirstDisplayDimension")
    found = []
    while display is not None:
        d = g(display,"GetDimension2",0)
        if abs(g(d,"SystemValue")*1000-old_mm)<1e-7:
            found.append(d)
        display = g(f,"GetNextDisplayDimension",display)
    if len(found)==1:
        found[0].SystemValue = new_mm/1000
        return new_mm
    if not found and old_mm < 20:
        return correct_dimension(doc,feature,old_mm*2,new_mm*2)
    raise RuntimeError(f"{feature} 驱动尺寸{old_mm}未唯一找到")


def change_spring(doc, height, pitch):
    eq = g(doc,"GetEquationMgr")
    member = eq._oleobj_.GetIDsOfNames("Equation")
    values = {"H":height,"p":pitch}
    for i in range(g(eq,"GetCount")):
        line = eq._oleobj_.Invoke(member,0,pythoncom.DISPATCH_PROPERTYGET,True,i)
        key = line.split("=")[0].strip().strip('"')
        if key in values:
            eq._oleobj_.Invoke(member,0,pythoncom.DISPATCH_PROPERTYPUT,False,i,f'"{key}" = {values.pop(key)}mm')
    assert not values
    g(eq,"EvaluateAll")


if STAGE == "fit-parts":
    corrections = []
    for key, feature, old, new in [("cover","轴向截面",17,18),
                                    ("o10","截面草图",5.5,5.6),("o13","截面草图",16,16.25),
                                    ("big_spring",None,6.5,1.0532),("small_spring",None,2.3,.695)]:
        doc = open_copy(key)
        if feature:
            correct_dimension(doc,feature,old,new)
        else:
            change_spring(doc,old,new)
        assert doc.ForceRebuild3(False) and len(g(doc,"GetBodies2",0,False))==1
        save(doc,PARTS[key])
        corrections.append({"part":key,"dimension_or_assembly_height":new if feature else old})
        sw.CloseDoc(g(doc,"GetTitle"))
        print(key,"装配副本已保存",flush=True)
    status = win32com.client.VARIANT(16387,0)
    assert sw.ActivateDoc3(TITLE,False,0,status) is not None
    (OUT/"fit-parts.json").write_text(json.dumps(corrections,indent=2),encoding="utf-8")

elif STAGE == "lower":
    assert g(asm,"GetComponentCount",True)==1 and (OUT/"fit-parts.json").is_file()
    body = part("body")
    connector = insert("connector",(0,0,20),UP)
    mate(cyl(body,2,17.5),cyl(connector,1,17.5),1,"连接件止口同心",rotation=True)
    mate(plane_face(body,2,20),plane_face(connector,1,0),0,"连接件贴阀体肩面")
    checkpoint("连接件止口与阀体肩面")
    disc = insert("disc",(0,0,31),DOWN)
    a=2**-.5
    gasket = insert("gasket",(0,0,24),((a,-a,0),(a,a,0),Z))
    mate(cyl(disc,1,17),cyl(gasket,2,17),1,"配气盘与盘垫粘接同轴定相",rotation=True)
    mate(plane_face(disc,1,4),plane_face(gasket,2,3),0,"配气盘与盘垫粘接面")
    ring = insert("o13",(0,0,23))
    mate(cyl(body,2,16.25),cyl(ring,2,16.25),1,"外O圈槽根同心",rotation=True)
    mate(plane_face(body,2,21),plane_face(ring,2,-2),0,"外O圈贴槽肩")
    checkpoint("粘接叠层及外密封圈")

elif STAGE == "upper":
    assert g(asm,"GetComponentCount",True)==5
    connector, disc = part("connector"), part("disc")
    cover = insert("cover",(0,0,41),(Y,Z,X))
    mate(cyl(connector,1,18),cyl(cover,1,18),1,"阀上盖主轴同心可转")
    mate(plane_face(connector,1,21),plane_face(cover,1,0),0,"阀上盖杯口贴肩")
    core = insert("core",(0,0,29.3),UP)
    mate(cyl(cover,1,3.5),cyl(core,1,3.5),1,"内芯与上盖同轴定相",rotation=True)
    mate(datum(core,"Top Plane"),datum(cover,"Top Plane"),5,"内芯轴向止位",distance=11.7)
    mate(cyl(disc,1,17),cyl(core,1,7),1,"配气盘随内芯扁位转动",rotation=True)
    mate(datum(disc,"Top Plane"),datum(core,"Top Plane"),5,"内芯和配气盘弹簧座间距",distance=1.7)
    checkpoint("上盖、内芯与配气盘回转链")
    ring = insert("o10",(0,0,38.8))
    mate(cyl(core,1,5.6),cyl(ring,2,5.6),1,"内密封圈与内芯槽根同心",rotation=True)
    mate(plane_face(core,1,7.5),plane_face(ring,2,-2),0,"内密封圈贴芯轴槽肩")
    checkpoint("内密封圈")

elif STAGE == "hardware":
    assert g(asm,"GetComponentCount",True)==8
    body, connector, cover, core = [part(k) for k in ("body","connector","cover","core")]
    for i,(x,y) in enumerate(((20.75,0),(0,20.75),(-20.75,0),(0,-20.75)),1):
        screw = insert("m4",(x,y,-.1),UP)
        mate(cyl(body,2,2.5,(x,y,0)),cyl(screw,1,2),1,f"M4-{i}穿孔同心",rotation=True)
        mate(plane_face(body,2,3),plane_face(screw,1,3.1),0,f"M4-{i}头部贴沉孔底")
    screw = insert("m5",(0,0,58),DOWN)
    mate(cyl(cover,1,2.6),cyl(screw,1,2.5),1,"M5穿上盖中心同心",rotation=True)
    mate(plane_face(cover,1,12),plane_face(screw,1,5),0,"M5头部承面")
    checkpoint("五个联接螺钉")

elif STAGE in ("detents","center-detents"):
    assert len(parts("m4"))==4 and len(parts("m5"))==1
    repair = STAGE != "detents"
    if repair:
        assert g(asm,"GetComponentCount",True)==20
        remove_mates()
    else:
        assert g(asm,"GetComponentCount",True) in (13,14) and not parts("pin")
    connector, cover, core, body = [part(k) for k in ("connector","cover","core","body")]
    body_transform = g(g(body,"Transform2"),"ArrayData")
    assert all(abs(a-b)<1e-8 for a,b in zip(body_transform[:9],[v for basis in ID for v in basis]))
    base_z = body_transform[11]*1000
    position(cover,(0,0,41+base_z),(Y,Z,X),solve=True)
    assert asm.ForceRebuild3(False)
    cover_basis = g(g(cover,"Transform2"),"ArrayData")[:9]
    assert math.dist(cover_basis,(*Y,*Z,*X)) < 1e-6, cover_basis

    def in_body(key,local,basis,existing=None):
        xyz = [v+base*1000 for v,base in zip(local,body_transform[9:12])]
        if existing is not None:
            position(existing,xyz,basis)
            return existing
        return insert(key,xyz,basis)

    c, s = math.cos(math.radians(22.6)), math.sin(math.radians(22.6))
    location, basis = (0,-12.719629,52.472892),(X,(0,-c,s),(0,-s,-c))
    handle = in_body("handle",location,basis,part("handle") if parts("handle") else None)
    lock_pose(cover,handle,"手柄M8装入锁定")
    pin = in_body("pin",(0,14,46),DOWN,part("pin") if repair else None)
    lock_pose(connector,pin,"限位挡芯轴入连接件孔")
    old_balls, old_springs = (parts("ball"),parts("big_spring")) if repair else ([],[])
    for i,x in enumerate((-14,14),1):
        ball = in_body("ball",(x,0,42.05),UP,old_balls[i-1] if repair else None)
        spring = in_body("big_spring",(x,0,47.5),UP,old_springs[i-1] if repair else None)
        lock_pose(cover,ball,f"钢球{i}中位就位")
        lock_pose(cover,spring,f"钢球弹簧{i}压缩态就位")
    spring = in_body("small_spring",(0,0,29),UP,part("small_spring") if repair else None)
    lock_pose(core,spring,"芯轴弹簧压缩态就位")
    checkpoint("手柄、挡芯轴、钢球与两组装配态弹簧")

elif STAGE == "key-phase":
    assert g(asm,"GetComponentCount",True)==20
    body, cover, core, disc = [part(k) for k in ("body","cover","core","disc")]
    base = g(g(body,"Transform2"),"ArrayData")[11]*1000
    names = {"内芯与上盖同轴定相","配气盘随内芯扁位转动"}
    remove_mates(names)
    position(core,(0,0,29.3+base),UP,solve=True)
    position(disc,(0,0,31+base),DOWN,solve=True)
    # 锁定前明确两组扁面的相位。重建后回读三件，防止求解绕主轴偷转。
    mate(cyl(cover,1,3.5),cyl(core,1,3.5),1,"内芯与上盖同轴定相",rotation=True)
    mate(cyl(disc,1,17),cyl(core,1,7),1,"配气盘随内芯扁位转动",rotation=True)
    assert asm.ForceRebuild3(False)
    for comp,basis in [(cover,(Y,Z,X)),(core,UP),(disc,DOWN)]:
        assert math.dist(g(g(comp,"Transform2"),"ArrayData")[:9],
                         [v for axis in basis for v in axis])<1e-6,g(comp,"Name2")
    checkpoint("上盖与内芯扁位、配气盘槽中位对正")

elif STAGE == "seal-phase":
    disc, gasket = [part(k) for k in ("disc","gasket")]
    remove_mates({"配气盘与盘垫粘接同轴定相"})
    a = 2**-.5
    basis = ((a,-a,0),(a,a,0),Z)
    xyz = [value*1000 for value in g(g(gasket,"Transform2"),"ArrayData")[9:12]]
    position(gasket,xyz,basis,solve=True)
    mate(cyl(disc,1,17),cyl(gasket,2,17),1,"配气盘与盘垫粘接同轴定相",rotation=True)
    assert asm.ForceRebuild3(False)
    assert math.dist(g(g(gasket,"Transform2"),"ArrayData")[:9],
                     [v for axis in basis for v in axis])<1e-6
    checkpoint("盘垫两定位柱对正配气盘孔")

elif STAGE == "snapshot":
    evidence = {"time":now(),"mates":collect_mate_feature_summary(asm),"components":[]}
    for key in PARTS:
        for c in parts(key):
            evidence["components"].append({"key":key,"name":g(c,"Name2"),"path":PARTS[key].name,
                                           "transform":list(g(g(c,"Transform2"),"ArrayData")),
                                           "fixed":g(c,"IsFixed")})
    (OUT/"assembly-evidence.json").write_text(json.dumps(evidence,ensure_ascii=False,indent=2),encoding="utf-8")
    save_preview(asm,str(OUT/"assembly-isometric.bmp"),"isometric")
    print("组件",len(evidence["components"]),flush=True)

elif STAGE == "hold-neutral":
    assert g(asm,"GetComponentCount",True)==20
    assert asm.FeatureByName("中位展示定位") is None
    mate(datum(part("cover"),"Front Plane"),datum(part("connector"),"Right Plane"),3,"中位展示定位")
    checkpoint("静态装配中位稳定")

elif STAGE == "key-face-mates":
    cover, core, disc = [part(k) for k in ("cover","core","disc")]
    remove_mates({"内芯与上盖同轴定相","配气盘随内芯扁位转动"})
    mate(cyl(cover,1,3.5),cyl(core,1,3.5),1,"内芯与上盖同轴")
    mate(cyl(disc,1,17),cyl(core,1,7),1,"配气盘与内芯同轴")
    mate(datum(cover,"Front Plane"),datum(core,"Right Plane"),3,"上盖槽与内芯扁位平行")
    mate(datum(core,"Right Plane"),datum(disc,"Right Plane"),3,"配气盘槽与内芯扁位平行")
    checkpoint("两处扁位角向由平行配合定位")

elif STAGE in ("explode","finish-explode"):
    assert g(asm,"GetComponentCount",True)==20
    cfg = g(g(asm,"ConfigurationManager"),"ActiveConfiguration")
    sequence = [
        ("m5",155),("handle",-70),("cover",125),("big_spring",105),
        ("ball",90),("pin",70),("core",64),("o10",48),("connector",44),
        ("small_spring",28),("disc",20),("gasket",12),("o13",6),("m4",-25),
    ]
    if STAGE=="explode":
        assert g(cfg,"GetNumberOfExplodeSteps")==0
        assert g(asm,"CreateExplodedView")
        for key,distance in sequence:
            asm.ClearSelection2(True)
            sd = g(asm.SelectionManager,"CreateSelectData")
            sd.Mark = 1
            for i,component in enumerate(parts(key)):
                assert component.Select4(bool(i),sd,False),key
            if key=="handle":
                assert cyl(part("handle"),1,6).Select2(True,2)
            err = win32com.client.VARIANT(16387,0)
            step = cfg.AddExplodeStep2(abs(distance)/1000,2,distance<0,0,0,
                                       False,False,False,err)
            assert step is not None and err.value==0,(key,err.value)
    assert g(cfg,"GetNumberOfExplodeSteps")==len(sequence)+1
    first_step = g(cfg,"GetExplodeStep",0)
    assert g(first_step,"Name")=="链1"
    assert cfg.DeleteExplodeStep("链1")
    assert g(cfg,"GetNumberOfExplodeSteps")==len(sequence)
    assert asm.ShowExploded(True)
    try:
        asm.ViewZoomtofit2()
        save_preview(asm,str(OUT/"assembly-exploded.bmp"),"isometric")
    finally:
        assert asm.ShowExploded(False)
    checkpoint("原生爆炸视图14步，恢复合拢保存")

elif STAGE == "final":
    assert g(asm,"GetComponentCount",True)==20
    save(asm,TARGET)
    sw.CloseDoc(TITLE)
    e,w = (win32com.client.VARIANT(16387,0) for _ in range(2))
    reopened = sw.OpenDoc6(str(TARGET),2,1,"",e,w)
    assert reopened is not None and e.value==0 and Path(g(reopened,"GetPathName")).resolve()==TARGET
    assert reopened.ForceRebuild3(False)
    components = list(g(reopened,"GetComponents",True))
    counts = {}
    for component in components:
        path = Path(g(component,"GetPathName")).resolve()
        assert path in PARTS.values() and g(component,"GetModelDoc2") is not None
        assert not g(component,"IsSuppressed")
        counts[path.name] = counts.get(path.name,0)+1
    assert len(components)==20 and len(counts)==15
    for key,count in {"m4":4,"ball":2,"big_spring":2}.items():
        assert counts[PARTS[key].name]==count
    cfg = g(g(reopened,"ConfigurationManager"),"ActiveConfiguration")
    assert g(cfg,"GetNumberOfExplodeSteps")==14
    mates = collect_mate_feature_summary(reopened)
    assert len(mates)==35
    from sw_review import collect_model_summary
    health = collect_model_summary(reopened)
    (OUT/"reopened.json").write_text(json.dumps({"time":now(),"components":counts,
             "mate_count":len(mates)-1,"explode_steps":14,"mates":mates,"feature_issues":{k:health.get(k) for k in
             ("faulty_features","warning_features","unknown_feature_health")}},ensure_ascii=False,indent=2),encoding="utf-8")
    reopened.ViewZoomtofit2()
    save_preview(reopened,str(OUT/"assembly-isometric.bmp"),"isometric")
    save(reopened,TARGET)
    print("重开：15种、20实例、34条配合，引用全部读取",flush=True)

else:
    raise ValueError(STAGE)
