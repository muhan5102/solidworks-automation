"""LJT05.23内腔与孔；在全部加料之后执行。"""
from pathlib import Path
import sys,json,math,time
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/"practice"));sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,dimension_existing_relations,sketch,sketch_line,round_edges,offset_plane
from sw_connect import connect_solidworks
from sw_part import extrude_cut,_ensure_sketch_selected,linear_pattern
from shared.single_part import relation,circle,anchor,drive
from sw_review import save_preview
sw,m=connect_solidworks(2024,.1);out=Path(sys.argv[1]).resolve();target=out/"支架_LJT05.23.SLDPRT"
assert g(sw,"GetProcessID")==42548 and Path(g(m,"GetPathName")).resolve()==target
timer=time.perf_counter()


def poly(points):
    sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
    try:es=[sketch_line(m,a[0]/1000,a[1]/1000,b[0]/1000,b[1]/1000) for a,b in zip(points,points[1:]+points[:1])]
    finally:sm.AddToDB=old
    for a,b in zip(es,es[1:]+es[:1]):relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
    for e,a,b in zip(es,points,points[1:]+points[:1]):
        if a[0]==b[0] or a[1]==b[1]:relation(m,[e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
    if points[0]==(0,0):anchor(m,g(es[0],"GetStartPoint2"),x=True,y=True)
    return es


def select_face(axis,pos,point):
    faces=[]
    for b in g(m,"GetBodies2",0,False):
        for f in g(b,"GetFaces"):
            s=g(f,"GetSurface")
            if not g(s,"IsPlane"):continue
            p=list(g(s,"PlaneParams"));box=list(g(f,"GetBox"))
            if abs(abs(p[axis])-1)<1e-8 and abs(p[3+axis]-pos)<1e-8 and all(box[i]-1e-8<=point[i]<=box[i+3]+1e-8 for i in range(3)):faces.append(f)
    assert len(faces)==1,(axis,pos,len(faces))
    m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData");sd.X,sd.Y,sd.Z=point
    assert faces[0].Select4(False,sd)


def locate_wizard(f,name,world):
    f.Name=name;sf=g(f,"GetFirstSubFeature");sf.Name=name+"孔位";g(sf,"GetNextSubFeature").Name=name+"原生轮廓"
    m.ClearSelection2(True);assert sf.Select2(False,0);m.EditSketch()
    sk=g(m.SketchManager,"ActiveSketch");points=list(g(sk,"GetSketchPoints2"));assert len(points)==1
    p=points[0]
    if abs(g(p,"X"))<1e-9 or abs(g(p,"Y"))<1e-9:anchor(m,p,x=abs(g(p,"X"))<1e-9,y=abs(g(p,"Y"))<1e-9)
    dimension_existing_relations(m)
    a=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"));q=[g(p,"X"),g(p,"Y"),g(p,"Z")]
    xyz=[sum(q[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
    assert max(abs(xyz[i]-world[i]) for i in range(3))<1e-7,(name,xyz,world)
    m.SketchManager.InsertSketch(True)


old=sw.GetUserPreferenceToggle(10);sw.SetUserPreferenceToggle(10,False)
try:
    minor=42-1.0825317547305484*2
    if "--resume" not in sys.argv:
        # 预先按ISO牙型关系确定简化小径，非加工底孔/制造验收。
        minor=42-1.0825317547305484*2
        with sketch(m,"Front Plane") as ref:
            pts=[(0,0),(minor/2+2,0),(minor/2,2),(minor/2,30),(23,30),(23,33),(13.5,33),(13.5,57),(0,57)]
            es=poly(pts)
            sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
            try:axis=sm.CreateCenterLine(0,0,0,0,.057,0)
            finally:sm.AddToDB=db
            relation(m,[g(axis,"GetStartPoint2"),g(es[0],"GetStartPoint2")],"sgCOINCIDENT")
            relation(m,[g(axis,"GetEndPoint2"),g(es[-1],"GetStartPoint2")],"sgCOINCIDENT")
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name="下座内腔截面"
        _ensure_sketch_selected(m,"下座内腔截面")
        sd=g(m.SelectionManager,"CreateSelectData");sd.Mark=16;assert axis.Select4(True,sd)
        f=m.FeatureManager.FeatureRevolve2(True,True,False,True,False,False,0,0,2*math.pi,0,False,False,0,0,0,0,0,True,False,True)
        assert f is not None;f.Name="下座内腔与环槽"
        # Φ15精密通孔直接以左端面建立草图；不新增重合基准面。
        select_face(0,-.073,(-.073,.110,0));m.SketchManager.InsertSketch(True)
        sk=g(m.SketchManager,"ActiveSketch");a=list(g(g(sk,"ModelToSketchTransform"),"ArrayData"));inv=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"))
        world=(-.073,.110,0);q=[sum(world[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
        c=circle(m,q[0],q[1],.0075)
        if abs(q[0])<1e-9 or abs(q[1])<1e-9:anchor(m,g(c,"GetCenterPoint2"),x=abs(q[0])<1e-9,y=abs(q[1])<1e-9)
        drive(m,[c],(q[0]+.02,q[1]+.02,0),.015,"diameter");dimension_existing_relations(m)
        m.SketchManager.InsertSketch(True)
        f=g(m,"FirstFeature");last=None
        while f is not None:
            if g(f,"GetTypeName2")=="ProfileFeature":last=f
            f=g(f,"GetNextFeature")
        last.Name="上套筒通孔草图"
        f=extrude_cut(m,"上套筒通孔草图",0,direction=inv[6]>0);assert f is not None;f.Name="上套筒通孔"
        assert m.ForceRebuild3(False) and save_document(m,str(target));print("[孔腔] 下座内腔/槽与上通孔已保存",flush=True)
    # ISO攻丝底孔，母孔位于50×60矩形的左下角。螺纹采用原生装饰表达。
    select_face(0,-.056,(-.056,.085,-.030))
    # 本机两个ISO普通孔库入口均返回None，按止损规则转已验证的母孔切除。
    m.SketchManager.InsertSketch(True)
    sk=g(m.SketchManager,"ActiveSketch");t=list(g(g(sk,"ModelToSketchTransform"),"ArrayData"));inv=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"))
    point=(-.056,.085,-.030);q=[sum(point[j]*t[j*3+i] for j in range(3))*t[12]+t[9+i] for i in range(3)]
    c=circle(m,q[0],q[1],.0025);drive(m,[c],(q[0]+.012,q[1]+.012,0),.005,"diameter");dimension_existing_relations(m)
    m.SketchManager.InsertSketch(True)
    f=g(m,"FirstFeature");last=None
    while f is not None:
        if g(f,"GetTypeName2")=="ProfileFeature":last=f
        f=g(f,"GetNextFeature")
    last.Name="M6安装底孔草图"
    f=extrude_cut(m,"M6安装底孔草图",0,direction=-inv[6]>0);assert f is not None;f.Name="M6安装底孔"
    matches=[e for e,p,r in round_edges(m) if abs(p[0]+56)<1e-5 and abs(p[1]-85)<1e-5 and abs(p[2]+30)<1e-5 and abs(r-2.5)<1e-5]
    assert len(matches)==1
    m.ClearSelection2(True);assert matches[0].Select2(False,0)
    th=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M6x1.0",.006,0,.012,"4×M6-6H，贯通12")
    assert th is not None;th.Name="M6装饰螺纹"
    p=linear_pattern(m,"M6安装底孔",0,1,0,.050,2,0,0,1,.060,2)
    assert p is not None;p.Name="四个M6孔二维阵列"
    assert m.ForceRebuild3(False) and save_document(m,str(target));print("[孔群] 四个M6孔50×60二维阵列已保存",flush=True)
    # 圆柱顶面的径向孔需要切平面，区别于重合模型平面的冗余基准。
    plane=offset_plane(m,"Top Plane",125,"径向孔切平面")
    m.ClearSelection2(True)
    from sw_connect import create_empty_dispatch_variant
    assert m.Extension.SelectByID2(plane,"PLANE",-.045,.125,0,False,0,create_empty_dispatch_variant(),0)
    f=m.FeatureManager.HoleWizard5(1,8,141,"M3",0,.005,.010,-1,
        .008,math.pi/2,0,-1,-1,0,0,-1,0,-1,-1,-1,"",False,False,True,False,False,False)
    if f is not None:
        locate_wizard(f,"径向沉头孔",(-.045,.125,0))
        radial_method="HoleWizard5"
    else:
        # 曲面孔入口未能建成时，整合单壁直孔及90度锥口为一次回转切除。
        with sketch(m,"Front Plane") as ref:
            es=poly([(-45,115),(-42.5,115),(-42.5,123.5),(-41,125),(-45,125)])
            sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
            try:axis=sm.CreateCenterLine(-.045,.115,0,-.045,.125,0)
            finally:sm.AddToDB=db
            relation(m,[g(axis,"GetStartPoint2"),g(es[0],"GetStartPoint2")],"sgCOINCIDENT")
            relation(m,[g(axis,"GetEndPoint2"),g(es[-1],"GetStartPoint2")],"sgCOINCIDENT")
            dimension_existing_relations(m)
        m.FeatureByName(str(ref)).Name="径向沉头孔截面";_ensure_sketch_selected(m,"径向沉头孔截面")
        sd=g(m.SelectionManager,"CreateSelectData");sd.Mark=16;assert axis.Select4(True,sd)
        f=m.FeatureManager.FeatureRevolve2(True,True,False,True,False,False,0,0,2*math.pi,0,False,False,0,0,0,0,0,True,False,True)
        assert f is not None;f.Name="径向沉头孔";radial_method="一次回转切除"
        m.ClearSelection2(True);assert m.FeatureByName(plane).Select2(False,0);assert m.Extension.DeleteSelection2(0)
    (out/"hole-method.json").write_text(json.dumps({"M6":"母孔切除+原生二维阵列，装饰螺纹","M6_wizard_attempts":"ISO145与144均返回None，未继续试规格","radial":radial_method},ensure_ascii=False,indent=2),encoding="utf-8")
    assert m.ForceRebuild3(False)
    # M42底孔口C2已经整合到内腔截面，螺纹从Y2开始至Y30环槽。
    candidates=[(e,p,r) for e,p,r in round_edges(m) if abs(p[1]-2)<1e-5 and abs(r-minor/2)<1e-5]
    assert len(candidates)==1
    m.ClearSelection2(True);assert candidates[0][0].Select2(False,0)
    f=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M42x2",.042,0,.028,"M42×2-6H，底面至环槽30，入口C2")
    assert f is not None;f.Name="M42装饰螺纹"
    if m.FeatureByName(plane) is not None:
        m.ClearSelection2(True);assert m.FeatureByName(plane).Select2(False,0);m.BlankRefGeom()
    m.ClearSelection2(True);assert m.FeatureByName("弯臂腹板边界").Select2(False,0);m.BlankSketch()
    assert m.ForceRebuild3(False) and save_document(m,str(target))
    g(m,"ActiveView").DisplayMode=5
    m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
    save_preview(m,str(out/"holes.bmp"),"isometric")
    print(f"[孔腔完成] 本段{time.perf_counter()-timer:.2f}秒",flush=True)
finally:sw.SetUserPreferenceToggle(10,old);m.GraphicsRedraw2()
