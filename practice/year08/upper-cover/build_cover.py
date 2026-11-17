"""第8届阀上盖：COM准备复杂草图，主体回转/阵列/导出由MCP执行。"""
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time
import win32com.client

ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,sketch,sketch_line,save_document,dimension_existing_relations
from shared.single_part import relation,anchor,drive,quiet_dimensions
from sw_part import extrude_cut
from sw_review import save_preview

out=Path(sys.argv[1]).resolve();stage=sys.argv[2];target=out/"阀上盖_SDF-04.SLDPRT"
sw=win32com.client.Dispatch("SldWorks.Application.32");m=g(sw,"ActiveDoc")
assert g(sw,"GetProcessID")==int(sys.argv[3]) and m is not None
timer=time.perf_counter()


def select_floor(point):
    faces=[]
    for b in g(m,"GetBodies2",0,False):
        for f in g(b,"GetFaces"):
            s=g(f,"GetSurface")
            if not g(s,"IsPlane"):continue
            p=list(g(s,"PlaneParams"));box=list(g(f,"GetBox"))
            if abs(abs(p[1])-1)<1e-8 and abs(p[4]-.002)<1e-8 and all(box[i]-1e-8<=point[i]<=box[i+3]+1e-8 for i in range(3)):faces.append(f)
    assert len(faces)==1
    m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData");sd.X,sd.Y,sd.Z=point
    assert faces[0].Select4(False,sd)


def begin_floor(point):
    select_floor(point);m.SketchManager.InsertSketch(True)
    sk=g(m.SketchManager,"ActiveSketch")
    return sk,g(sk,"Name"),list(g(g(sk,"ModelToSketchTransform"),"ArrayData"))


def local(a,p):
    return [sum(p[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]


def checkpoint(name):
    assert m.ForceRebuild3(False) and save_document(m,str(target))
    m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
    (out/"checkpoint.json").write_text(json.dumps({"stage":name,"saved":datetime.now().astimezone().isoformat(),
        "target":target.name,"next":"按计划继续孔槽及最终验收"},ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"[{name}] 已保存，本段{time.perf_counter()-timer:.2f}秒",flush=True)


with quiet_dimensions(sw):
    if stage=="profile":
        assert not target.exists() and not g(m,"GetPathName")
        # 径向X、轴向Y，杯口Y0，外端Y18。中间Φ5.2与Φ7是两个不同轴段。
        pts=[(17,0),(20,0),(20,9.6),(16.5,18),(14.5,18),(14.5,17.5),
             (4.5,17.5),(4.5,12),(3.6,12),(2.6,11),(2.6,9),(3.5,9),(3.5,2),(17,2)]
        with sketch(m,"Front Plane") as ref:
            sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
            try:
                axis=sm.CreateCenterLine(0,0,0,0,.018,0)
                es=[sketch_line(m,a[0]/1000,a[1]/1000,b[0]/1000,b[1]/1000) for a,b in zip(pts,pts[1:]+pts[:1])]
            finally:sm.AddToDB=old
            for a,b in zip(es,es[1:]+es[:1]):relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
            for e,a,b in zip(es,pts,pts[1:]+pts[:1]):
                if a[0]==b[0] or a[1]==b[1]:relation(m,[e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
            anchor(m,g(axis,"GetStartPoint2"),x=True,y=True);relation(m,[axis],"sgVERTICAL")
            relation(m,[g(axis,"GetEndPoint2"),g(es[3],"GetStartPoint2")],"sgHORIZONTALPOINTS2D")
            anchor(m,g(es[0],"GetStartPoint2"),y=True)
            dimension_existing_relations(m)
            for e,p in zip(es,pts):
                q=g(e,"GetStartPoint2");assert math.dist([g(q,"X")*1000,g(q,"Y")*1000],p)<1e-5
        m.FeatureByName(str(ref)).Name="轴向截面"
        checkpoint("回转截面待MCP执行")
    elif stage=="slots":
        assert Path(g(m,"GetPathName")).resolve()==target
        # 长圆定位槽沿AA径向X；宽8、圆心距9，总长17，底位Y6.5。
        sk,ref,a=begin_floor((.006,.002,0))
        q0=local(a,(-.0045,.002,0));q1=local(a,(.0045,.002,0))
        from sw_part import sketch_slot
        # 当前封装参数名为radius，但传给原生API的是槽宽；回读应为端R4。
        sketch_slot(m,q0[0],q0[1],q1[0],q1[1],.008)
        dimension_existing_relations(m)
        inv=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"))
        m.SketchManager.InsertSketch(True);m.FeatureByName(ref).Name="8H9长圆定位槽"
        f=extrude_cut(m,"8H9长圆定位槽",.0045,direction=inv[7]>0)
        assert f is not None;f.Name="内芯扁位定位槽"
        # 90°弧槽，中心R14，端R2.5，底Y5.5（从杯底Y2切3.5）。
        sk,ref,a=begin_floor((.012,.002,0))
        def uv(r,deg):
            ang=math.radians(deg)
            # 图面水平u=Z、竖直v=X；杯内侧观察的上半槽。
            return local(a,(r*math.sin(ang)/1000,.002,r*math.cos(ang)/1000))[:2]
        def center(r,deg):return uv(r,deg)
        o=local(a,(0,.002,0))[:2];ex=local(a,(0,.002,.001))[:2];ey=local(a,(.001,.002,0))[:2]
        orient=1 if (ex[0]-o[0])*(ey[1]-o[1])-(ex[1]-o[1])*(ey[0]-o[0])>0 else -1
        spec=[(o,uv(16.5,45),uv(16.5,135),1),
              (center(14,135),uv(16.5,135),uv(11.5,135),1),
              (o,uv(11.5,135),uv(11.5,45),-1),
              (center(14,45),uv(11.5,45),uv(16.5,45),1)]
        sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
        try:arcs=[sm.CreateArc(c[0],c[1],0,p[0],p[1],0,q[0],q[1],0,d*orient) for c,p,q,d in spec]
        finally:sm.AddToDB=old
        for e,n in zip(arcs,arcs[1:]+arcs[:1]):relation(m,[g(e,"GetEndPoint2"),g(n,"GetStartPoint2")],"sgCOINCIDENT")
        relation(m,[arcs[0],arcs[2]],"sgCONCENTRIC")
        for e,n in zip(arcs,arcs[1:]+arcs[:1]):relation(m,[e,n],"sgTANGENT")
        # 原图90°用圆心径向构造线夹角驱动，不用换算XY尺寸替代。
        sm.AddToDB=True
        try:
            c0=g(arcs[3],"GetCenterPoint2");c1=g(arcs[1],"GetCenterPoint2")
            l0=sm.CreateCenterLine(0,0,0,g(c0,"X"),g(c0,"Y"),0)
            l1=sm.CreateCenterLine(0,0,0,g(c1,"X"),g(c1,"Y"),0)
        finally:sm.AddToDB=old
        for l,c in [(l0,c0),(l1,c1)]:
            anchor(m,g(l,"GetStartPoint2"),x=True,y=True)
            relation(m,[g(l,"GetEndPoint2"),c],"sgCOINCIDENT")
        drive(m,[l0,l1],(.020,.010,0),math.pi/2)
        drive(m,[arcs[1]],(.018,.012,0),.0025)
        dimension_existing_relations(m)
        m.SketchManager.InsertSketch(True);m.FeatureByName(ref).Name="90度弧槽轮廓"
        f=extrude_cut(m,"90度弧槽轮廓",.0035,direction=inv[7]>0)
        assert f is not None;f.Name="弧形限位槽"
        checkpoint("长圆定位槽与90度限位槽")
    elif stage=="blind":
        assert Path(g(m,"GetPathName")).resolve()==target
        select_floor((0,.002,-.014))
        # 本机GB_DATA_HW_DrillSizes.SIZE为“Ø6”，不是裸数字6。
        f=m.FeatureManager.HoleWizard5(2,13,355,"Ø6",0,.006,.011,-1,
            -1,math.pi,0,0,0,0,-1,-1,-1,-1,-1,-1,"",False,False,True,False,False,False)
        assert f is not None,"GB Φ6盲孔入口未创建，保留主体"
        f.Name="GB钢球弹簧母孔"
        sf=g(f,"GetFirstSubFeature");sf.Name="弹簧孔位";g(sf,"GetNextSubFeature").Name="弹簧孔原生轮廓"
        m.ClearSelection2(True);assert sf.Select2(False,0);m.EditSketch()
        sk=g(m.SketchManager,"ActiveSketch");p=list(g(sk,"GetSketchPoints2"))[0]
        anchor(m,p,x=abs(g(p,"X"))<1e-8,y=abs(g(p,"Y"))<1e-8)
        dimension_existing_relations(m);m.SketchManager.InsertSketch(True)
        checkpoint("GB六毫米弹簧母孔待阵列")
    elif stage=="mirror":
        assert Path(g(m,"GetPathName")).resolve()==target
        assert m.FeatureByName("两钢球弹簧孔镜像") is None
        m.ClearSelection2(True)
        assert m.FeatureByName("GB钢球弹簧母孔").Select2(False,1)
        assert m.FeatureByName("前视基准面").Select2(True,2)
        f=m.FeatureManager.InsertMirrorFeature2(False,False,False,False,0)
        assert f is not None;f.Name="两钢球弹簧孔镜像"
        checkpoint("两钢球弹簧孔镜像")
    elif stage=="inclined_axis":
        assert Path(g(m,"GetPathName")).resolve()==target
        # 在已消费的轴向截面补入构造定位，不增加悬空草图或重合平面。
        m.ClearSelection2(True);assert m.FeatureByName("轴向截面").Select2(False,0);m.EditSketch()
        sm=m.SketchManager;active=g(sm,"ActiveSketch")
        rot=next(e for e in g(active,"GetSketchSegments") if g(e,"GetType")==0 and g(e,"ConstructionGeometry") and abs(g(g(e,"GetEndPoint2"),"X"))<1e-8)
        angle=math.radians(22.6);s,c=math.sin(angle),math.cos(angle)
        distance=(28.4-16.5)/(c+5*s/12)
        q=(-28.4+c*distance,18-s*distance)
        old=sm.AddToDB;sm.AddToDB=True
        try:
            cone=sm.CreateCenterLine(-.020,.0096,0,-.0165,.018,0)
            normal=sm.CreateCenterLine(q[0]/1000,q[1]/1000,0,-.0284,.018,0)
            horizontal=sm.CreateCenterLine(-.0284,.018,0,0,.018,0)
        finally:sm.AddToDB=old
        relation(m,[g(horizontal,"GetEndPoint2"),g(rot,"GetEndPoint2")],"sgCOINCIDENT")
        relation(m,[g(horizontal,"GetStartPoint2"),g(normal,"GetEndPoint2")],"sgCOINCIDENT")
        relation(m,[horizontal],"sgHORIZONTAL")
        relation(m,[g(normal,"GetStartPoint2"),cone],"sgCOINCIDENT")
        relation(m,[g(cone,"GetEndPoint2"),g(rot,"GetEndPoint2")],"sgHORIZONTALPOINTS2D")
        drive(m,[horizontal],(-.015,.025,0),.0284)
        drive(m,[normal,horizontal],(-.025,.017,0),angle)
        drive(m,[g(cone,"GetStartPoint2"),g(rot,"GetStartPoint2")],(-.010,.005,0),.020,kind="horizontal")
        drive(m,[g(cone,"GetEndPoint2"),g(rot,"GetStartPoint2")],(-.008,.021,0),.0165,kind="horizontal")
        dimension_existing_relations(m)
        sm.InsertSketch(True)
        checkpoint("22点6度斜孔构造轴，后续使用finish_m8.py")
    else:raise ValueError(stage)
    if stage!="profile":save_preview(m,str(out/(stage+".bmp")),"isometric")
