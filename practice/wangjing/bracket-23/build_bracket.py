"""LJT05.23：先主体加料，孔腔随后；按自然阶段串行。"""
from pathlib import Path
from datetime import datetime
import sys,json,math,time
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/"practice"));sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,dimension_existing_relations,new_document,sketch,sketch_line
from sw_connect import connect_solidworks
from sw_part import extrude_midplane,revolve_boss,sketch_arc,_ensure_sketch_selected
from shared.single_part import relation,anchor,drive
from sw_review import save_preview
out=Path(sys.argv[1]).resolve();stage=sys.argv[2];target=out/"支架_LJT05.23.SLDPRT"
sw,m=connect_solidworks(2024,.1)
assert g(sw,"GetProcessID")==42548
if stage=="body":
    assert not target.exists() and g(sw,"GetFirstDocument") is None
    m=new_document(sw,"part")
else:assert Path(g(m,"GetPathName")).resolve()==target
timer=time.perf_counter()


def poly(points):
    sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
    try:edges=[sketch_line(m,a[0]/1000,a[1]/1000,b[0]/1000,b[1]/1000) for a,b in zip(points,points[1:]+points[:1])]
    finally:sm.AddToDB=old
    for a,b in zip(edges,edges[1:]+edges[:1]):relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
    for e,a,b in zip(edges,points,points[1:]+points[:1]):
        if a[0]==b[0] or a[1]==b[1]:relation(m,[e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
    zerosx=[g(e,"GetStartPoint2") for e,a in zip(edges,points) if a[0]==0]
    zerosy=[g(e,"GetStartPoint2") for e,a in zip(edges,points) if a[1]==0]
    if zerosx:anchor(m,zerosx[0],x=True)
    if zerosy:anchor(m,zerosy[0],y=True)
    return edges


def checkpoint(name):
    assert m.ForceRebuild3(False) and save_document(m,str(target))
    m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
    (out/"checkpoint.json").write_text(json.dumps({"stage":name,"saved":datetime.now().astimezone().isoformat(),"target":target.name,
        "next":"孔腔与统一收尾"},ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"[{name}] 已保存，本段{time.perf_counter()-timer:.2f}秒",flush=True)


old=sw.GetUserPreferenceToggle(10);sw.SetUserPreferenceToggle(10,False)
try:
    if stage in ["body","wing","rib"]:
        if stage=="body":
            # 下座竖直轴Y，底面Y0；先实心，后统一回转切内腔。
            with sketch(m,"Front Plane") as ref:
                pts=[(0,0),(30,0),(30,40),(22.5,40),(22.5,57),(0,57)]
                edges=poly(pts)
                sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
                try:axis=sm.CreateCenterLine(0,0,0,0,.057,0)
                finally:sm.AddToDB=db
                relation(m,[g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgCOINCIDENT")
                relation(m,[g(axis,"GetEndPoint2"),g(edges[5],"GetStartPoint2")],"sgCOINCIDENT")
                dimension_existing_relations(m)
            m.FeatureByName(str(ref)).Name="下座外形截面"
            f=revolve_boss(m,"下座外形截面",2*math.pi);assert f is not None;f.Name="下座实心回转"
            with sketch(m,"Front Plane") as ref:
                poly([(-68,75),(-56,75),(-56,145),(-68,145)])
                dimension_existing_relations(m)
            m.FeatureByName(str(ref)).Name="安装板截面"
            f=extrude_midplane(m,"安装板截面",.080);assert f is not None;f.Name="安装板"
            with sketch(m,"Front Plane") as ref:
                edges=poly([(-73,110),(-73,124),(-68,124),(-68,125),(-35,125),(-35,110)])
                sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
                try:axis=sm.CreateCenterLine(-.073,.110,0,-.035,.110,0)
                finally:sm.AddToDB=db
                relation(m,[g(axis,"GetStartPoint2"),g(edges[0],"GetStartPoint2")],"sgCOINCIDENT")
                relation(m,[g(axis,"GetEndPoint2"),g(edges[5],"GetStartPoint2")],"sgCOINCIDENT")
                dimension_existing_relations(m)
            m.FeatureByName(str(ref)).Name="上套筒外形截面"
            f=revolve_boss(m,"上套筒外形截面",2*math.pi);assert f is not None;f.Name="上套筒实心回转"
        # 弯臂T形翼缘：内R30、厚8、宽38；上端相对板背面偏4。
        if stage=="wing":
            for name in ["弯臂翼缘","弯臂翼缘截面"]:
                m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,0);assert m.Extension.DeleteSelection2(0)
        cy=12+math.sqrt(38**2-4**2)
        inner_y=cy-math.sqrt(30**2-6**2)
        outer_y=cy-math.sqrt(38**2-6**2)
        if stage!="rib":
            with sketch(m,"Front Plane") as ref:
                sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
                try:
                    es=[sketch_line(m,-.064,.095,-.056,.095),sketch_line(m,-.056,.095,-.056,cy/1000),
                        sketch_arc(m,-.026,cy/1000,-.056,cy/1000,-.020,inner_y/1000,1),
                        sketch_line(m,-.020,inner_y/1000,-.020,outer_y/1000),
                        sketch_arc(m,-.026,cy/1000,-.020,outer_y/1000,-.064,cy/1000,-1),
                        sketch_line(m,-.064,cy/1000,-.064,.095)]
                finally:sm.AddToDB=db
                for a,b in zip(es,es[1:]+es[:1]):relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
                relation(m,[es[0]],"sgHORIZONTAL")
                for i in [1,3,5]:relation(m,[es[i]],"sgVERTICAL")
                relation(m,[es[1],es[2]],"sgTANGENT")
                relation(m,[es[4],es[5]],"sgTANGENT")
                relation(m,[es[2],es[4]],"sgCONCENTRIC")
                drive(m,[es[0]],(-.062,.100,0),.008)
                drive(m,[es[2]],(-.050,.026,0),.030)
                dimension_existing_relations(m)
                for arc,r in [(es[2],.030),(es[4],.038)]:assert abs(g(arc,"GetRadius")-r)<1e-8
            m.FeatureByName(str(ref)).Name="弯臂翼缘截面"
            f=extrude_midplane(m,"弯臂翼缘截面",.038);assert f is not None;f.Name="弯臂翼缘"
            checkpoint("主要加料阶段")
        if stage=="body":
            with sketch(m,"Front Plane") as ref:
                sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
                try:
                    e=sketch_line(m,-.038,.095,-.038,(.052+math.sqrt(.030**2-.0145**2)))
                    a=sketch_arc(m,-.008,(.052+math.sqrt(.030**2-.0145**2)),-.038,(.052+math.sqrt(.030**2-.0145**2)),-.008,(.052+math.sqrt(.030**2-.0145**2)-.030),1)
                finally:sm.AddToDB=db
                relation(m,[g(e,"GetEndPoint2"),g(a,"GetStartPoint2")],"sgCOINCIDENT")
                relation(m,[e],"sgVERTICAL");relation(m,[e,a],"sgTANGENT")
                drive(m,[a],(-.018,.068,0),.030)
                relation(m,[g(a,"GetCenterPoint2"),g(a,"GetEndPoint2")],"sgVERTICALPOINTS2D")
                dimension_existing_relations(m)
        else:
            ref="草图5"
            if g(m.SketchManager,"ActiveSketch") is None:
                m.ClearSelection2(True);assert m.FeatureByName(ref).Select2(False,0);m.EditSketch()
            sk=g(m.SketchManager,"ActiveSketch");a=next(seg for seg in g(sk,"GetSketchSegments") if g(seg,"GetType")==1)
            relation(m,[g(a,"GetCenterPoint2"),g(a,"GetEndPoint2")],"sgVERTICALPOINTS2D")
            dimension_existing_relations(m);m.SketchManager.InsertSketch(True)
        m.FeatureByName(str(ref)).Name="弯臂腹板边界"
        _ensure_sketch_selected(m,"弯臂腹板边界")
        m.FeatureManager.InsertRib(True,False,.006,0,False,False,False,0.,False,False)
        f=g(m,"FirstFeature");ribs=[]
        while f is not None:
            if g(f,"GetTypeName2")=="Rib":ribs.append(f)
            f=g(f,"GetNextFeature")
        assert len(ribs)==1,"本题原生Rib未创建，主体保存点保留"
        ribs[0].Name="弯臂腹板"
        assert len(g(m,"GetBodies2",0,False))==1
        checkpoint("全部加料完成")
        save_preview(m,str(out/"body.bmp"),"isometric")
    else:raise ValueError(stage)
finally:
    sw.SetUserPreferenceToggle(10,old);m.GraphicsRedraw2()
