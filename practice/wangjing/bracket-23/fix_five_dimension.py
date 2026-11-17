"""5为R30与Φ45左侧理论交点相对顶面的落差，不是圆弧最低点。"""
from pathlib import Path
import sys,json,math,shutil
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3];sys.path.insert(0,str(ROOT/"practice"));sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,dimension_existing_relations
from sw_connect import connect_solidworks,create_empty_dispatch_variant
from shared.single_part import relation,anchor,drive,dimensions
sw,m=connect_solidworks(2024,.1);out=Path(sys.argv[1]).resolve();target=out/"支架_LJT05.23.SLDPRT"
assert Path(g(m,"GetPathName")).resolve()==target and g(sw,"GetProcessID")==42548
assert save_document(m,str(target));backup=out/"支架_LJT05.23_5尺寸修正前.SLDPRT";assert not backup.exists();shutil.copy2(target,backup)
for source,dest in [(out/"verification.json",out/"verification_5尺寸修正前.json"),(target.with_suffix(".step"),out/"支架_LJT05.23_5尺寸修正前.step")]:
    assert source.is_file() and not dest.exists();source.rename(dest)
shutil.copy2(out/"review.png",out/"review_5尺寸修正前.png")
old=sw.GetUserPreferenceToggle(10);sw.SetUserPreferenceToggle(10,False)
try:
    # 先置到推导位置，避免添加关系时任意移动已验证的其它端点。
    low=.052+math.sqrt(.030**2-(.0225-.008)**2)-.030
    m.Parameter("D3@弯臂腹板边界").SystemValue=low
    m.ClearSelection2(True);assert m.FeatureByName("弯臂腹板边界").Select2(False,0);m.EditSketch()
    sk=g(m.SketchManager,"ActiveSketch");arc=next(e for e in g(sk,"GetSketchSegments") if g(e,"GetType")==1)
    m.ClearSelection2(True);assert m.Extension.SelectByID2("D3@弯臂腹板边界","DIMENSION",0,0,0,False,0,create_empty_dispatch_variant(),0)
    assert m.Extension.DeleteSelection2(0)
    sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
    try:
        datum=sm.CreateCenterLine(-.040,.057,0,0,.057,0)
        point=sm.CreatePoint(-.0225,.052,0)
    finally:sm.AddToDB=db
    relation(m,[datum],"sgHORIZONTAL");anchor(m,g(datum,"GetEndPoint2"),x=True)
    relation(m,[point,arc],"sgCOINCIDENT")
    # 理论外壁X=-45/2，与原图5落差；Y57=110-53。
    m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData");assert point.Select4(False,sd)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0)
    dd=m.AddHorizontalDimension2(-.015,.042,0);assert dd is not None
    assert abs(g(g(dd,"GetDimension2",0),"SystemValue")-.0225)<1e-8
    drive(m,[point,datum],(-.026,.054,0),.005)
    dimension_existing_relations(m)
    assert abs(g(point,"X")+.0225)<1e-8 and abs(g(point,"Y")-.052)<1e-8
    center=g(arc,"GetCenterPoint2");assert abs(g(center,"Y")-(low+.030))<1e-8
    m.SketchManager.InsertSketch(True)
    assert m.ForceRebuild3(False) and save_document(m,str(target))
    (out/"five-dimension-repair.json").write_text(json.dumps({"arc_center_y_mm":g(center,"Y")*1000,
        "intersection_mm":[-22.5,52],"top_y_mm":57,"drawing_drop_mm":5,"driving_dimensions":dimensions(m,"弯臂腹板边界")},ensure_ascii=False,indent=2),encoding="utf-8")
finally:sw.SetUserPreferenceToggle(10,old);m.GraphicsRedraw2()
m.ClearSelection2(True);assert m.FeatureByName("弯臂腹板边界").Select2(False,0);m.BlankSketch();g(m,"ActiveView").DisplayMode=5;m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print("R30与Φ45理论交点已采用5毫米直接驱动，等待受影响特征复核",flush=True)
