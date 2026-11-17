"""仅替换未能稳定重开的M8孔向导，保留主体和两只弹簧孔。"""
import math
from pathlib import Path
import shutil
import sys
import win32com.client

sys.path.insert(0,str(Path(__file__).resolve().parents[3]/"practice"))
from shared.cad import member as g,sketch,sketch_line,save_document,dimension_existing_relations,round_edges
from shared.single_part import relation,drive,quiet_dimensions
from sw_part import _ensure_sketch_selected

out=Path(sys.argv[1]).resolve();pid=int(sys.argv[2]);target=out/"阀上盖_SDF-04.SLDPRT"
sw=win32com.client.Dispatch("SldWorks.Application.32");m=g(sw,"ActiveDoc")
assert g(sw,"GetProcessID")==pid and Path(g(m,"GetPathName")).resolve()==target
backup=out/"阀上盖_孔向导重开失败.SLDPRT";assert not backup.exists();shutil.copy2(target,backup)
failed=m.FeatureByName("GB斜向M8孔")
if failed is not None:
    m.ClearSelection2(True);assert failed.Select2(False,0)
    assert m.Extension.DeleteSelection2(1)
for name in ["M8斜孔入口面"]:
    f=m.FeatureByName(name)
    if f is not None:
        m.ClearSelection2(True);assert f.Select2(False,0);assert m.Extension.DeleteSelection2(0)
profile=g(m.FeatureByName("轴向截面"),"GetSpecificFeature2")
normal=next(e for e in g(profile,"GetSketchSegments") if g(e,"GetType")==0 and g(e,"ConstructionGeometry") and abs(g(g(e,"GetEndPoint2"),"X")+.0284)<1e-8 and abs(g(g(e,"GetStartPoint2"),"Y")-.018)>1e-5)
origin=g(normal,"GetStartPoint2");q=[g(origin,"X"),g(origin,"Y")]
a=math.radians(22.6);n=(-math.cos(a),math.sin(a));t=(math.sin(a),math.cos(a));r=.0034;depth=.007;tip=r/math.tan(math.radians(59))
pts=[q,[q[i]+r*t[i] for i in range(2)],[q[i]-depth*n[i]+r*t[i] for i in range(2)],
     [q[i]-(depth+tip)*n[i] for i in range(2)]]
with quiet_dimensions(sw):
    with sketch(m,"Front Plane") as ref:
        sm=m.SketchManager;old=sm.AddToDB;sm.AddToDB=True
        try:
            es=[sketch_line(m,*p,*v) for p,v in zip(pts,pts[1:]+pts[:1])]
            axis=sm.CreateCenterLine(*q,0,*pts[-1],0)
        finally:sm.AddToDB=old
        for e,next_e in zip(es,es[1:]+es[:1]):relation(m,[g(e,"GetEndPoint2"),g(next_e,"GetStartPoint2")],"sgCOINCIDENT")
        relation(m,[g(axis,"GetStartPoint2"),g(es[0],"GetStartPoint2")],"sgCOINCIDENT")
        relation(m,[g(axis,"GetEndPoint2"),g(es[2],"GetEndPoint2")],"sgCOINCIDENT")
        relation(m,[g(axis,"GetStartPoint2"),origin],"sgCOINCIDENT")
        relation(m,[axis,normal],"sgPARALLEL")
        relation(m,[es[0],axis],"sgPERPENDICULAR")
        relation(m,[es[1],axis],"sgPARALLEL")
        drive(m,[es[0]],(q[0]-.002,q[1]+.005,0),r)
        drive(m,[es[1]],(q[0]+.008,q[1]+.002,0),depth)
        dimension_existing_relations(m)
    m.FeatureByName(str(ref)).Name="M8定深孔与钻尖截面"
    _ensure_sketch_selected(m,"M8定深孔与钻尖截面")
    sd=g(m.SelectionManager,"CreateSelectData");sd.Mark=16;assert axis.Select4(True,sd)
    f=m.FeatureManager.FeatureRevolve2(True,True,False,True,False,False,0,0,2*math.pi,0,False,False,0,0,0,0,0,True,False,True)
    assert f is not None;f.Name="M8定深孔及118度钻尖"
    center=[(q[i]-depth*n[i])*1000 for i in range(2)]+[0]
    candidates=[edge for edge,p,radius in round_edges(m) if math.dist(p,center)<.001 and abs(radius-3.4)<.001]
    assert len(candidates)==1
    m.ClearSelection2(True);assert candidates[0].Select2(False,0)
    f=m.FeatureManager.InsertCosmeticThread3(13,"Tapped Hole","M8x1.25",.008,0,.007,"M8，深7，GB，22.6°斜向")
    assert f is not None;f.Name="M8装饰螺纹"
assert m.ForceRebuild3(False) and save_document(m,str(target))
print("已用一次回转切除形成定深7与118°钻尖，并保留原生装饰螺纹；待重开核验",flush=True)
