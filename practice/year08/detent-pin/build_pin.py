"""第8届挡轴，球头与端倒角合并在一个回转截面。"""
import sys,json,math,time
from pathlib import Path
from datetime import datetime
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from shared.cad import member,mm,sketch,sketch_line,revolve_boss,dimension_existing_relations,new_document,save_document
from sw_connect import open_document,create_empty_dispatch_variant
from sw_review import save_preview
from sw_sketch_quality import inspect_model_sketches
import win32com.client
started=time.monotonic()
out=Path(sys.argv[1]).resolve()
filename=sys.argv[2] if len(sys.argv)>2 else "SDF-07-detent-pin.SLDPRT"
assert Path(filename).name==filename and Path(filename).suffix.lower()==".sldprt"
dest=out/filename
assert out.is_dir() and not dest.exists()
sw=win32com.client.Dispatch("SldWorks.Application.32");sw.Visible=True;sw.UserControl=True
pid=int(member(sw,"GetProcessID"));assert pid==37208
model=new_document(sw,"part")
with sketch(model,"Front Plane") as ref:
 sm=model.SketchManager;old=sm.AddToDB;sm.AddToDB=True
 try:
  axis=sm.CreateCenterLine(0,0,0,0,.007,0)
  arc=sm.CreateArc(0,.002,0,0,0,0,.002,.002,0,1)
  pts=[(2,2),(2,6.7),(1.7,7),(0,7),(0,0)]
  lines=[sketch_line(model,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1])) for a,b in zip(pts,pts[1:])]
 finally:sm.AddToDB=old
 sd=member(model.SelectionManager,"CreateSelectData");sd.Mark=0
 entities=[arc]+lines
 for a,b in zip(entities,entities[1:]+entities[:1]):
  model.ClearSelection2(True);assert member(a,"GetEndPoint2").Select4(False,sd);assert member(b,"GetStartPoint2").Select4(True,sd);model.SketchAddConstraints("sgCOINCIDENT")
 for segment,relation in [(lines[0],"sgVERTICAL"),(lines[2],"sgHORIZONTAL"),(lines[3],"sgVERTICAL"),(axis,"sgVERTICAL")]:
  model.ClearSelection2(True);segment.Select4(False,sd);model.SketchAddConstraints(relation)
 for point in (member(axis,"GetStartPoint2"),member(arc,"GetStartPoint2")):
  model.ClearSelection2(True);assert point.Select4(False,sd);assert model.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,create_empty_dispatch_variant(),0);model.SketchAddConstraints("sgCOINCIDENT")
 model.ClearSelection2(True);arc.Select4(False,sd);lines[0].Select4(True,sd);model.SketchAddConstraints("sgTANGENT")
 quality=dimension_existing_relations(model)
 contours=list(member(member(sm,"ActiveSketch"),"GetSketchContours") or []);assert contours and all(member(c,"IsClosed") for c in contours)
feature=revolve_boss(model,ref,2*math.pi);assert feature is not None
feature.Name="主体回转";member(model,"FeatureByName",str(ref)).Name="轴向截面"
model.SetMaterialPropertyName2("","","")
for k,v in [("名称","挡轴"),("图号","SDF-07"),("材料","45"),("技术要求","未注倒角C0.3；表面镀锌"),("尺寸","Φ4 g9公称建模；总长7；球头SR2")]:model.Extension.CustomPropertyManager("").Add3(k,30,v,2)
assert model.ForceRebuild3(False);assert save_document(model,str(dest))
sw.CloseDoc(member(model,"GetTitle"));model=open_document(sw,str(dest),silent=True);assert model.ForceRebuild3(False)
faces=[]
for body in member(model,"GetBodies2",0,False):
 for face in member(body,"GetFaces"):
  surf=member(face,"GetSurface")
  if member(surf,"IsSphere"):faces.append({"sphere":list(member(surf,"SphereParams"))})
  if member(surf,"IsCylinder"):faces.append({"cylinder":list(member(surf,"CylinderParams"))})
quality=inspect_model_sketches(model)
checks={"one_solid":len(member(model,"GetBodies2",0,False))==1,"fully_defined":quality["status"]=="pass","SR2":any("sphere" in f and abs(f["sphere"][3]-.002)<1e-8 for f in faces),"D4":any("cylinder" in f and abs(f["cylinder"][6]-.002)<1e-8 for f in faces)}
from shared.cad import round_edges
edges=[(position,radius) for _,position,radius in round_edges(model)]
def has_edge(y,radius):
 return any(math.dist(position,[0,y,0])<1e-5 and abs(r-radius)<1e-5 for position,r in edges)
sphere=next((f["sphere"] for f in faces if "sphere" in f),None)
checks["length_7"]=sphere is not None and abs(sphere[1]-.002)<1e-8 and has_edge(7,1.7)
checks["C0p3"]=has_edge(6.7,2) and has_edge(7,1.7)
assert all(checks.values()),checks
model.ShowNamedView2("*Isometric",7);model.ViewZoomtofit2();model.GraphicsRedraw2()
save_preview(model,str(out/"挡轴_isometric.bmp"),"isometric")
report={"checks":checks,"surfaces":faces,"sketches":quality,"end":datetime.now().astimezone().isoformat(),"execution_seconds":round(time.monotonic()-started,3),"material":"45仅属性，未配置物性","thread":"不适用"}
(out/"verification.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(checks,ensure_ascii=False))
