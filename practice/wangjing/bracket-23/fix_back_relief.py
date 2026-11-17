"""按用户圈图补背面35高避空带，保留Φ28凸台；不猜槽底圆弧。"""
from pathlib import Path
import sys,shutil
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3];sys.path.insert(0,str(ROOT/"practice"));sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,dimension_existing_relations,sketch_line
from sw_connect import connect_solidworks
from sw_part import extrude_cut
from shared.single_part import relation,circle,drive,anchor
from sw_review import save_preview
sw,m=connect_solidworks(2024,.1);out=Path(sys.argv[1]).resolve();target=out/"支架_LJT05.23.SLDPRT"
assert Path(g(m,"GetPathName")).resolve()==target and g(sw,"GetProcessID")==42548
assert save_document(m,str(target))
backup=out/"支架_LJT05.23_避空及外侧圆角修正前.SLDPRT";assert not backup.exists();shutil.copy2(target,backup)
# 下游同值圆角稍后按原图补入外侧肩部，再统一生成；保留四角R10。
m.ClearSelection2(True);assert m.FeatureByName("同值铸造过渡R3").Select2(False,0);assert m.Extension.DeleteSelection2(0)
faces=[];world=(-.068,.110,.030)
for b in g(m,"GetBodies2",0,False):
    for face in g(b,"GetFaces"):
        su=g(face,"GetSurface")
        if not g(su,"IsPlane"):continue
        p=list(g(su,"PlaneParams"));box=list(g(face,"GetBox"))
        if abs(abs(p[0])-1)<1e-8 and abs(p[3]+.068)<1e-8 and all(box[i]-1e-8<=world[i]<=box[i+3]+1e-8 for i in range(3)):faces.append(face)
assert len(faces)==1
m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData");sd.X,sd.Y,sd.Z=world;assert faces[0].Select4(False,sd)
m.SketchManager.InsertSketch(True);sk=g(m.SketchManager,"ActiveSketch");t=list(g(g(sk,"ModelToSketchTransform"),"ArrayData"));inv=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"))
def local(p):return [sum(p[j]*t[j*3+i] for j in range(3))*t[12]+t[9+i] for i in range(3)]
points=[local(p) for p in [(-.068,.0925,-.045),(-.068,.1275,-.045),(-.068,.1275,.045),(-.068,.0925,.045)]]
assert all(abs(p[2])<1e-8 for p in points)
sm=m.SketchManager;db=sm.AddToDB;sm.AddToDB=True
try:es=[sketch_line(m,a[0],a[1],b[0],b[1]) for a,b in zip(points,points[1:]+points[:1])]
finally:sm.AddToDB=db
for a,b in zip(es,es[1:]+es[:1]):relation(m,[g(a,"GetEndPoint2"),g(b,"GetStartPoint2")],"sgCOINCIDENT")
for e,a,b in zip(es,points,points[1:]+points[:1]):relation(m,[e],"sgVERTICAL" if abs(a[0]-b[0])<1e-8 else "sgHORIZONTAL")
center=local((-.068,.110,0));c=circle(m,center[0],center[1],.014)
if abs(center[0])<1e-9 or abs(center[1])<1e-9:anchor(m,g(c,"GetCenterPoint2"),x=abs(center[0])<1e-9,y=abs(center[1])<1e-9)
old=sw.GetUserPreferenceToggle(10);sw.SetUserPreferenceToggle(10,False)
try:
    drive(m,[es[0]],(points[0][0]+.01,(points[0][1]+points[1][1])/2,0),.035)
    drive(m,[c],(center[0]+.02,center[1]+.02,0),.028,"diameter")
    dimension_existing_relations(m)
finally:sw.SetUserPreferenceToggle(10,old)
m.SketchManager.InsertSketch(True)
f=g(m,"FirstFeature");last=None
while f is not None:
    if g(f,"GetTypeName2")=="ProfileFeature":last=f
    f=g(f,"GetNextFeature")
last.Name="背面35高避空带轮廓"
f=extrude_cut(m,"背面35高避空带轮廓",.004,direction=inv[6]>0);assert f is not None;f.Name="背面避空保留28凸台"
assert m.ForceRebuild3(False) and save_document(m,str(target))
m.ClearSelection2(True);g(m,"ActiveView").DisplayMode=5;save_preview(m,str(out/"back-relief.bmp"),"left")
print("35高、4深的背面避空带已补，Φ28凸台保留；槽底圆弧待核实",flush=True)
