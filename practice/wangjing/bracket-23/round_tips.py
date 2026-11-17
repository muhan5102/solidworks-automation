"""按T形断面将6厚腹板与8厚翼缘自由端收圆。"""
from pathlib import Path
import sys,json,math
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3];sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document
from sw_connect import connect_solidworks
from sw_part import fillet
sw,m=connect_solidworks(2024,.1);out=Path(sys.argv[1]).resolve();target=out/"支架_LJT05.23.SLDPRT"
assert Path(g(m,"GetPathName")).resolve()==target and g(sw,"GetProcessID")==42548


def candidates(which):
    selected=[];evidence=[];seen=set()
    for b in g(m,"GetBodies2",0,False):
        for face in g(b,"GetFaces"):
            for e in g(face,"GetEdges"):
                types=[]
                for f in g(e,"GetTwoAdjacentFaces2"):
                    if f is None:continue
                    s=g(f,"GetSurface")
                    if g(s,"IsPlane"):types.append(("Plane",list(g(s,"PlaneParams"))))
                    elif g(s,"IsCylinder"):types.append(("Cylinder",list(g(s,"CylinderParams"))))
                side=any(k=="Plane" and abs(abs(v[2])-1)<1e-8 and abs(abs(v[5])-(.003 if which=="rib" else .019))<1e-8 for k,v in types)
                front=False
                for k,v in types:
                    if which=="rib":
                        if k=="Plane" and abs(abs(v[0])-1)<1e-8 and abs(v[3]+.038)<1e-8:front=True
                        if k=="Cylinder" and abs(abs(v[5])-1)<1e-8 and abs(v[0]+.008)<1e-8 and abs(v[1]-(.052+math.sqrt(.030**2-.0145**2)))<1e-8 and abs(v[6]-.030)<1e-8:front=True
                    else:
                        if k=="Plane" and abs(abs(v[0])-1)<1e-8 and min(abs(v[3]+.056),abs(v[3]+.064))<1e-8:front=True
                        if k=="Cylinder" and abs(abs(v[5])-1)<1e-8 and abs(v[0]+.026)<1e-8 and min(abs(v[6]-.030),abs(v[6]-.038))<1e-8:front=True
                if not(side and front):continue
                data=g(e,"GetCurveParams3");p=list(g(data,"StartPoint"));q=list(g(data,"EndPoint"));key=tuple(sorted([tuple(round(v,8) for v in p),tuple(round(v,8) for v in q)]))
                if key in seen:continue
                seen.add(key);selected.append(e);evidence.append({"p":p,"q":q,"surfaces":types})
    return selected,evidence


for which,radius,name,count in [("rib",.003,"腹板自由端R3",4),("wing",.004,"翼缘自由端R4",8)]:
    selected,evidence=candidates(which)
    (out/(which+"-tips.json")).write_text(json.dumps(evidence,indent=2),encoding="utf-8")
    print(name,"候选",len(selected),flush=True);assert len(selected)==count
    m.ClearSelection2(True);f=fillet(m,radius,selected)
    assert f is not None,(name,"半厚度圆头的恒定圆角未生成，停止继续收圆")
    f.Name=name
    assert m.ForceRebuild3(False) and save_document(m,str(target))
m.ClearSelection2(True);g(m,"ActiveView").DisplayMode=5;m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
