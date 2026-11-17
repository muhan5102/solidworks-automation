"""LJT05.23：同值边成组，一次R10、一次R3、一次C1。"""
from pathlib import Path
import sys,json
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3];sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,round_edges
from sw_connect import connect_solidworks
from sw_part import fillet,chamfer
from sw_review import save_preview
sw,m=connect_solidworks(2024,.1);out=Path(sys.argv[1]).resolve();target=out/"支架_LJT05.23.SLDPRT"
assert Path(g(m,"GetPathName")).resolve()==target and g(sw,"GetProcessID")==42548


def edges():
    seen=set()
    for b in g(m,"GetBodies2",0,False):
        for face in g(b,"GetFaces"):
            for e in g(face,"GetEdges"):
                data=g(e,"GetCurveParams3");p=list(g(data,"StartPoint"));q=list(g(data,"EndPoint"))
                types=[]
                for af in g(e,"GetTwoAdjacentFaces2"):
                    if af is None:continue
                    su=g(af,"GetSurface")
                    for kind in ["Plane","Cylinder","Torus","Cone"]:
                        if g(su,"Is"+kind):types.append((kind,list(g(su,kind+"Params"))));break
                key=(tuple(sorted(tuple(round(v,8) for v in x) for x in [p,q])),str(sorted(types)))
                if key in seen:continue
                seen.add(key)
                yield e,p,q,types


if m.FeatureByName("安装板四角R10") is None:
    selected=[e for e,p,q,t in edges() if abs(p[1]-q[1])<1e-8 and abs(p[2]-q[2])<1e-8
              and abs(abs(p[2])-.040)<1e-8 and min(abs(p[1]-.075),abs(p[1]-.145))<1e-8 and abs(p[0]-q[0])>.011]
    assert len(selected)==4
    m.ClearSelection2(True);f=fillet(m,.010,selected);assert f is not None;f.Name="安装板四角R10"

if m.FeatureByName("同值铸造过渡R3") is None:
    selected=[];evidence=[]
    for e,p,q,types in edges():
        ribside=any(k=="Plane" and abs(abs(v[2])-1)<1e-8 and abs(abs(v[5])-.003)<1e-8 for k,v in types)
        support=False
        for kind,v in types:
            if kind=="Cylinder":
                # 图示T形截面两侧与翼缘、上下座的内交线。
                if abs(abs(v[4])-1)<1e-8 and min(abs(v[6]-.030),abs(v[6]-.0225))<1e-8:support=True
                if abs(abs(v[3])-1)<1e-8 and abs(v[1]-.110)<1e-8 and abs(v[6]-.015)<1e-8:support=True
                if abs(abs(v[5])-1)<1e-8 and abs(v[0]+.026)<1e-8 and abs(v[6]-.030)<1e-8:support=True
            if kind=="Plane":
                if abs(abs(v[0])-1)<1e-8 and abs(v[3]+.056)<1e-8:support=True
                if abs(abs(v[1])-1)<1e-8 and abs(v[4]-.040)<1e-8:support=True
        # 下座肩部和安装板-套筒根部的图示过渡。
        shoulder=any(k=="Cylinder" and min(abs(v[6]-.0225),abs(v[6]-.030))<1e-8 and abs(abs(v[4])-1)<1e-8 for k,v in types) and all(abs(x[1]-.040)<1e-8 for x in [p,q])
        plate_tube=any(k=="Cylinder" and abs(abs(v[3])-1)<1e-8 and abs(v[6]-.015)<1e-8 for k,v in types) \
            and any(k=="Plane" and abs(abs(v[0])-1)<1e-8 and abs(v[3]+.056)<1e-8 for k,v in types)
        plate_wing=all(abs(x[1]-.075)<1e-8 for x in [p,q]) and max(p[0],q[0])<=-.056+1e-8 \
            and min(p[0],q[0])>=-.064-1e-8 and max(abs(p[2]),abs(q[2]))<=.019+1e-8
        if ribside and support or shoulder or plate_tube or plate_wing:
            selected.append(e);evidence.append({"p":p,"q":q,"rib_root":ribside and support,"shoulder":shoulder,"plate_tube":plate_tube,"plate_wing":plate_wing})
    (out/"R3-selected.json").write_text(json.dumps(evidence,indent=2),encoding="utf-8")
    assert 12<=len(selected)<=20,len(selected)
    m.ClearSelection2(True);f=fillet(m,.003,selected);assert f is not None,"同值R3组合失败，需按实际依赖定位"
    f.Name="同值铸造过渡R3"
    print("[圆角] 同值R3一次多选",len(selected),"条边",flush=True)

selected=[e for e,p,r in round_edges(m) if abs(r-7.5)<1e-5 and abs(p[1]-110)<1e-5
          and min(abs(p[0]+73),abs(p[0]+35))<1e-5 or abs(r-13.5)<1e-5 and abs(p[1]-57)<1e-5]
assert len(selected)==3,len(selected)
m.ClearSelection2(True)
for edge in selected:assert edge.Select2(True,0)
f=chamfer(m,.001);assert f is not None;f.Name="三处孔口C1"
for name,value in [("名称","支架"),("图号","LJT05.23"),("材料","HT200"),("螺纹表达","4×M6-6H及M42×2-6H为装饰表达，无实体牙型")]:m.Extension.CustomPropertyManager("").Add3(name,30,value,2)
assert m.ForceRebuild3(False) and save_document(m,str(target))
g(m,"ActiveView").DisplayMode=5;m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
save_preview(m,str(out/"finished.bmp"),"isometric")
