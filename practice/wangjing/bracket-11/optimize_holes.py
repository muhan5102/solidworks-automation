"""保留首轮，派生副本采用直接选面、孔向导及原生线性阵列。"""
from pathlib import Path
from datetime import datetime
import sys,json,shutil
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/"practice"))
sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,dimension_existing_relations
from sw_connect import connect_solidworks
from shared.single_part import anchor
from sw_part import linear_pattern
from sw_review import save_preview
sw,m=connect_solidworks(2024,.1)
assert g(sw,"GetProcessID")==42548
old=ROOT/"outputs/practice/wangjing/bracket-11/20261117-r01/支架_LJT05.11.SLDPRT"
out=ROOT/"outputs/practice/wangjing/bracket-11/20261117-r02"
resume="--resume" in sys.argv
if not resume: out.mkdir(exist_ok=False)
target=out/"支架_LJT05.11_指令优化.SLDPRT"
if not resume:
    assert Path(g(m,"GetPathName")).resolve()==old and g(m.SketchManager,"ActiveSketch") is None
    assert save_document(m,str(old))
    sw.CloseDoc(g(m,"GetTitle"))
    shutil.copy2(old,target)
    start=json.loads((old.parent/"start.json").read_text(encoding="utf-8"))
    start.update(repair_started=datetime.now().astimezone().isoformat(),reason="用户指出多余基准面、重复切孔、孔向导和重复特征指令未落实")
    (out/"start.json").write_text(json.dumps(start,ensure_ascii=False,indent=2),encoding="utf-8")
    import win32com.client
    e=win32com.client.VARIANT(16387,0);w=win32com.client.VARIANT(16387,0)
    m=sw.OpenDoc6(str(target),1,1,"",e,w)
    assert m is not None and e.value==0
else:
    assert Path(g(m,"GetPathName")).resolve()==target


def delete(name):
    f=m.FeatureByName(name)
    assert f is not None,name
    m.ClearSelection2(True)
    assert f.Select2(False,0) and m.Extension.DeleteSelection2(0),name


if not resume:
    # 拆除本轮已知孔特征；圆角不引用这些孔口。源版本不改。
    for name in ["两安装沉孔","两安装沉孔草图","两安装通孔","两安装通孔草图",
                 "夹耳沉孔","夹耳沉孔草图","夹耳间隙孔","夹耳间隙孔草图"]:
        delete(name)
    assert m.ForceRebuild3(False)


def select_plane_face(axis,position,point):
    candidates=[]
    for body in g(m,"GetBodies2",0,False):
        for face in g(body,"GetFaces"):
            su=g(face,"GetSurface")
            if not g(su,"IsPlane"):continue
            p=list(g(su,"PlaneParams"));box=list(g(face,"GetBox"))
            if abs(abs(p[axis])-1)<1e-8 and abs(p[3+axis]-position)<1e-8:
                if all(box[i]-1e-8<=point[i]<=box[i+3]+1e-8 for i in range(3)):
                    candidates.append(face)
    assert len(candidates)==1,(axis,position,len(candidates))
    # 原生面选择附带实际孔中心，不建立重合基准面。
    m.ClearSelection2(True)
    sd=g(m.SelectionManager,"CreateSelectData")
    sd.X,sd.Y,sd.Z=point
    assert candidates[0].Select4(False,sd)


def wizard(axis,position,point,diameter,seat,depth,end,hole_depth,size,name):
    select_plane_face(axis,position,point)
    f=m.FeatureManager.HoleWizard5(0,8,139,size,end,diameter,hole_depth,-1,
        seat,depth,0,-1,-1,0,0,0,0,0,0,-1,"",False,False,True,False,False,False)
    assert f is not None,name
    f.Name=name
    sf=g(f,"GetFirstSubFeature")
    sf.Name=name+"孔位"
    g(sf,"GetNextSubFeature").Name=name+"原生轮廓"
    m.ClearSelection2(True);assert sf.Select2(False,0);m.EditSketch()
    sk=g(m.SketchManager,"ActiveSketch")
    points=list(g(sk,"GetSketchPoints2") or [])
    assert len(points)==1
    p=points[0]
    if abs(g(p,"X"))<1e-9 or abs(g(p,"Y"))<1e-9:
        anchor(m,p,x=abs(g(p,"X"))<1e-9,y=abs(g(p,"Y"))<1e-9)
    dimension_existing_relations(m)
    # 使用草图变换独立核对孔位的世界坐标。
    a=list(g(g(g(sk,"ModelToSketchTransform"),"Inverse"),"ArrayData"))
    local=[g(p,"X"),g(p,"Y"),g(p,"Z")]
    world=[sum(local[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
    assert max(abs(world[i]-point[i]) for i in range(3))<1e-7,(name,world,point)
    m.SketchManager.InsertSketch(True)
    d=g(f,"GetDefinition")
    assert abs(g(d,"CounterBoreDiameter")-seat)<1e-8
    # SW2024该孔型HoleDiameter回读为0；以实际圆柱面核对通孔径。
    radii=[g(g(face,"GetSurface"),"CylinderParams")[6] for body in g(m,"GetBodies2",0,False) for face in g(body,"GetFaces") if g(g(face,"GetSurface"),"IsCylinder")]
    assert any(abs(r-diameter/2)<1e-8 for r in radii)
    assert abs(g(d,"CounterBoreDepth")-depth)<1e-8
    return f


if m.FeatureByName("底座沉孔母孔") is None:
    wizard(0,.060,(.060,-.100,-.020),.015,.028,.003,1,.016,"M12","底座沉孔母孔")
pattern=linear_pattern(m,"底座沉孔母孔",0,0,1,.040,2)
assert pattern is not None
pattern.Name="两组沉孔40阵列"
print("底座孔向导母孔与40毫米线性阵列已创建",flush=True)
assert save_document(m,str(target))
wizard(1,.009,(-.025,.009,0),.011,.018,.003,0,.0075,"M10","夹耳沉孔向导")
assert m.ForceRebuild3(False)
# 替换孔之后，已无引用的重合基准面移除；其余面由后续局部改序处理。
for name in ["底座背面","夹耳上表面"]:delete(name)
assert save_document(m,str(target))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
save_preview(m,str(out/"wizard.bmp"),"isometric")
(out/"checkpoint.json").write_text(json.dumps({"saved":datetime.now().astimezone().isoformat(),"target":target.name,
    "stage":"两组孔向导及底座原生线性阵列","next":"实体参数、孔位与下游关联复查"},ensure_ascii=False,indent=2),encoding="utf-8")
