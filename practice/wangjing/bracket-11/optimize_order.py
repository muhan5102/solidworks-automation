"""当前副本：实心加料后统一开孔，两个普通切孔直接依附模型面。"""
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/"practice"));sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g,save_document,dimension_existing_relations,round_edges
from sw_connect import connect_solidworks
from sw_part import extrude_cut
from shared.single_part import circle,drive,anchor
sw,m=connect_solidworks(2024,.1)
out=ROOT/"outputs/practice/wangjing/bracket-11/20261117-r02";target=out/"支架_LJT05.11_指令优化.SLDPRT"
assert Path(g(m,"GetPathName")).resolve()==target and g(sw,"GetProcessID")==42548


def delete(name):
    m.ClearSelection2(True);f=m.FeatureByName(name);assert f is not None,name
    assert f.Select2(False,0) and m.Extension.DeleteSelection2(0),name


for name in ["M10装饰螺纹","M10攻丝底孔","M10攻丝底孔草图","主孔贯通","主孔贯通草图"]:delete(name)
for name in ["套筒端面","夹耳下表面"]:delete(name)
m.ClearSelection2(True);assert m.FeatureByName("套筒截面").Select2(False,0);m.EditSketch()
sk=g(m.SketchManager,"ActiveSketch")
inside=[seg for seg in g(sk,"GetSketchSegments") if abs(g(seg,"GetRadius")-.010)<1e-8]
assert len(inside)==1
m.ClearSelection2(True);assert inside[0].Select4(False,g(m.SelectionManager,"CreateSelectData"))
assert m.Extension.DeleteSelection2(0)
m.SketchManager.InsertSketch(True)
assert m.ForceRebuild3(False)
m.FeatureByName("夹紧套筒").Name="套筒实心毛坯"


def cut_on_face(axis,position,world_center,radius,depth,direction,name):
    candidates=[]
    for b in g(m,"GetBodies2",0,False):
        for face in g(b,"GetFaces"):
            su=g(face,"GetSurface")
            if not g(su,"IsPlane"):continue
            p=list(g(su,"PlaneParams"));box=list(g(face,"GetBox"))
            if abs(abs(p[axis])-1)<1e-8 and abs(p[axis+3]-position)<1e-8 and all(box[i]-1e-8<=world_center[i]<=box[i+3]+1e-8 for i in range(3)):
                candidates.append(face)
    assert len(candidates)==1,(name,len(candidates))
    m.ClearSelection2(True);sd=g(m.SelectionManager,"CreateSelectData");sd.X,sd.Y,sd.Z=world_center
    assert candidates[0].Select4(False,sd)
    m.SketchManager.InsertSketch(True)
    sk=g(m.SketchManager,"ActiveSketch");t=g(sk,"ModelToSketchTransform");a=list(g(t,"ArrayData"));inv=list(g(g(t,"Inverse"),"ArrayData"))
    q=[sum(world_center[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
    assert abs(q[2])<1e-8
    c=circle(m,q[0],q[1],radius)
    if abs(q[0])<1e-9 or abs(q[1])<1e-9:anchor(m,g(c,"GetCenterPoint2"),x=abs(q[0])<1e-9,y=abs(q[1])<1e-9)
    old=sw.GetUserPreferenceToggle(10);sw.SetUserPreferenceToggle(10,False)
    try:
        drive(m,[c],(q[0]+.02,q[1]+.02,0),2*radius,"diameter")
        dimension_existing_relations(m)
    finally:sw.SetUserPreferenceToggle(10,old)
    m.SketchManager.InsertSketch(True)
    # 从最后新增的ProfileFeature获取草图，避免依赖默认编号。
    f=g(m,"FirstFeature");profiles=[]
    while f is not None:
        if g(f,"GetTypeName2")=="ProfileFeature":profiles.append(f)
        f=g(f,"GetNextFeature")
    profiles[-1].Name=name+"草图"
    f=extrude_cut(m,name+"草图",depth,direction=sum(inv[6+i]*direction[i] for i in range(3))>0)
    assert f is not None,name
    f.Name=name


cut_on_face(2,.025,(0,0,.025),.010,0,(0,0,-1),"主孔贯通")
cut_on_face(1,-.009,(-.025,-.009,0),.00425,.0075,(0,1,0),"M10攻丝底孔")
matches=[e for e,p,r in round_edges(m) if abs(p[0]+25)<1e-5 and abs(p[1]+9)<1e-5 and abs(r-4.25)<1e-5]
assert len(matches)==1
m.ClearSelection2(True);assert matches[0].Select2(False,0)
f=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M10x1.5",.010,0,.0075,"M10-6H，通至夹紧槽")
assert f is not None;f.Name="M10装饰螺纹"
# 隐藏被Rib消费但本机仍显示的草图线，保留全部参数化节点。
for name in ["腹板边界"]:
    m.ClearSelection2(True);assert m.FeatureByName(name).Select2(False,0);m.BlankSketch()
for name,value in [("名称","支架"),("图号","LJT05.11"),("材料","HT200"),("螺纹表达","M10-6H，装饰螺纹，无实体牙型")]:
    m.Extension.CustomPropertyManager("").Add3(name,30,value,2)
assert m.ForceRebuild3(False) and save_document(m,str(target))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print("已采用实心套筒毛坯，主孔在连接加料后一次切出；本题额外基准面为0",flush=True)
