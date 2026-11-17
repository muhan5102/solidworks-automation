"""LJT05.11局部修正、按图过渡与装饰螺纹。"""
from pathlib import Path
import sys, math
sys.stdout.reconfigure(encoding="utf-8")
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/"practice"))
from shared.cad import member as g, dimension_existing_relations, round_edges, save_document
from sw_connect import connect_solidworks
from sw_part import fillet
from sw_edge_select import _edge_curve
from sw_review import save_preview
sys.path.insert(0,str(ROOT/"practice"))
from shared.single_part import dimensions
out=Path(sys.argv[1]).resolve()
sw,m=connect_solidworks(2024,.1)
target=out/"支架_LJT05.11.SLDPRT"
assert g(sw,"GetProcessID")==42548 and Path(g(m,"GetPathName")).resolve()==target
if sys.argv[2]=="profile":
    m.ClearSelection2(True)
    assert m.FeatureByName("腹板边界").Select2(False,0)
    m.EditSketch()
    sk=g(m.SketchManager,"ActiveSketch")
    rm=g(sk,"RelationManager")
    # 本草图自动添加的两个外部逻辑关系改变了预期端点；保留两条驱动尺寸。
    removed=[]
    for rel in list(g(rm,"GetRelations",0)):
        if g(rel,"GetRelationType")==9:
            assert rm.DeleteRelation(rel)
            removed.append(9)
    assert len(removed)==2
    m.Parameter("D1@腹板边界").SystemValue=.007
    dimension_existing_relations(m)
    dims=dimensions(m,"腹板边界")
    # 新增端点高度驱动替换为由R17.5和图示7推得的交点高度。
    for row in dims:
        if abs(row["value"]-.00851305318105438)<1e-7:
            m.Parameter(row["name"].split("@支架")[0]).SystemValue=math.sqrt(.0175**2-.007**2)
    m.SketchManager.InsertSketch(True)
    assert m.ForceRebuild3(False)
    seg=list(g(g(m.FeatureByName("腹板边界"),"GetSpecificFeature2"),"GetSketchSegments"))[0]
    points=[[g(g(seg,k),"X"),g(g(seg,k),"Y")] for k in ["GetStartPoint2","GetEndPoint2"]]
    assert abs(points[0][0]+.007)<1e-8 and abs(points[0][1]+math.sqrt(.0175**2-.007**2))<1e-8,points
    assert abs(points[1][0]-.044)<1e-8 and abs(points[1][1]+.115)<1e-8,points
    print("腹板端点已恢复原图7及底部5定位",points)
elif sys.argv[2]=="finish":
    matches=[e for e,p,r in round_edges(m) if abs(p[0]+25)<1e-5 and abs(p[1]+9)<1e-5 and abs(r-4.25)<1e-5]
    assert len(matches)==1
    m.ClearSelection2(True)
    assert matches[0].Select2(False,0)
    f=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M10x1.5",.010,0,.0075,"M10-6H，通至夹紧槽")
    assert f is not None
    f.Name="M10装饰螺纹"

    def edge_list():
        seen=set()
        for body in g(m,"GetBodies2",0,False):
            for face in g(body,"GetFaces"):
                for e in g(face,"GetEdges"):
                    if not g(_edge_curve(e),"IsLine"):continue
                    p,q=g(e,"GetStartVertex"),g(e,"GetEndVertex")
                    if p is None or q is None:continue
                    p,q=[v*1000 for v in g(p,"GetPoint")],[v*1000 for v in g(q,"GetPoint")]
                    key=tuple(sorted([tuple(round(v,6) for v in p),tuple(round(v,6) for v in q)]))
                    if key in seen:continue
                    seen.add(key)
                    yield e,p,q

    # 图示底座正视四角，沿X方向；不选择孔口或其它凸边。
    selected=[e for e,p,q in edge_list() if abs(p[1]-q[1])<1e-6 and abs(p[2]-q[2])<1e-6
              and abs(abs(p[2])-41)<1e-6 and min(abs(p[1]+120),abs(p[1]+70))<1e-6]
    assert len(selected)==4
    m.ClearSelection2(True)
    f=fillet(m,.002,selected)
    assert f is not None
    f.Name="底座四角R2"
    # T形剖面中腹板两侧与翼板的两条纵向内交线。
    selected=[e for e,p,q in edge_list() if abs(abs(p[2])-4)<1e-6 and abs(p[2]-q[2])<1e-6
              and max(p[1],q[1])>-10 and min(p[1],q[1])<-65 and min(p[0],q[0])>10]
    assert len(selected)==2
    m.ClearSelection2(True)
    f=fillet(m,.002,selected)
    assert f is not None
    f.Name="腹板翼板过渡R2"
    # 腹板接底座的两条竖向内交线，以及翼板落在底座上的两段。
    selected=[e for e,p,q in edge_list() if abs(abs(p[2])-4)<1e-6 and abs(p[2]-q[2])<1e-6
              and abs(p[0]-44)<1e-6 and abs(q[0]-44)<1e-6 and abs(p[1]-q[1])>30]
    assert len(selected)==2
    m.ClearSelection2(True)
    f=fillet(m,.002,selected)
    assert f is not None
    f.Name="腹板底座过渡R2"
    for name in ["套筒端面","底座背面","夹耳上表面","夹耳下表面"]:
        m.ClearSelection2(True)
        assert m.FeatureByName(name).Select2(False,0)
        m.BlankRefGeom()
else:raise ValueError(sys.argv[2])
assert m.ForceRebuild3(False)
assert save_document(m,str(target))
m.ClearSelection2(True)
m.ShowNamedView2("*Isometric",7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
save_preview(m,str(out/(sys.argv[2]+".bmp")),"isometric")
