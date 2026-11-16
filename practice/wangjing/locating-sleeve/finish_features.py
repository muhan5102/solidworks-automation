"""主体回转后补母孔/削平、约束与装饰螺纹；圆周阵列单独由MCP调用。"""
from pathlib import Path
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/"year08/connector"))
from build_connector import (member as g, mm, circle_cut, sketch, sketch_line,
                             dimension_existing_relations, save_document, round_edges, offset_plane)
from sw_part import extrude_cut, _get_plane_name_candidates, _select_first_candidate
import win32com.client

stage=sys.argv[1]
out=Path(sys.argv[2]).resolve()
s=win32com.client.Dispatch("SldWorks.Application.32")
assert g(s,"GetProcessID")==34092
m=g(s,"ActiveDoc")
assert Path(g(m,"GetPathName")).resolve()==out/"定位套_LJT03.28.SLDPRT"
if stage=="mother":
    f=m.FeatureByName("旋转1")
    if f is not None:f.Name="主体回转"
    offset_plane(m,"Top Plane",15,"法兰前端面")
    if m.FeatureByName("回转轴") is None:
        m.ClearSelection2(True)
        assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Front Plane"),"PLANE")
        assert _select_first_candidate(m.Extension,_get_plane_name_candidates("Right Plane"),"PLANE",append=True)
        assert m.InsertAxis2(True)
        f=g(m,"FirstFeature")
        while f is not None:
            if g(f,"GetTypeName2")=="RefAxis":f.Name="回转轴"
            f=g(f,"GetNextFeature")
    # 本轮失败MCP残留，删除前逐一核对名称及当前任务文档。
    for name in ["草图4","安装沉孔_通孔","草图3"]:
        f=m.FeatureByName(name)
        if f is not None:
            assert f.Select2(False,0)
            assert m.Extension.DeleteSelection2(1 if name=="安装沉孔_通孔" else 0)
    circle_cut(m,s,"法兰前端面",[(0,15,78)],6.25,33,1,"安装通孔")
    circle_cut(m,s,"法兰前端面",[(0,15,78)],10,8,1,"安装沉孔")
elif stage=="finish":
    # 削平只切法兰：在Y15端面按变换建立矩形，沿轴向切33。
    corners=[(-110,15,-110),(110,15,-110),(110,15,-70),(-110,15,-70)]
    with sketch(m,"法兰前端面") as ref:
        active=g(m.SketchManager,"ActiveSketch")
        matrix=list(g(g(active,"ModelToSketchTransform"),"ArrayData"))
        inv=list(g(g(g(active,"ModelToSketchTransform"),"Inverse"),"ArrayData"))
        points=[]
        for c in corners:
            p=[mm(v) for v in c]
            xyz=[sum(p[j]*matrix[j*3+i] for j in range(3))*matrix[12]+matrix[9+i] for i in range(3)]
            assert abs(xyz[2])<1e-8
            points.append(xyz)
        # AddToDB + 显式约束防止偶然自动关系改变矩形。
        old=m.SketchManager.AddToDB
        m.SketchManager.AddToDB=True
        try:
            edges=[sketch_line(m,*a[:2],*b[:2]) for a,b in zip(points,points[1:]+points[:1])]
        finally:m.SketchManager.AddToDB=old
        sd=g(m.SelectionManager,"CreateSelectData")
        for a,b in zip(edges,edges[1:]+edges[:1]):
            m.ClearSelection2(True)
            assert g(a,"GetEndPoint2").Select4(False,sd) and g(b,"GetStartPoint2").Select4(True,sd)
            m.SketchAddConstraints("sgCOINCIDENT")
        for e,a,b in zip(edges,points,points[1:]+points[:1]):
            m.ClearSelection2(True);assert e.Select4(False,sd)
            m.SketchAddConstraints("sgVERTICAL" if abs(a[0]-b[0])<1e-10 else "sgHORIZONTAL")
        dimension_existing_relations(m)
    f=extrude_cut(m,ref,mm(33),direction=inv[7]>0)
    assert f is not None
    f.Name="法兰削平"
    m.FeatureByName(str(ref)).Name="削平轮廓"
    # 端面至环槽右侧：30-8=22；实际圆柱2mm入口倒角后余20mm。
    minor=(80-1.0825317547305484*2)/2
    matches=[e for e,p,r in round_edges(m) if abs(p[1]-188)<1e-5 and abs(r-minor)<1e-5]
    assert len(matches)==1
    m.ClearSelection2(True);assert matches[0].Select2(False,0)
    f=m.FeatureManager.InsertCosmeticThread3(8,"Tapped Hole","M80x2",mm(80),0,mm(20),"M80×2-6H，端面至退刀槽22")
    assert f is not None
    f.Name="M80内螺纹"
else:
    raise ValueError(stage)
assert m.ForceRebuild3(False)
assert save_document(m,str(out/"定位套_LJT03.28.SLDPRT"))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
print(f"[{stage}] 已重建、保存并刷新")
