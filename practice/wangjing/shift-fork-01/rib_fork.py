"""本题例外：开放曲面放样限定筋顶，拉伸到曲面。

依据用户提供视频BV1snZYYDE4U；其他零件仍优先原生Rib。
"""
import math
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.single_part import g, relation, quiet_dimensions
from shared.cad import sketch, sketch_line, dimension_existing_relations, offset_plane
from sw_part import _ensure_sketch_selected


def create_rib(sw, model):
    assert model.FeatureByName("加强筋") is None
    bottom = offset_plane(model, "Top Plane", 8, "筋脚平面")
    top = offset_plane(model, "Top Plane", 28, "筋顶平面")
    root = .120-math.sqrt(.015**2-.003**2)
    with quiet_dimensions(sw):
        with sketch(model, bottom) as ref:
            line = sketch_line(model, .070, -.003, .070, .003)
            relation(model, [line], "sgVERTICAL")
            dimension_existing_relations(model)
        model.FeatureByName(str(ref)).Name = "筋脚宽度"
        with sketch(model, top) as ref:
            sm = model.SketchManager
            old = sm.AddToDB
            sm.AddToDB = True
            try:
                arc = sm.CreateArc(.120, 0., 0., root, .003, 0.,
                                   root, -.003, 0., 1)
                assert arc is not None
            finally:
                sm.AddToDB = old
            relation(model, [g(arc, "GetStartPoint2"),
                             g(arc, "GetEndPoint2")], "sgVERTICALPOINTS2D")
            dimension_existing_relations(model)
        model.FeatureByName(str(ref)).Name = "筋顶圆弧"
    model.ClearSelection2(True)
    assert model.FeatureByName("筋脚宽度").Select2(False, 1)
    assert model.FeatureByName("筋顶圆弧").Select2(True, 1)
    model.InsertLoftRefSurface2(False, True, False, 1., 0, 0)
    f = g(model, "FirstFeature")
    loft = None
    while f is not None:
        if g(f, "GetTypeName2") == "BlendRefSurface":
            loft = f
        f = g(f, "GetNextFeature")
    assert loft is not None, "未生成放样曲面"
    loft.Name = "筋顶放样曲面"
    return boss_to_loft(model, bottom, top)


def boss_to_loft(model, bottom="筋脚平面", top="筋顶平面"):
    sheets = g(model, "GetBodies2", 1, False)
    assert len(sheets or []) == 1
    faces = g(sheets[0], "GetFaces")
    assert len(faces) == 1
    face = faces[0]
    with sketch(model, bottom) as ref:
        model.ClearSelection2(True)
        assert face.Select4(False, g(model.SelectionManager, "CreateSelectData"))
        assert model.SketchManager.SketchUseEdge3(False, False)
    model.FeatureByName(str(ref)).Name = "筋投影轮廓"
    _ensure_sketch_selected(model, "筋投影轮廓")
    data = g(model.SelectionManager, "CreateSelectData")
    data.Mark = 1
    assert face.Select4(True, data)
    feature = model.FeatureManager.FeatureExtrusion3(
        True, False, False, 4, 0, .020, 0., False, False, False, False,
        0., 0., False, False, False, False, True, False, True, 0, 0., False)
    assert feature is not None, "拉伸到筋顶放样曲面失败"
    feature.Name = "加强筋"
    assert model.ForceRebuild3(False)
    assert len(g(model, "GetBodies2", 0, False)) == 1
    model.ClearSelection2(True)
    sheets[0].HideBody(True)
    for name in [bottom, top]:
        model.ClearSelection2(True)
        assert model.FeatureByName(name).Select2(False, 0)
        model.BlankRefGeom()
    model.ClearSelection2(True)
    model.ShowNamedView2("*Isometric", 7)
    model.ViewZoomtofit2()
    model.GraphicsRedraw2()
    return feature
