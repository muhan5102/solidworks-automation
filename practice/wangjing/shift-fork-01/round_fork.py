"""只处理原图画出的叉根、杆毂和筋根过渡，未注半径选R2。"""
import math
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.single_part import g
from sw_part import fillet


def near(a, b):
    return abs(a-b) < 1e-7


def edge_rows(model):
    for body in g(model, "GetBodies2", 0, False):
        for edge in g(body, "GetEdges"):
            curve = g(edge, "GetCurve")
            va, vb = g(edge, "GetStartVertex"), g(edge, "GetEndVertex")
            yield edge, curve, (None if va is None else list(g(va, "GetPoint"))), (
                None if vb is None else list(g(vb, "GetPoint")))


def apply(model, name, predicate, count):
    assert model.FeatureByName(name) is None
    edges = [e for e, c, a, b in edge_rows(model) if predicate(c, a, b)]
    assert len(edges) == count, (name, len(edges), count)
    model.ClearSelection2(True)
    data = g(model.SelectionManager, "CreateSelectData")
    for edge in edges:
        assert edge.Select4(True, data)
    feature = fillet(model, .002)
    assert feature is not None, name
    feature.Name = name
    assert model.ForceRebuild3(False)
    model.ClearSelection2(True)


def finish_rounds(model):
    fork_x = math.sqrt(.027**2-.015**2)
    def fork(curve, a, b):
        if g(curve, "IsCircle"):
            p = g(curve, "CircleParams")
            return near(p[0], 0) and near(p[6], .027) and (
                near(p[1], .008) or near(p[1], .002))
        return a is not None and b is not None and g(curve, "IsLine") and (
            near(a[0], fork_x) and near(b[0], fork_x) and
            near(abs(a[2]), .015) and near(a[2], b[2]) and
            near(min(a[1], b[1]), .002) and near(max(a[1], b[1]), .008))
    if model.FeatureByName("叉根过渡R2") is None:
        apply(model, "叉根过渡R2", fork, 4)

    root_x = .120-math.sqrt(.015**2-.003**2)
    def hub_rib(curve, a, b):
        if g(curve, "IsCircle"):
            p = g(curve, "CircleParams")
            return near(p[0], .120) and near(p[6], .015) and (
                near(p[1], .002) or near(p[1], .008))
        if a is None or b is None or not g(curve, "IsLine"):
            return False
        same_side = near(abs(a[2]), .003) and near(a[2], b[2])
        foot = near(a[1], .008) and near(b[1], .008) and (
            near(min(a[0], b[0]), .070) and near(max(a[0], b[0]), root_x))
        # 到顶的筋根竖边未明确图示圆角；对其加R2会在顶面留下截平小面。
        return same_side and foot
    apply(model, "杆毂与筋脚过渡R2", hub_rib, 5)
