"""独立按LJT05.01题面尺寸复核；放样曲面是隐藏构造体。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.single_part import g, prepare_review
import win32com.client


def near(a, b):
    return abs(a-b) < 1e-7


def geometry(model, rows):
    cylinders = [r["Cylinder"] for r in rows if "Cylinder" in r]
    cones = [r["Cone"] for r in rows if "Cone" in r]
    def cyl(center, radius, axis):
        return any(all(near(a, b) for a, b in zip(c[:3], center))
                   and near(c[6], radius)
                   and all(near(abs(a), abs(b)) for a, b in zip(c[3:6], axis))
                   for c in cylinders)
    foot = g(model.FeatureByName("筋脚宽度"), "GetSpecificFeature2")
    line = list(g(foot, "GetSketchSegments"))[0]
    a, b = g(line, "GetStartPoint2"), g(line, "GetEndPoint2")
    arc_sketch = g(model.FeatureByName("筋顶圆弧"), "GetSpecificFeature2")
    arc = list(g(arc_sketch, "GetSketchSegments"))[0]
    # 局部X-Z坐标；世界Y由真实草图变换确认。
    foot_transform = list(g(g(foot, "ModelToSketchTransform"), "Inverse").ArrayData)
    top_transform = list(g(g(arc_sketch, "ModelToSketchTransform"), "Inverse").ArrayData)
    outer_top = []
    for body in g(model, "GetBodies2", 0, False):
        for face in g(body, "GetFaces"):
            surf = g(face, "GetSurface")
            if not g(surf, "IsPlane"):
                continue
            pp = g(surf, "PlaneParams")
            if not near(pp[4], .028):
                continue
            for edge in g(face, "GetEdges"):
                cur = g(edge, "GetCurve")
                if g(cur, "IsCircle") and near(g(cur, "CircleParams")[6], .015):
                    outer_top.append(edge)
                elif g(cur, "IsLine"):
                    outer_top.append(None)
    return {
        "叉外R27及叉内直径27": cyl((0., .010, 0.), .027, (0, 1, 0))
            and cyl((0., .010, 0.), .0135, (0, 1, 0)),
        "毂外直径30及中心距120": cyl((.120, .028, 0.), .015, (0, 1, 0)),
        "主孔直径15贯通及两口C1": cyl((.120, 0., 0.), .0075, (0, 1, 0))
            and sum(near(c[0], .120) and near(abs(c[4]), 1.) for c in cones) == 2,
        "径向孔直径5高14轴向Z": cyl((.120, .014, 0.), .0025, (0, 0, 1)),
        "径向沉头直径11半角45": any(near(c[0], .120) and near(c[1], .014)
            and near(c[6], .0025) and near(c[7], 0.7853981633974483) for c in cones)
            and cyl((.120, .014, 0.), .0055, (0, 0, 1)),
        "筋脚中心距50宽6高度8": near(g(a, "X"), .070)
            and near(g(b, "X"), .070)
            and near(abs(g(a, "Y")-g(b, "Y")), .006)
            and near(foot_transform[10], .008),
        "筋顶共圆R15高度28": near(g(arc, "GetRadius"), .015)
            and near(g(g(arc, "GetCenterPoint2"), "X"), .120)
            and near(top_transform[10], .028),
        "圆毂顶面无额外直线截平台": bool(outer_top) and None not in outer_top,
        "图示两组R2保留": all(model.FeatureByName(n) is not None
            for n in ["叉根过渡R2", "杆毂与筋脚过渡R2"]),
        "筋为实体拉伸并引用放样曲面": g(model.FeatureByName("加强筋"), "GetTypeName2")
            in ("ICE", "Extrusion") and model.FeatureByName("筋顶放样曲面") is not None,
    }


if __name__ == "__main__":
    sw = win32com.client.Dispatch("SldWorks.Application.32")
    prepare_review(sw, Path(sys.argv[1]).resolve(), geometry, sketch_count=8,
                   views=["isometric", "front", "top", "right_wire"])
