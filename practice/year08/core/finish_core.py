"""按内芯原图补做未注R0.5/C0.5；以几何签名选择，不依赖边序号。"""
import json
from pathlib import Path
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member, mm, save_document
from sw_part import fillet, chamfer
from sw_edge_select import _edge_curve

ROUNDS = [(19.8, 3.5)]
CHAMFERS = [(4, 8.5), (7.5, 7), (13, 7)]


def near(a, b):
    return abs(a - b) < 1e-5


def boundary_edges(model, targets, flat_ends):
    matches = []
    found = {target: 0 for target in targets}
    for body in member(model, "GetBodies2", 0, False):
        for edge in member(body, "GetEdges"):
            curve = _edge_curve(edge)
            if curve is None:
                continue
            if member(curve, "IsCircle"):
                p = list(member(curve, "CircleParams"))
                for y, r in targets:
                    if near(p[0]*1000, 0) and near(p[2]*1000, 0) and near(p[1]*1000, y) and near(p[6]*1000, r):
                        matches.append((edge, {"circle_y_r_mm": [y, r]}))
                        found[(y, r)] += 1
            elif member(curve, "IsLine"):
                vertices = [member(edge, k) for k in ("GetStartVertex", "GetEndVertex")]
                if any(v is None for v in vertices):
                    continue
                a, b = [[c*1000 for c in member(v, "GetPoint")] for v in vertices]
                for y, x in flat_ends:
                    if near(a[1], y) and near(b[1], y) and near(abs(a[0]), x) and near(a[0], b[0]):
                        matches.append((edge, {"flat_y_x_mm": [y, a[0]]}))
    assert all(found.values()), found
    return matches


def finish(model):
    evidence = []
    for name, kind, targets, flats, count in [
        ("末端圆角", "R", ROUNDS, [], 1),
        ("台阶倒角", "C", CHAMFERS, [], 3),
    ]:
        if model.FeatureByName(name) is not None:
            continue
        selected = boundary_edges(model, targets, flats)
        assert len(selected) == count, (name, len(selected), [p for _, p in selected])
        model.ClearSelection2(True)
        for edge, _ in selected:
            assert edge.Select2(True, 0)
        feature = fillet(model, mm(.5)) if kind == "R" else chamfer(model, mm(.5))
        assert feature is not None, f"{name}创建失败，未降低图示尺寸"
        feature.Name = name
        assert model.ForceRebuild3(False)
        assert len(member(model, "GetBodies2", 0, False)) == 1
        model.ClearSelection2(True)
        model.GraphicsRedraw2()
        evidence.append({"name": name, "type": kind, "size_mm": .5,
                         "selected_edges": [p for _, p in selected]})
        print(f"[完成] {name} {kind}0.5，选择{len(selected)}条边", flush=True)
    return evidence


if __name__ == "__main__":
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    out = Path(sys.argv[1]).resolve()
    target = out / "内芯.SLDPRT"
    assert target.is_file()
    sw = win32com.client.Dispatch("SldWorks.Application.32")
    assert int(member(sw, "GetProcessID")) == int(sys.argv[2])
    model = sw.ActiveDoc
    assert Path(member(model, "GetPathName")).resolve() == target
    sw.Visible = True
    started = time.monotonic()
    evidence = finish(model)
    assert save_document(model)
    (out / "edge-treatment.json").write_text(json.dumps(
        {"treatments": evidence, "execution_seconds": time.monotonic()-started},
        ensure_ascii=False, indent=2), encoding="utf-8")
