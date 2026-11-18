"""检查已打开的 SolidWorks 单实体零件，可按题面核对轴向钻尖盲孔。"""

import argparse
from contextlib import redirect_stdout
import json
import math
from pathlib import Path
import sys


def read_blind_hole(value):
    try:
        x, y, z, axis, diameter, depth, tip = value.split(",")
        entry = tuple(map(float, (x, y, z)))
        diameter, depth, tip = map(float, (diameter, depth, tip))
        axis = axis.strip().upper()
        if axis not in ("X+", "X-", "Y+", "Y-", "Z+", "Z-"):
            raise ValueError
        if not all(map(math.isfinite, (*entry, diameter, depth, tip))):
            raise ValueError
        if not (diameter > 0 and 0 < depth < tip):
            raise ValueError
        return entry, axis, diameter, depth, tip
    except ValueError as error:
        raise argparse.ArgumentTypeError("需要入口X,Y,Z,方向,直径,全径深,钻尖总深，例如 0,0,-33,Y+,6.8,22,24.043") from error


def check_blind_hole(rows, hole, tolerance):
    """以实体圆柱、导入锥和终止锥验证沿坐标轴的完整孔壁，尺寸为毫米。"""
    entry, direction, diameter, depth, tip_depth = hole
    index = "XYZ".index(direction[0])
    sign = 1 if direction[1] == "+" else -1
    transverse = [i for i in range(3) if i != index]
    radius = diameter / 2

    def near(actual, expected):
        return abs(actual - expected) <= tolerance

    def intervals(kind):
        for row in rows:
            if not row.get("internal") or kind not in row:
                continue
            params = row[kind]
            if not math.isclose(abs(params[index + 3]), 1, abs_tol=1e-8):
                continue
            if math.hypot(*(params[i] * 1000 - entry[i] for i in transverse)) > tolerance:
                continue
            start, end = sorted(sign * (row["box"][i] - entry[index]) for i in (index, index + 3))
            yield start, end, params, row["area_mm2"]

    cylinders = list(intervals("Cylinder"))
    cones = list(intervals("Cone"))
    report = {"entry_mm": entry, "direction": direction, "diameter_mm": diameter,
              "depth_mm": depth, "tip_depth_mm": tip_depth, "passed": False,
              "cylinder_spans_mm": [[round(a, 5), round(b, 5), round(p[6] * 2000, 5)]
                                    for a, b, p, _ in cylinders],
              "cone_spans_mm": [[round(a, 5), round(b, 5)] for a, b, _, _ in cones]}
    bores = [(start, end) for start, end, params, area in cylinders
             if end > start and start >= -tolerance and near(end, depth)
             and near(params[6] * 1000, radius)
             and abs(area - math.tau * radius * (end - start)) <= math.tau * radius * tolerance]
    if len(bores) != 1:
        report["message"] = "未找到唯一、完整且终点符合预期的圆柱孔壁"
        return report
    start, end = bores[0]

    def cone_matches(cone, first, last, is_tip):
        low, high, params, area = cone
        if not (high > low and near(low, first) and near(high, last)):
            return False
        height = high - low
        radial_change = height * abs(math.tan(params[7]))
        if is_tip and not near(radial_change, radius):
            return False
        radius_sum = radius if is_tip else 2 * radius + radial_change
        full_area = math.pi * radius_sum * math.hypot(height, radial_change)
        return abs(area - full_area) <= math.pi * radius_sum * tolerance

    closed_tip = any(cone_matches(cone, end, tip_depth, True) for cone in cones)
    entrance = near(start, 0) or any(cone_matches(cone, 0, start, False) for cone in cones)
    report.update(passed=entrance and closed_tip, entrance_matches=entrance, closed_tip_matches=closed_tip)
    if not report["passed"]:
        report["message"] = "孔口或完整钻尖与预期不符，检查孔口几何、误穿或相交"
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("part", type=Path, help="已打开且活动的 SLDPRT 完整路径")
    parser.add_argument("--pid", required=True, type=int, help="既有 SolidWorks 的进程 ID")
    parser.add_argument("--blind", type=read_blind_hole, action="append", default=[],
                        help="入口X,Y,Z,方向,直径,全径深,钻尖总深，单位毫米，可重复提供")
    parser.add_argument("--tol-mm", type=float, default=.02, help="几何数值比对容差，非制造公差")
    args = parser.parse_args()
    if not math.isfinite(args.tol_mm) or args.tol_mm <= 0:
        parser.error("--tol-mm 必须是有限正数")
    target = args.part.resolve()
    if target.suffix.lower() != ".sldprt" or not target.is_file():
        parser.error("请指定磁盘上存在的 SLDPRT")

    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from shared.single_part import g, surfaces
    from sw_connect import _attach_existing_with_pid_check, solidworks_processes
    from sw_review import collect_model_summary
    from sw_sketch_quality import inspect_model_sketches

    processes = solidworks_processes() or {}
    year = processes.get(args.pid)
    if year is None or list(processes.values()).count(year) != 1:
        parser.error("目标进程或版本不唯一，无法安全附着")
    with redirect_stdout(sys.stderr):
        sw = _attach_existing_with_pid_check(processes, year)
    if g(sw, "GetProcessID") != args.pid:
        parser.error("SolidWorks 进程与目标不一致")
    model = g(sw, "ActiveDoc")
    if model is None or g(model, "GetType") != 1 or Path(g(model, "GetPathName")).resolve() != target:
        parser.error("活动零件与目标完整路径不一致")

    health, sketches = collect_model_summary(model), inspect_model_sketches(model)
    bodies = g(model, "GetBodies2", 0, False) or []
    rows = surfaces(model, hole_evidence=True) if args.blind else []
    holes = [check_blind_hole(rows, hole, args.tol_mm) for hole in args.blind]
    failed = (len(bodies) != 1 or health.get("faulty_features") or sketches["status"] == "fail"
              or any(not hole["passed"] for hole in holes))
    review = (sketches["status"] != "pass" or any(health.get(key) for key in
              ("feature_error", "warning_features", "unknown_feature_health")))
    print(json.dumps({
        "part": target.name,
        "status": "fail" if failed else "review_required" if review else "checks_passed",
        "solid_bodies": len(bodies),
        "feature_issues": {key: health.get(key) for key in
                           ("feature_error", "faulty_features", "warning_features", "unknown_feature_health")},
        "sketches": {s["name"]: s["status"] for s in sketches["sketches"]},
        "blind_holes": holes,
        "limits": "只读活动零件，不执行保存重开或目视；盲孔仅检查指定的轴向完整圆柱与锥底。",
    }, ensure_ascii=False, indent=2))
    return 1 if failed else 2 if review else 0


if __name__ == "__main__":
    raise SystemExit(main())
