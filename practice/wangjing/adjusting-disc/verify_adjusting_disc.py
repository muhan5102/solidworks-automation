"""独立按原图数值核验保存后回读的B-Rep；不导入建模脚本参数。"""
import json
import math
from pathlib import Path
import sys

out = Path(sys.argv[1]).resolve()
report = json.loads((out/"verification.json").read_text(encoding="utf-8"))
faces = report["faces"]
cylinders = [f for f in faces if "cylinder" in f]


def close(a, b):
    # API单位米，以下比较统一转为毫米；1e-5 mm用于数值回读，不是制造公差。
    return abs(a-b) < 1e-5


def cylinder(radius, x, z, lo, hi):
    return [f for f in cylinders if close(f["cylinder"][6]*1000, radius)
            and close(f["cylinder"][0]*1000, x)
            and close(f["cylinder"][2]*1000, z)
            and abs(abs(f["cylinder"][4])-1) < 1e-8
            and close(f["box_m"][1]*1000, lo) and close(f["box_m"][4]*1000, hi)]


checks = {
    "四张草图全约束无Fix": report["sketches"]["status"] == "pass"
        and len(report["sketches"]["sketches"]) == 4,
    "特征无错误无警告": not report["health"]["faulty_features"]
        and not report["health"]["warning_features"],
    "后孔D64至11": len(cylinder(32, 0, 0, 0, 11)) == 1,
    "内环槽D74宽13位置11至24": len(cylinder(37, 0, 0, 11, 24)) == 1,
    "内孔D60及前口C2起点": len(cylinder(30, 0, 0, 24, 56)) == 1,
    "前沉台D104深6": len(cylinder(52, 0, 0, 58, 64)) == 1,
    "法兰D180及前外缘C2起点": len(cylinder(90, 0, 0, 42, 62)) == 1,
    "轮毂D100及后外缘C2终点": len(cylinder(50, 0, 0, 2, 37)) == 1,
    "退刀槽5乘1实际D98": len(cylinder(49, 0, 0, 37, 42)) == 1,
}
for radius, label in [(64, "内"), (78, "外")]:
    fs = cylinder(radius, 0, 0, 42, 64)
    expected_area = radius * math.pi / 6 * 22
    checks[f"六弧槽{label}壁半径与30度跨度"] = len(fs) == 6 and all(
        abs(f["area_m2"]*1e6-expected_area) < 1e-4 for f in fs)
for theta in range(0, 360, 30):
    # 六段各跨30度，端圆中心每30度出现，共12处；方向由两个独立孔另行确定。
    a = math.radians(theta)
    fs = cylinder(7, 71*math.cos(a), -71*math.sin(a), 42, 64)
    checks[f"槽端R7_分布圆D142_角{theta}"] = len(fs) == 1 and abs(
        fs[0]["area_m2"]*1e6-math.pi*7*22) < 1e-4
for radius, angle, label in [(4, 15, "D8H7上右通孔"), (5, -45, "D10H7下右销孔")]:
    a = math.radians(angle)
    checks[label] = len(cylinder(radius, 71*math.cos(a), -71*math.sin(a), 42, 64)) == 1
cones = [f for f in faces if "cone" in f]
checks["仅三处图示C2且位置正确"] = len(cones) == 3 and all(
    any(close(f["box_m"][1]*1000, lo) and close(f["box_m"][4]*1000, hi)
        and close(math.tan(f["cone"][7]), 1) for f in cones)
    for lo, hi in [(0, 2), (56, 58), (62, 64)])
result = {"checks": checks, "passed": all(checks.values()),
          "note": "B-Rep几何检查需与原图视图目视合用；材料、制造偏差和工程图未由此验证。"}
(out/"independent-checks.json").write_text(
    json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(result, ensure_ascii=True, indent=2))
assert result["passed"]
