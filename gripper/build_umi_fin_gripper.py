# -*- coding: utf-8 -*-
"""
UMI 风格鱼鳍形软爪（Fin-Ray Soft Gripper）参数化建模脚本
目标软件：SOLIDWORKS 2024（中文版）

流程：
  1. 用 numpy 参数化生成二维主截面（外轮廓 + 16 个斜置气腔 + 3 个贯穿圆孔）
  2. 连接 SolidWorks（优先连接已运行实例，否则自动启动）
  3. 在前视基准面绘制闭合草图（绝对坐标、无几何约束、AddToDB 高速绘制）
  4. 选中草图特征后拉伸生成实体（本机验证可用的路径）
  5. 保存 .SLDPRT 到桌面（中文名，失败自动回退 ASCII 名）

用法：
  python build_umi_fin_gripper.py                # 完整建模
  python build_umi_fin_gripper.py --check-only   # 只打印参数并校验几何

坐标约定（重要）：
  前视基准面：X = 轮廓宽度(≈40mm)，Y = 轮廓高度/长度(≈123.36mm)，
  拉伸方向 Z = 厚度(25.8mm)
  对应需求中的 X40 / Y25.8 / Z123.36 -> SW.X=40，SW.Y=123.36，SW.Z=25.8

本机适配说明（已实测）：
  - sw.RevisionNumber 是【属性】不是方法；EditRebuild3 同样以属性访问
  - SelectByID2 在本机 typelib 下不可用 -> 用 FeatureByPositionReverse 遍历特征树选基准面
  - 拉伸前需先选中草图特征，再调用 FeatureExtrusion2（盲拉伸）
"""

import os
import sys
import time

# ---------------------------------------------------------------------------
# pywin32 引导：优先使用本机 Python（D:\Python）自带的 pywin32；
# 若系统未装，则回退到工作区 pylibs2（供 Codex 内置 Python 使用）
# ---------------------------------------------------------------------------
_PYLIB = os.path.join(os.path.expanduser("~"), "Documents", "Codex", "2026-08-15", "he-2", "pylibs2")
try:
    import win32com.client
except Exception:
    if os.path.isdir(_PYLIB):
        for _p in (_PYLIB,
                   os.path.join(_PYLIB, "win32"),
                   os.path.join(_PYLIB, "win32", "lib")):
            if _p not in sys.path:
                sys.path.insert(0, _p)
        if hasattr(os, "add_dll_directory"):
            os.add_dll_directory(os.path.join(_PYLIB, "pywin32_system32"))
    import win32com.client

import numpy as np

# 控制台统一 UTF-8 输出，避免 GBK 控制台编码报错
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# ===========================================================================
# 1. 关键设计参数
# ===========================================================================
BB_X = 40.0          # 轮廓总宽 mm（SolidWorks X）
BB_Y = 123.36        # 轮廓总高/长度 mm（SolidWorks Y）
BB_Z = 25.8          # 拉伸厚度 mm（SolidWorks Z）

N_CHAMBERS = 16      # 波纹气腔数量
PITCH = 6.0          # 气腔节距 mm
YC0 = 13.0           # 第一个气腔中心到尖端的距离 mm
CH_H = 4.1           # 气腔高度 mm（腔间壁厚 = PITCH - CH_H ≈ 1.9 ≈ T）
T_WALL = 1.8         # 参考壁厚 mm
SLANT_DEG = 15.0     # 气腔倾斜角
ML = 2.2             # 气腔左侧边距 mm
MR = 3.4             # 气腔右侧边距 mm（常规）
MR_HOLE = 10.0       # 侧孔附近气腔的右侧加厚边距 mm（外鼓区）
HOLE_R = 2.0         # 贯穿圆孔半径 mm（孔径 4mm）
HOLES = [(10.0, 112.0),   # 左上
         (30.0, 112.0),   # 右上
         (33.5, 79.0)]    # 右侧中上部（外鼓区）

TPL = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates\gb_part.prtdot"
OUT_DIR = os.path.dirname(os.path.abspath(__file__))
DESKTOP_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
PART_NAME_CN = "\u9c7c\u9ccd\u8f6f\u722a.SLDPRT"      # 鱼鳍软爪.SLDPRT
PART_NAME_ASCII = "UMI_FinGripper.SLDPRT"
LOG_FILE = os.path.join(OUT_DIR, "build_log.txt")


def mm2m(v):
    """SolidWorks API 底层单位永远是米，所有几何坐标必须显式换算。"""
    return float(v) / 1000.0


def invoke(obj, name, *args):
    """兼容本机 typelib：先按方法调用，若不可调用则按属性读取。"""
    attr = getattr(obj, name)
    try:
        return attr(*args)
    except TypeError:
        return attr


def Log(msg):
    """同时输出到控制台与日志文件。"""
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


# ===========================================================================
# 2. 二维主截面几何生成
# ===========================================================================
def outer_control_points():
    """外轮廓控制点（逆时针）：上边缘平直、右侧中上部外鼓、底部收尖。"""
    return np.array([
        [ 4.0, 123.36], [12.0, 123.36], [20.0, 123.36], [28.0, 123.36], [36.0, 123.36],
        [36.0, 104.0], [37.5, 90.0], [38.8, 78.0], [37.2, 66.0],
        [35.0, 50.0], [32.0, 34.0], [28.0, 20.0], [23.5, 8.0], [19.5, 1.0],
        [18.0, 0.0],
        [16.0, 1.2], [12.0, 10.0], [8.8, 26.0], [6.8, 46.0],
        [5.8, 68.0], [5.2, 90.0], [4.7, 108.0], [4.0, 123.36],
    ], dtype=float)


def catmull_rom_closed(pts, samples=10):
    """闭合 Catmull-Rom 样条采样，输出平滑的密集折线点列。
    跳过首尾重复控制点对应的零长度段（否则会产生回折尖刺，导致无法拉伸）。"""
    n = len(pts)
    last = n - 1 if np.allclose(pts[0], pts[-1]) else n
    out = []
    for i in range(last):
        p0, p1, p2, p3 = pts[(i-1) % n], pts[i], pts[(i+1) % n], pts[(i+2) % n]
        for t in np.linspace(0.0, 1.0, samples, endpoint=False):
            t2, t3 = t * t, t * t * t
            x = 0.5 * (2*p1[0] + (-p0[0]+p2[0])*t +
                       (2*p0[0]-5*p1[0]+4*p2[0]-p3[0])*t2 +
                       (-p0[0]+3*p1[0]-3*p2[0]+p3[0])*t3)
            y = 0.5 * (2*p1[1] + (-p0[1]+p2[1])*t +
                       (2*p0[1]-5*p1[1]+4*p2[1]-p3[1])*t2 +
                       (-p0[1]+3*p1[1]-3*p2[1]+p3[1])*t3)
            out.append([float(x), float(y)])
    out.append(out[0])
    arr = np.array(out)
    # 样条拐点处可能轻微过冲：夹取到设计边界，保证 Bounding Box 精确
    arr[:, 0] = np.clip(arr[:, 0], 0.0, BB_X)
    arr[:, 1] = np.clip(arr[:, 1], 0.0, BB_Y)
    # 保险：顶边不允许 x 回折到设计角点之外
    top = arr[:, 1] == BB_Y
    arr[top, 0] = np.clip(arr[top, 0], 4.0, 36.0)
    return arr


def edge_x(dense, y):
    """求闭合多边形在高度 y 处与水平线的左右交点 x 值。"""
    xs = []
    for i in range(len(dense) - 1):
        x1, y1 = dense[i]
        x2, y2 = dense[i + 1]
        if (y1 <= y <= y2) or (y2 <= y <= y1):
            if abs(y2 - y1) < 1e-9:
                continue
            t = (y - y1) / (y2 - y1)
            xs.append(x1 + t * (x2 - x1))
    if not xs:
        return (0.0, 0.0)
    return min(xs), max(xs)


def build_chambers(dense):
    """生成 N_CHAMBERS 个斜置、逐级排布的闭合四边形气腔。"""
    quads = []
    dx = CH_H * np.tan(np.radians(SLANT_DEG))
    for i in range(N_CHAMBERS):
        yc = YC0 + i * PITCH
        yb, yt = yc - CH_H / 2.0, yc + CH_H / 2.0
        mr = MR_HOLE if abs(yc - HOLES[2][1]) < 9.0 else MR
        xl0 = edge_x(dense, yb)[0] + ML
        xr0 = edge_x(dense, yb)[1] - mr
        xl1 = edge_x(dense, yt)[0] + ML + dx
        xr1 = edge_x(dense, yt)[1] - mr + dx
        quads.append(np.array([
            [xl0, yb], [xr0, yb], [xr1, yt], [xl1, yt],
        ], dtype=float))
    return quads


def poly_area(poly):
    """鞋带公式求多边形面积（mm²）。"""
    x, y = poly[:, 0], poly[:, 1]
    return 0.5 * abs(np.dot(x, np.roll(y, 1)) - np.dot(y, np.roll(x, 1)))


def build_geometry():
    """组装全部几何，返回 (外轮廓点, 气腔列表, 孔列表, 统计信息)。"""
    dense = catmull_rom_closed(outer_control_points())
    quads = build_chambers(dense)

    outer_area = poly_area(dense)
    chamber_area = sum(poly_area(q) for q in quads)
    hole_area = len(HOLES) * np.pi * HOLE_R * HOLE_R
    net_area = outer_area - chamber_area - hole_area

    left_walls, right_walls = [], []
    for q in quads:
        yb, yt = q[0][1], q[2][1]
        left_walls.append(q[0][0] - edge_x(dense, yb)[0])
        right_walls.append(edge_x(dense, yt)[1] - q[2][0])

    stats = dict(
        outer_area=outer_area,
        chamber_area=chamber_area,
        hole_area=hole_area,
        net_area=net_area,
        expected_volume=net_area * BB_Z,
        min_left_wall=min(left_walls),
        min_right_wall=min(right_walls),
        wall_between=PITCH - CH_H,
        bbox=(float(dense[:, 0].min()), float(dense[:, 0].max()),
              float(dense[:, 1].min()), float(dense[:, 1].max())),
    )
    return dense, quads, HOLES, stats


def print_parameters(stats):
    """建模前打印关键参数供核对。"""
    print("=" * 64)
    print("UMI 鱼鳍形软爪 - 关键设计参数")
    print("=" * 64)
    print(f"整体 Bounding Box   : X {stats['bbox'][0]:.2f} ~ {stats['bbox'][1]:.2f} mm"
          f"  |  Y {stats['bbox'][2]:.2f} ~ {stats['bbox'][3]:.2f} mm"
          f"  |  拉伸厚度 Z = {BB_Z} mm")
    print(f"气腔数量            : N = {N_CHAMBERS}")
    print(f"气腔节距            : D = {PITCH} mm (腔中心 {YC0} ~ {YC0 + (N_CHAMBERS-1)*PITCH:.2f} mm)")
    print(f"气腔尺寸            : 高 {CH_H} mm, 倾角 {SLANT_DEG}°, "
          f"腔间壁厚 ≈ {stats['wall_between']:.2f} mm (参考 T={T_WALL})")
    print(f"侧壁厚              : 左侧最小 {stats['min_left_wall']:.2f} mm, "
          f"右侧最小 {stats['min_right_wall']:.2f} mm")
    print(f"贯穿圆孔            : {len(HOLES)} 个, 孔径 {HOLE_R*2:.1f} mm")
    for i, (hx, hy) in enumerate(HOLES, 1):
        print(f"   孔{i}: 中心 ({hx:.1f}, {hy:.1f}) mm")
    print(f"截面面积            : 外轮廓 {stats['outer_area']:.1f} mm2, "
          f"气腔 {stats['chamber_area']:.1f} mm2, 孔 {stats['hole_area']:.1f} mm2")
    print(f"实体净面积 / 体积   : {stats['net_area']:.1f} mm2 / {stats['expected_volume']:.1f} mm3")
    print("坐标约定            : SW 前视面 X=宽(40)/Y=高(123.36), 拉伸 Z=厚(25.8)")
    print("=" * 64)


# ===========================================================================
# 3. SolidWorks 建模
# ===========================================================================
def connect_solidworks():
    """连接已运行实例，否则启动新实例，并等待 API 就绪。"""
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        Log("[OK] 已连接正在运行的 SolidWorks 实例")
    except Exception:
        Log("[INFO] 未发现运行中的 SolidWorks，正在启动新实例...")
        sw = win32com.client.Dispatch("SldWorks.Application")

    t0 = time.time()
    while time.time() - t0 < 240:
        try:
            rev = str(sw.RevisionNumber)   # 本机为属性
            Log(f"[OK] SolidWorks 就绪，版本: {rev}")
            try:
                sw.Visible = True          # 使窗口可见，便于观察
                Log("[OK] 已将 SolidWorks 窗口设为可见")
            except Exception:
                pass
            return sw
        except Exception:
            time.sleep(1)
    raise RuntimeError("SolidWorks 未在 240s 内就绪")


def find_feature(model, keywords):
    """按名称在特征树中查找特征（FeatureByPositionReverse 遍历）。"""
    for i in range(80):
        try:
            f = model.FeatureByPositionReverse(i)
        except Exception:
            break
        if f is None:
            break
        try:
            n = str(invoke(f, "Name"))
        except Exception:
            n = ""
        if any(k in n for k in keywords):
            return f, n
    return None, None


def select_front_plane(model):
    """选中前视基准面（不依赖 SelectByID2，本机不可用）。"""
    candidates = []
    for i in range(40):
        try:
            feat = model.FeatureByPositionReverse(i)
        except Exception:
            break
        if feat is None:
            break
        try:
            fname = str(invoke(feat, "Name"))
        except Exception:
            fname = ""
        if not fname:
            continue
        priority = 0 if ("前视" in fname or "Front" in fname) else 1
        candidates.append((priority, fname, feat))
    candidates.sort(key=lambda c: c[0])
    for priority, fname, feat in candidates:
        try:
            if feat.Select2(False, 0):
                Log(f"[OK] 已选中基准面: {fname}")
                return True
        except Exception:
            pass
    return False


def draw_polyline(sk, pts):
    """逐段绘制闭合折线（米制），跳过零长度线段。"""
    for i in range(len(pts) - 1):
        x1, y1 = pts[i]
        x2, y2 = pts[i + 1]
        if abs(x2 - x1) < 1e-9 and abs(y2 - y1) < 1e-9:
            continue
        sk.CreateLine(mm2m(x1), mm2m(y1), 0.0, mm2m(x2), mm2m(y2), 0.0)
    x1, y1 = pts[-1]
    x2, y2 = pts[0]
    if not (abs(x2 - x1) < 1e-9 and abs(y2 - y1) < 1e-9):
        sk.CreateLine(mm2m(x1), mm2m(y1), 0.0, mm2m(x2), mm2m(y2), 0.0)


def build_part(sw, dense, quads, holes):
    """新建零件 -> 前视基准面画草图 -> 选中草图特征 -> 拉伸 -> 保存。"""
    model = sw.NewDocument(TPL, 0, 0, 0)
    if model is None:
        model = sw.ActiveDoc
    if model is None:
        raise RuntimeError("新建零件文档失败（模板可能未找到）")
    Log("[OK] 新零件文档已创建")
    time.sleep(1.0)

    sk = model.SketchManager

    # ---- 选择前视基准面并进入草图 ----
    Log("[步骤] 选择前视基准面...")
    model.ClearSelection2(True)
    if not select_front_plane(model):
        raise RuntimeError("无法选择前视基准面")
    try:
        sk.InsertSketch2(True)
    except Exception:
        sk.InsertSketch(True)

    # ---- 高速绘制草图 ----
    sk.AddToDB = True
    sk.DisplayWhenAdded = False
    draw_polyline(sk, dense)                      # 外轮廓（220 段圆滑折线）
    for q in quads:                               # 16 个斜置气腔
        draw_polyline(sk, np.vstack([q, q[0:1]]))
    for hx, hy in holes:                          # 3 个贯穿圆孔
        sk.CreateCircleByRadius(mm2m(hx), mm2m(hy), 0.0, mm2m(HOLE_R))
    sk.AddToDB = False
    sk.DisplayWhenAdded = True
    try:
        sk.InsertSketch2(True)
    except Exception:
        sk.InsertSketch(True)
    Log("[OK] 草图绘制完成（外轮廓 + 16 气腔 + 3 孔）")

    # ---- 选中草图特征后拉伸（本机验证可用的路径） ----
    feat = None
    sketch_feat, sketch_name = find_feature(model, ("草图", "Sketch"))
    if sketch_feat is not None:
        model.ClearSelection2(True)
        try:
            sketch_feat.Select2(False, 0)
            Log(f"[OK] 已选中草图特征: {sketch_name}")
        except Exception as e:
            Log(f"[WARN] 选中草图特征失败: {e}")
        feat = model.FeatureManager.FeatureExtrusion2(
            True, False, False, 0, 0, mm2m(BB_Z), 0,
            False, False, False, False, 0, 0,
            False, False, False, False,
            True, True, True, 0, 0, False)
    if feat is None:
        raise RuntimeError("拉伸特征创建失败")
    Log(f"[OK] 拉伸特征: {feat.Name}")

    # ---- 保存：桌面优先，中文名 -> ASCII 回退 ----
    os.makedirs(OUT_DIR, exist_ok=True)
    save_paths = [os.path.join(DESKTOP_DIR, PART_NAME_CN),
                  os.path.join(DESKTOP_DIR, PART_NAME_ASCII),
                  os.path.join(OUT_DIR, PART_NAME_CN),
                  os.path.join(OUT_DIR, PART_NAME_ASCII),
                  os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"),
                               PART_NAME_ASCII)]
    saved = None
    for path in save_paths:
        try:
            err = model.SaveAs3(path, 0, 1)
            if os.path.exists(path):
                saved = path
                Log(f"[OK] 零件已保存: {path} (code={err})")
                break
        except Exception as e:
            Log(f"[WARN] 保存失败 {path}: {e}")
    if saved is None:
        raise RuntimeError("所有保存路径均失败")
    return saved


def main():
    try:
        os.remove(LOG_FILE)
    except OSError:
        pass
    Log("=== UMI 鱼鳍软爪建模开始 ===")

    dense, quads, holes, stats = build_geometry()
    print_parameters(stats)

    if "--check-only" in sys.argv:
        Log("[INFO] --check-only：仅校验几何，未调用 SolidWorks。")
        return

    sw = connect_solidworks()
    saved = build_part(sw, dense, quads, holes)
    Log(f"\n完成！零件文件: {saved}")


if __name__ == "__main__":
    main()
