# -*- coding: utf-8 -*-
"""
一级圆柱齿轮减速器 - SolidWorks 2024 参数化建模脚本
设计数据来自 reducer_geometry.py（GB 规范：GB/T 1357/10095/1095-1096/276 等）

全部零件用已验证的 API 路径建模：
  前视/上视基准面草图（AddToDB 批量绘制、绝对坐标、无几何约束）
  -> 退出草图 -> 选中草图特征 -> FeatureExtrusion2 盲拉伸
嵌套轮廓自动成孔；轴类零件用多段圆柱叠加（免 revolve）；键槽用带缺口的截面（免 cut）。

用法：
  python build_reducer.py            # 全部零件 + 装配体
  python build_reducer.py --parts    # 只建零件
  python build_reducer.py --asm      # 只建装配体
"""
import math
import os
import sys
import time

import numpy as np

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
import reducer_geometry as rg

_PYLIB = os.path.join(os.path.expanduser("~"), "Documents", "Codex", "2026-08-15", "he-2", "pylibs2")
try:
    import win32com.client
except Exception:
    if os.path.isdir(_PYLIB):
        for _p in (_PYLIB, os.path.join(_PYLIB, "win32"),
                   os.path.join(_PYLIB, "win32", "lib")):
            if _p not in sys.path:
                sys.path.insert(0, _p)
        if hasattr(os, "add_dll_directory"):
            os.add_dll_directory(os.path.join(_PYLIB, "pywin32_system32"))
    import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

TPL_PART = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates\gb_part.prtdot"
TPL_ASM = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates\gb_assembly.asmdot"
OUT_DIR = rg.OUT_DIR
DESKTOP_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "一级减速器")
LOG_FILE = os.path.join(OUT_DIR, "reducer_build_log.txt")
_SW = None


def mm2m(v):
    return float(v) / 1000.0


def Log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


# ===========================================================================
# SW 基础操作
# ===========================================================================
def connect_solidworks():
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        Log("[OK] 已连接正在运行的 SolidWorks 实例")
    except Exception:
        Log("[INFO] 未发现运行中的 SolidWorks，正在启动新实例...")
        sw = win32com.client.Dispatch("SldWorks.Application")
    t0 = time.time()
    while time.time() - t0 < 240:
        try:
            rev = str(sw.RevisionNumber)
            Log(f"[OK] SolidWorks 就绪，版本: {rev}")
            try:
                sw.Visible = True
            except Exception:
                pass
            time.sleep(5.0)   # 等 UI/COM 完全就绪，避免紧接着的批量操作压垮新实例
            return sw
        except Exception:
            time.sleep(1.0)
    raise RuntimeError("SolidWorks 长时间未就绪")


def doc_title(doc):
    try:
        return str(invoke(doc, "GetTitle"))
    except Exception:
        return None


def in_sketch(sk):
    """判断当前是否处于草图编辑状态（无法判断时放行）。"""
    try:
        return sk.ActiveSketch is not None
    except Exception:
        try:
            return invoke(sk, "ActiveSketch") is not None
        except Exception:
            return True


def check_usable(m):
    """校验文档可操作：特征树可遍历即可。
    注意：不检查 ActiveSketch——本机 typelib 该属性不可靠（可能恒非空），
    会导致新文档被误判为“正在编辑草图”而反复新建文档。"""
    if m is None:
        return None
    try:
        if m.FeatureByPositionReverse(0) is None:
            return None
    except Exception:
        return None
    return m


def new_part(sw):
    before = doc_title(sw.ActiveDoc)
    last = None
    for attempt in range(8):
        try:
            m = sw.NewDocument(TPL_PART, 0, 0, 0)
            time.sleep(1.2)
            if check_usable(m) is not None:
                try:
                    m.ClearSelection2(True)
                except Exception:
                    pass
                return m
            cur = doc_title(sw.ActiveDoc)
            if cur and cur != before and cur.startswith("零件"):
                m2 = check_usable(sw.ActiveDoc)
                if m2 is not None:
                    try:
                        m2.ClearSelection2(True)
                    except Exception:
                        pass
                    return m2
                before = cur
        except Exception as e:
            last = e
        Log(f"[WARN] 新建零件文档未就绪，重试 {attempt + 1}/8 ...")
        time.sleep(3.0)
    raise RuntimeError(f"新建零件文档失败: {last}")


def feature_at(model, i):
    try:
        return model.FeatureByPositionReverse(i)
    except Exception:
        return None


def invoke(obj, name, *args):
    """兼容本机 typelib：先按方法调用，若不可调用则按属性读取。"""
    attr = getattr(obj, name)
    try:
        return attr(*args)
    except TypeError:
        return attr


def find_feature(model, substr):
    """从特征树底部（最新）向上找第一个名称含 substr 的特征。"""
    for i in range(0, 150):
        f = feature_at(model, i)
        if f is None:
            break
        try:
            name = str(invoke(f, "Name"))
        except Exception:
            name = ""
        if name and substr in name:
            return f
    return None


def find_plane(model, keys):
    """按名称匹配基准面并尝试 Select2（不依赖 GetTypeName/SelectByID2）。"""
    candidates = []
    for i in range(0, 80):
        try:
            f = model.FeatureByPositionReverse(i)
        except Exception:
            break
        if f is None:
            break
        try:
            name = str(invoke(f, "Name"))
        except Exception:
            name = ""
        if not name:
            continue
        priority = 0 if any(k in name for k in keys) else 1
        candidates.append((priority, name, f))
    candidates.sort(key=lambda c: c[0])
    for priority, name, f in candidates:
        try:
            if f.Select2(False, 0):
                return f
        except Exception:
            pass
    return None


def start_sketch(model, keys=("前视", "Front")):
    model.ClearSelection2(True)
    plane = find_plane(model, keys)
    if plane is None:
        raise RuntimeError(f"找不到基准面 {keys}")
    try:
        pname = str(invoke(plane, "Name"))
    except Exception:
        pname = "?"
    Log(f"[DBG] 选中基准面: {pname}")
    plane.Select2(False, 0)
    sk = model.SketchManager
    last = None
    for attempt in range(5):
        try:
            # 本机 typelib 的 InsertSketch2 返回值不可信：进草图成功也可能返回假值。
            # 与已跑通的鱼鳍脚本一致：只按异常判断，不检查返回值。
            sk.InsertSketch2(True)
            sk.AddToDB = True
            sk.DisplayWhenAdded = False
            return sk
        except Exception as e:
            last = e
        time.sleep(1.5)
    for attempt in range(5):
        try:
            sk.InsertSketch(True)
            sk.AddToDB = True
            sk.DisplayWhenAdded = False
            return sk
        except Exception as e:
            last = e
        time.sleep(1.5)
    raise RuntimeError(f"进入草图失败: {last}")


def draw_poly(sk, pts):
    pts = np.asarray(pts, dtype=float)
    n = len(pts)
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        sk.CreateLine(mm2m(x0), mm2m(y0), 0, mm2m(x1), mm2m(y1), 0)


def draw_circle(sk, cx, cy, r):
    sk.CreateCircleByRadius(mm2m(cx), mm2m(cy), 0, mm2m(r))


def extrude(model, depth, flip=False):
    """退出草图 -> 选中最新草图特征 -> 盲拉伸。退出后不要 ClearSelection！"""
    sk = model.SketchManager
    sk.AddToDB = False
    sk.DisplayWhenAdded = True
    try:
        sk.InsertSketch2(True)
    except Exception:
        sk.InsertSketch(True)
    feat = find_feature(model, "草图")
    if feat is None:
        feat = find_feature(model, "Sketch")
    if feat is None:
        raise RuntimeError("找不到草图特征")
    model.ClearSelection2(True)
    feat.Select2(False, 0)
    time.sleep(0.15)
    ext = model.FeatureManager.FeatureExtrusion2(
        True, flip, False, 0, 0, mm2m(depth), 0,
        False, False, False, False, 0, 0,
        False, False, False, False, True, True, True, 0, 0, False)
    if ext is None:
        raise RuntimeError(f"拉伸失败 depth={depth} flip={flip}")
    time.sleep(0.2)
    return ext


def save_part(model, fname):
    os.makedirs(DESKTOP_DIR, exist_ok=True)
    path = os.path.join(DESKTOP_DIR, fname)
    if os.path.exists(path) and os.path.getsize(path) > 0:
        Log(f"[SKIP] 零件已存在: {path}")
        return path
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
    err = model.SaveAs3(path, 0, 1)
    if not os.path.exists(path):
        raise RuntimeError(f"保存失败 {path} err={err}")
    Log(f"[OK] 已保存: {path}")
    try:
        title = doc_title(model)
        if title:
            _SW.CloseDoc(title)
    except Exception:
        pass
    return path


# ===========================================================================
# 零件构建
# ===========================================================================
def build_base(sw):
    m = new_part(sw)
    Log("[步骤] 箱座: 主体+内腔+凸缘孔...")
    sk = start_sketch(m)
    draw_poly(sk, [(-59.5, 152), (-59.5, 140), (-47.5, 140), (-47.5, 20),
                   (247.5, 20), (247.5, 140), (259.5, 140), (259.5, 152)])
    draw_poly(sk, [(-39.5, 140), (-39.5, 22), (239.5, 22), (239.5, 140)])
    for x in rg.FLANGE_BOLT_X:
        draw_circle(sk, x, 146, 5.5)
    for x in {p[0] for p in rg.PIN_POS}:
        draw_circle(sk, x, 146, rg.PIN_D / 2)
    extrude(m, rg.BOX_Z)

    Log("[步骤] 箱座: 底座凸缘+地脚孔...")
    sk = start_sketch(m)
    draw_poly(sk, [(-72.5, 0), (-72.5, 20), (272.5, 20), (272.5, 0)])
    for x in {p[0] for p in rg.BASE_BOLT_POS}:
        draw_circle(sk, x, 10, 9)
    extrude(m, rg.BASE_Z)

    Log("[步骤] 箱座: 轴承座（高速×2 / 低速×2）...")
    for cx in (0, rg.A):
        dd = rg.SEAT_D2[cx]
        bo = rg.SEAT_BORE[cx]
        n, rr = rg.SEAT_SCR[cx]
        for _side in ("front", "back"):
            sk = start_sketch(m)
            draw_circle(sk, cx, rg.H_CENTER, dd / 2)
            draw_circle(sk, cx, rg.H_CENTER, bo / 2)
            for k in range(n):
                ang = math.pi / 4 + 2 * math.pi * k / n
                draw_circle(sk, cx + rr * math.cos(ang),
                            rg.H_CENTER + rr * math.sin(ang), 4.5)
            if _side == "front":
                extrude(m, rg.SEAT_PRO, flip=True)
            else:
                extrude(m, rg.BOX_Z + rg.SEAT_PRO)

    Log("[步骤] 箱座: 放油孔+油标孔...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_circle(sk, rg.OIL_DRAIN["x"], rg.OIL_DRAIN["z"], rg.OIL_DRAIN["d"] / 2)
    extrude(m, 40)
    sk = start_sketch(m)
    draw_circle(sk, rg.OIL_GAUGE["x"], rg.OIL_GAUGE["y"], rg.OIL_GAUGE["d"] / 2)
    extrude(m, rg.DELTA)
    return save_part(m, "箱座.SLDPRT")


def build_cover(sw):
    off = 128.0
    m = new_part(sw)
    Log("[步骤] 箱盖: 主体+内腔(含视孔)+凸缘孔...")
    sk = start_sketch(m)
    draw_poly(sk, [(-59.5, 62.5), (-59.5, 12), (-47.5, 12), (-47.5, 0),
                   (247.5, 0), (247.5, 12), (259.5, 12), (259.5, 62.5)])
    draw_poly(sk, [(-39.5, 54.5), (-39.5, 12), (239.5, 12), (239.5, 54.5)])
    for x in rg.FLANGE_BOLT_X:
        draw_circle(sk, x, 6, 5.5)
    for x in {p[0] for p in rg.PIN_POS}:
        draw_circle(sk, x, 6, rg.PIN_D / 2)
    extrude(m, rg.BOX_Z)

    Log("[步骤] 箱盖: 顶部视孔框架（左片/右片）...")
    sk = start_sketch(m)
    draw_poly(sk, [(-47.5, 62.5), (-47.5, 54.5), (60, 54.5), (60, 62.5)])
    extrude(m, rg.BOX_Z)
    sk = start_sketch(m)
    draw_poly(sk, [(160, 62.5), (160, 54.5), (247.5, 54.5), (247.5, 62.5)])
    extrude(m, rg.BOX_Z)

    Log("[步骤] 箱盖: 轴承座...")
    for cx in (0, rg.A):
        dd = rg.SEAT_D2[cx]
        bo = rg.SEAT_BORE[cx]
        n, rr = rg.SEAT_SCR[cx]
        for _side in ("front", "back"):
            sk = start_sketch(m)
            draw_circle(sk, cx, 12, dd / 2)
            draw_circle(sk, cx, 12, bo / 2)
            for k in range(n):
                ang = math.pi / 4 + 2 * math.pi * k / n
                draw_circle(sk, cx + rr * math.cos(ang), 12 + rr * math.sin(ang), 4.5)
            if _side == "front":
                extrude(m, rg.SEAT_PRO, flip=True)
            else:
                extrude(m, rg.BOX_Z + rg.SEAT_PRO)
    return save_part(m, "箱盖.SLDPRT")


def build_gear(sw, z, m, bore, light_holes=None, width=50, hub=None):
    mdl = new_part(sw)
    Log(f"[步骤] 齿轮 z={z} m={m}...")
    sk = start_sketch(mdl)
    draw_poly(sk, rg.involute_gear_profile(z, m))
    draw_poly(sk, bore)
    if light_holes:
        for hx, hy, hr in light_holes:
            draw_circle(sk, hx, hy, hr)
    extrude(mdl, width)
    if hub:
        sk = start_sketch(mdl)
        draw_circle(sk, 0, 0, hub["od"] / 2)
        draw_poly(sk, bore)
        extrude(mdl, hub["total"] + hub["back"])
        sk = start_sketch(mdl)
        draw_circle(sk, 0, 0, hub["od"] / 2)
        draw_poly(sk, bore)
        extrude(mdl, hub["back"], flip=True)
    return mdl


def build_pinion(sw):
    m = build_gear(sw, rg.Z1, rg.M, rg.GEAR1_BORE, width=rg.B1)
    return save_part(m, "小齿轮.SLDPRT")


def build_big_gear(sw):
    holes = []
    for k in range(6):
        ang = 2 * math.pi * k / 6 + math.pi / 6
        holes.append((70 * math.cos(ang), 70 * math.sin(ang), 17.5))
    m = build_gear(sw, rg.Z2, rg.M, rg.GEAR2_BORE, light_holes=holes,
                   width=rg.B2, hub=dict(od=70, total=60, back=10))
    return save_part(m, "大齿轮.SLDPRT")


def build_shaft(sw, halfs, name):
    m = new_part(sw)
    Log(f"[步骤] {name}...")
    for seg in halfs:
        r = seg["rad"]
        h = seg["half"]
        key = seg.get("key")
        sk = start_sketch(m)
        if key:
            draw_poly(sk, rg.shaft_notch_profile(r, key[0], key[2]))
        else:
            draw_poly(sk, rg.circle_profile(r))
        extrude(m, h)
        sk = start_sketch(m)
        if key:
            draw_poly(sk, rg.shaft_notch_profile(r, key[0], key[2]))
        else:
            draw_poly(sk, rg.circle_profile(r))
        extrude(m, h, flip=True)
    return save_part(m, name)


def build_bearing(sw, d, D, B, name, balls=10, ball_d=8.0):
    m = new_part(sw)
    Log(f"[步骤] 轴承 {name}...")
    ring = (D - d) / 4.0
    sk = start_sketch(m)
    draw_circle(sk, 0, 0, D / 2)
    draw_circle(sk, 0, 0, D / 2 - ring)
    extrude(m, B)
    sk = start_sketch(m)
    draw_circle(sk, 0, 0, d / 2 + ring)
    draw_circle(sk, 0, 0, d / 2)
    extrude(m, B)
    pitch = (D + d) / 4.0
    for k in range(balls):
        ang = 2 * math.pi * k / balls
        sk = start_sketch(m)
        draw_circle(sk, pitch * math.cos(ang), pitch * math.sin(ang), ball_d / 2)
        extrude(m, B)
    return save_part(m, name)


def build_cap(sw, key, side, through, name):
    """凸缘式轴承端盖。side=front: 止口向 +Z、凸缘朝 -Z；side=back 相反。"""
    c = rg.CAP[key]
    m = new_part(sw)
    Log(f"[步骤] 端盖 {name}...")
    disc_flip = (side == "front")
    lip_flip = (side == "back")
    sk = start_sketch(m)
    draw_circle(sk, 0, 0, c["D3"] / 2)
    for k in range(c["scr"]):
        ang = math.pi / 4 + 2 * math.pi * k / c["scr"]
        draw_circle(sk, c["r"] * math.cos(ang), c["r"] * math.sin(ang), 4.5)
    if through:
        draw_circle(sk, 0, 0, c["hole"] / 2)
    extrude(m, c["th"], flip=disc_flip)
    sk = start_sketch(m)
    draw_circle(sk, 0, 0, c["D0"] / 2)
    if through:
        draw_circle(sk, 0, 0, c["hole"] / 2)
    extrude(m, c["lip"], flip=lip_flip)
    return save_part(m, name)


def build_key(sw, L, b, h, name):
    m = new_part(sw)
    Log(f"[步骤] 键 {name}...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_poly(sk, rg.stadium_profile(L, b))
    extrude(m, h)
    return save_part(m, name)


def build_pin(sw):
    m = new_part(sw)
    Log("[步骤] 定位销...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_circle(sk, 0, 0, rg.PIN_D / 2)
    extrude(m, 40)
    return save_part(m, "定位销.SLDPRT")


def build_sight_cover(sw):
    m = new_part(sw)
    Log("[步骤] 视孔盖...")
    sc = rg.SIGHT_COVER
    sk = start_sketch(m, ("上视", "Top"))
    draw_poly(sk, [(-sc["dx"], -sc["dz"]), (-sc["dx"], sc["dz"]),
                   (sc["dx"], sc["dz"]), (sc["dx"], -sc["dz"])])
    for sx, sz in ((-50, -35), (50, -35), (-50, 35), (50, 35)):
        draw_circle(sk, sx, sz, 3.25)
    draw_circle(sk, 0, 0, rg.VENT["d"] / 2)
    extrude(m, sc["th"])
    return save_part(m, "视孔盖.SLDPRT")


def build_vent(sw):
    m = new_part(sw)
    Log("[步骤] 通气器...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_poly(sk, rg.hex_profile(rg.VENT["hex_s"]))
    extrude(m, rg.VENT["hex_h"])
    sk = start_sketch(m, ("上视", "Top"))
    draw_circle(sk, 0, 0, rg.VENT["d"] / 2)
    extrude(m, rg.VENT["shank"], flip=True)
    return save_part(m, "通气器.SLDPRT")


def build_gauge(sw):
    m = new_part(sw)
    Log("[步骤] 油标...")
    sk = start_sketch(m)
    draw_circle(sk, 0, 0, 20)
    for k in range(3):
        ang = 2 * math.pi * k / 3 + math.pi / 6
        draw_circle(sk, 15 * math.cos(ang), 15 * math.sin(ang), 2)
    extrude(m, 6, flip=True)
    return save_part(m, "油标.SLDPRT")


def build_plug(sw):
    m = new_part(sw)
    Log("[步骤] 放油螺塞...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_circle(sk, 0, 0, rg.PLUG["d"] / 2)
    extrude(m, rg.PLUG["shank"])
    sk = start_sketch(m, ("上视", "Top"))
    draw_poly(sk, rg.hex_profile(rg.PLUG["hex_s"]))
    extrude(m, rg.PLUG["hex_h"], flip=True)
    return save_part(m, "放油螺塞.SLDPRT")


def build_bolt16(sw):
    m = new_part(sw)
    Log("[步骤] 螺栓 M16（地脚）...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_circle(sk, 0, 0, 8)
    extrude(m, 60)
    sk = start_sketch(m, ("上视", "Top"))
    draw_poly(sk, rg.hex_profile(24))
    extrude(m, 15, flip=True)
    return save_part(m, "螺栓M16.SLDPRT")


def build_bolt10(sw):
    m = new_part(sw)
    Log("[步骤] 螺栓 M10（箱盖连接）...")
    sk = start_sketch(m, ("上视", "Top"))
    draw_circle(sk, 0, 0, 5)
    extrude(m, 48, flip=True)
    sk = start_sketch(m, ("上视", "Top"))
    draw_poly(sk, rg.hex_profile(16))
    extrude(m, 8)
    return save_part(m, "螺栓M10.SLDPRT")


def build_screw8(sw, side):
    m = new_part(sw)
    name = f"螺钉M8{side}.SLDPRT"
    Log(f"[步骤] {name}...")
    sk = start_sketch(m)
    draw_poly(sk, rg.hex_profile(13))
    extrude(m, 5, flip=(side == "前"))
    sk = start_sketch(m)
    draw_circle(sk, 0, 0, 4)
    extrude(m, 20, flip=(side == "前"))
    return save_part(m, name)


def build_all_parts(sw):
    global _SW
    _SW = sw
    files = {}
    tasks = [
        ("箱座", "箱座.SLDPRT", lambda: build_base(_SW)),
        ("箱盖", "箱盖.SLDPRT", lambda: build_cover(_SW)),
        ("小齿轮", "小齿轮.SLDPRT", lambda: build_pinion(_SW)),
        ("大齿轮", "大齿轮.SLDPRT", lambda: build_big_gear(_SW)),
        ("高速轴", "高速轴.SLDPRT", lambda: build_shaft(_SW, rg.SHAFT1_HALF, "高速轴.SLDPRT")),
        ("低速轴", "低速轴.SLDPRT", lambda: build_shaft(_SW, rg.SHAFT2_HALF, "低速轴.SLDPRT")),
        ("轴承6205", "轴承6205.SLDPRT", lambda: build_bearing(_SW, 25, 52, 15, "轴承6205.SLDPRT", 10, 8.0)),
        ("轴承6208", "轴承6208.SLDPRT", lambda: build_bearing(_SW, 40, 80, 18, "轴承6208.SLDPRT", 12, 12.0)),
        ("透盖_高速前", "透盖_高速前.SLDPRT", lambda: build_cap(_SW, "hs", "front", True, "透盖_高速前.SLDPRT")),
        ("透盖_高速后", "透盖_高速后.SLDPRT", lambda: build_cap(_SW, "hs", "back", True, "透盖_高速后.SLDPRT")),
        ("闷盖_高速前", "闷盖_高速前.SLDPRT", lambda: build_cap(_SW, "hs", "front", False, "闷盖_高速前.SLDPRT")),
        ("闷盖_高速后", "闷盖_高速后.SLDPRT", lambda: build_cap(_SW, "hs", "back", False, "闷盖_高速后.SLDPRT")),
        ("透盖_低速前", "透盖_低速前.SLDPRT", lambda: build_cap(_SW, "ls", "front", True, "透盖_低速前.SLDPRT")),
        ("透盖_低速后", "透盖_低速后.SLDPRT", lambda: build_cap(_SW, "ls", "back", True, "透盖_低速后.SLDPRT")),
        ("闷盖_低速前", "闷盖_低速前.SLDPRT", lambda: build_cap(_SW, "ls", "front", False, "闷盖_低速前.SLDPRT")),
        ("闷盖_低速后", "闷盖_低速后.SLDPRT", lambda: build_cap(_SW, "ls", "back", False, "闷盖_低速后.SLDPRT")),
        ("键8x7x40", "键8x7x40.SLDPRT", lambda: build_key(_SW, 40, 8, 7, "键8x7x40.SLDPRT")),
        ("键6x6x32", "键6x6x32.SLDPRT", lambda: build_key(_SW, 32, 6, 6, "键6x6x32.SLDPRT")),
        ("键14x9x40", "键14x9x40.SLDPRT", lambda: build_key(_SW, 40, 14, 9, "键14x9x40.SLDPRT")),
        ("键10x8x50", "键10x8x50.SLDPRT", lambda: build_key(_SW, 50, 10, 8, "键10x8x50.SLDPRT")),
        ("定位销", "定位销.SLDPRT", lambda: build_pin(_SW)),
        ("视孔盖", "视孔盖.SLDPRT", lambda: build_sight_cover(_SW)),
        ("通气器", "通气器.SLDPRT", lambda: build_vent(_SW)),
        ("油标", "油标.SLDPRT", lambda: build_gauge(_SW)),
        ("放油螺塞", "放油螺塞.SLDPRT", lambda: build_plug(_SW)),
        ("螺栓M16", "螺栓M16.SLDPRT", lambda: build_bolt16(_SW)),
        ("螺栓M10", "螺栓M10.SLDPRT", lambda: build_bolt10(_SW)),
        ("螺钉M8前", "螺钉M8前.SLDPRT", lambda: build_screw8(_SW, "前")),
        ("螺钉M8后", "螺钉M8后.SLDPRT", lambda: build_screw8(_SW, "后")),
    ]
    for name, fname, fn in tasks:
        if os.path.exists(os.path.join(DESKTOP_DIR, fname)):
            Log(f"[SKIP] 零件已存在: {fname}")
            files[name] = os.path.join(DESKTOP_DIR, fname)
            continue
        try:
            try:
                files[name] = fn()
            except Exception as e:
                if ("RPC" in str(e)) or ("ActiveDoc" in str(e)):
                    Log(f"[WARN] {name} 触发断线，重连 SolidWorks 后重试...")
                    try:
                        _SW = connect_solidworks()
                    except Exception as e2:
                        Log(f"[ERROR] 重连失败: {e2}")
                    files[name] = fn()
                else:
                    raise
        except Exception as e:
            Log(f"[ERROR] 零件 {name} 失败: {e}")
    Log(f"[OK] 成功 {len(files)}/{len(tasks)} 个零件已保存到 {DESKTOP_DIR}")
    return files


# ===========================================================================
# 装配体
# ===========================================================================
def _cap_pos(side):
    """端盖安装面：前 z=-60，后 z=161。"""
    z = -60 if side == "前" else 161
    return [(0, 140, z), (125, 140, z)]


def _screw_pos(side, cx, n, rr, z):
    out = []
    for k in range(n):
        ang = math.pi / 4 + 2 * math.pi * k / n
        out.append((cx + rr * math.cos(ang), 140 + rr * math.sin(ang), z))
    return out


def build_assembly(sw, files):
    Log("[步骤] 创建装配体...")
    # 1) 先打开所有零件（本机 typelib 的 AddComponent 要求零件已打开）
    def open_all_parts():
        opened = 0
        for fname in sorted(os.listdir(DESKTOP_DIR)):
            if not fname.lower().endswith(".sldprt"):
                continue
            if fname.startswith("~$"):
                continue
            p = os.path.join(DESKTOP_DIR, fname)
            try:
                sw.OpenDoc(p, 1)
                opened += 1
            except Exception as e:
                Log(f"[WARN] 打开零件失败 {fname}: {e}")
            time.sleep(0.3)
        return opened

    try:
        opened = open_all_parts()
    except Exception as e:
        Log(f"[WARN] 批量打开中断（{e}），重连后重试...")
        sw2 = connect_solidworks()
        opened = open_all_parts()
    Log(f"[OK] 已打开 {opened} 个零件")

    # 2) 新建装配体文档
    before = doc_title(sw.ActiveDoc)
    m = sw.NewDocument(TPL_ASM, 0, 0, 0)
    time.sleep(1.5)
    cur = doc_title(m) or doc_title(sw.ActiveDoc)
    if "装配体" not in (cur or ""):
        m = sw.ActiveDoc
        cur = doc_title(m) or ""
    if "装配体" not in cur:
        raise RuntimeError(f"装配体文档创建失败 (ActiveDoc={cur}, before={before})")
    Log(f"[OK] 装配体文档: {cur}")

    def add(fname, x, y, z):
        path = os.path.join(DESKTOP_DIR, fname)
        if not os.path.exists(path):
            Log(f"[WARN] 装配跳过缺失零件: {fname}")
            return
        try:
            ok = m.AddComponent(path, mm2m(x), mm2m(y), mm2m(z))
            if not ok:
                Log(f"[WARN] AddComponent 失败: {fname}")
        except Exception as e:
            Log(f"[WARN] AddComponent 异常 {fname}: {e}")
        time.sleep(0.05)

    add("箱座.SLDPRT", 0, 0, 0)
    add("箱盖.SLDPRT", 0, 128, 0)
    add("小齿轮.SLDPRT", 0, 140, -27.5)
    add("大齿轮.SLDPRT", 125, 140, -25)
    add("高速轴.SLDPRT", 0, 140, 0)
    add("低速轴.SLDPRT", 125, 140, 0)
    add("轴承6205.SLDPRT", 0, 140, 27.5)
    add("轴承6205.SLDPRT", 0, 140, -42.5)
    add("轴承6208.SLDPRT", 125, 140, 35)
    add("轴承6208.SLDPRT", 125, 140, -53)
    for side in ("前", "后"):
        for (cx, cy, cz) in _cap_pos(side):
            add(f"透盖_高速{side}.SLDPRT", cx, cy, cz)
            add(f"闷盖_高速{side}.SLDPRT", cx, cy, cz)
            add(f"透盖_低速{side}.SLDPRT", cx, cy, cz)
            add(f"闷盖_低速{side}.SLDPRT", cx, cy, cz)
    # 键（轴键槽底面为局部 y=0）
    add("键8x7x40.SLDPRT", 0, 151, 0)
    add("键6x6x32.SLDPRT", 0, 146.5, 87.5)
    add("键6x6x32.SLDPRT", 0, 146.5, -87.5)
    add("键14x9x40.SLDPRT", 125, 157, 0)
    add("键10x8x50.SLDPRT", 125, 151, 103)
    add("键10x8x50.SLDPRT", 125, 151, -103)
    for (px, pz) in rg.PIN_POS:
        add("定位销.SLDPRT", px, 120, pz)
    add("视孔盖.SLDPRT", 110, 190.5, 0)
    add("通气器.SLDPRT", 110, 196.5, 0)
    add("油标.SLDPRT", 170, 55, 0)
    add("放油螺塞.SLDPRT", rg.OIL_DRAIN["x"], 0, rg.OIL_DRAIN["z"])
    for (bx, bz) in rg.BASE_BOLT_POS:
        add("螺栓M16.SLDPRT", bx, 0, bz)
    for x in rg.FLANGE_BOLT_X:
        for z in rg.FLANGE_BOLT_Z:
            add("螺栓M10.SLDPRT", x, 152, z)
    for (sx, sy, sz) in _screw_pos("前", 0, 4, 37, -60):
        add("螺钉M8前.SLDPRT", sx, sy, sz)
    for (sx, sy, sz) in _screw_pos("前", 125, 6, 52, -60):
        add("螺钉M8前.SLDPRT", sx, sy, sz)
    for (sx, sy, sz) in _screw_pos("后", 0, 4, 37, 161):
        add("螺钉M8后.SLDPRT", sx, sy, sz)
    for (sx, sy, sz) in _screw_pos("后", 125, 6, 52, 161):
        add("螺钉M8后.SLDPRT", sx, sy, sz)

    # 3) 关闭旧的同名装配体（若打开）并保存
    try:
        sw.CloseDoc("一级减速器.SLDASM")
    except Exception:
        pass
    asm_path = os.path.join(DESKTOP_DIR, "一级减速器.SLDASM")
    if os.path.exists(asm_path):
        try:
            os.remove(asm_path)
        except OSError:
            pass
    err = m.SaveAs3(asm_path, 0, 1)
    if not os.path.exists(asm_path):
        raise RuntimeError(f"装配体保存失败 err={err}")
    Log(f"[OK] 装配体已保存: {asm_path}")
    return asm_path


def main():
    global _SW
    try:
        os.remove(LOG_FILE)
    except OSError:
        pass
    Log("=== 一级圆柱齿轮减速器建模开始 ===")
    Log(f"设计: P={rg.P_RAW}kW n1={rg.N1:.0f} i={rg.I_RATIO} a={rg.A} "
        f"m={rg.M} z1={rg.Z1} z2={rg.Z2} 输出目录={DESKTOP_DIR}")
    if "--check" in sys.argv:
        Log("[INFO] --check：仅打印设计参数")
        return
    sw = connect_solidworks()
    only_asm = "--asm" in sys.argv
    only_parts = "--parts" in sys.argv
    files = {}
    if not only_asm:
        files = build_all_parts(sw)
    if not only_parts:
        asm = build_assembly(_SW, files)
        Log(f"\n完成！装配体: {asm}")
        import shutil
        for f in ("reducer_design_report.md", "reducer_design.json",
                  "reducer_preview.png"):
            src = os.path.join(OUT_DIR, f)
            if os.path.exists(src):
                try:
                    shutil.copy(src, os.path.join(DESKTOP_DIR, f))
                    Log(f"[OK] 设计文件已复制: {f}")
                except Exception as e:
                    Log(f"[WARN] 复制 {f} 失败: {e}")
    else:
        Log(f"\n完成！零件目录: {DESKTOP_DIR}")


if __name__ == "__main__":
    main()
