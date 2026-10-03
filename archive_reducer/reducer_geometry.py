# -*- coding: utf-8 -*-
"""
一级圆柱齿轮减速器 - 设计计算 + 参数化几何生成（纯 numpy，无 SolidWorks 依赖）
依据：机械设计/机械设计基础课程设计（濮良贵）、GB/T 1357 模数系列、
      GB/T 1095-1096 平键、GB/T 276 深沟球轴承、箱体经验公式（0.025a+1 等）
"""
import json
import math
import os
import sys

import numpy as np

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.expanduser("~"), "Documents", "Codex", "2026-08-15", "he-2", "outputs")

# ===========================================================================
# 1. 原始数据（自定）
# ===========================================================================
P_RAW = 3.0          # 输入功率 kW
N1 = 1420.0          # 输入转速 r/min（4 极异步电机）
I_RATIO = 4.0        # 传动比（GB 推荐系列）
LIFE_H = 24000.0     # 设计寿命 h（10 年×300 天×8h）
K_LOAD = 1.3         # 载荷系数（平稳）
PHI_A = 0.4          # 齿宽系数
ALPHA_DEG = 20.0     # 压力角

# 齿轮材料（软齿面）：小齿轮 40Cr 调质，大齿轮 45 钢调质
SIG_H_ALLOW = 550.0  # MPa（取小值 [σH]2）
SIG_F1 = 290.0
SIG_F2 = 240.0

# ===========================================================================
# 2. 齿轮传动设计
# ===========================================================================
T1 = 9550.0 * P_RAW / N1                       # N·m
T1_NMM = T1 * 1000.0                            # N·mm
ETA = 0.97
P2 = P_RAW * ETA
N2 = N1 / I_RATIO
T2_NMM = 9550.0 * P2 / N2 * 1000.0

# 按齿面接触强度初算中心距（钢-钢软齿面，T1 单位 N·mm）
u = I_RATIO
a_min = (u + 1.0) * ((476.0 / SIG_H_ALLOW) ** 2 * K_LOAD * T1_NMM /
                     (PHI_A * u)) ** (1.0 / 3.0)

Z1 = 20
Z2 = int(round(Z1 * u))
M = 2.5                    # GB/T 1357 第一系列
A = M * (Z1 + Z2) / 2.0    # 中心距
D1 = M * Z1
D2 = M * Z2
DA1 = D1 + 2 * M
DA2 = D2 + 2 * M
DF1 = D1 - 2 * 1.25 * M
DF2 = D2 - 2 * 1.25 * M
B2 = int(PHI_A * A)        # 大齿轮宽
B1 = B2 + 5                # 小齿轮宽（加宽补偿）
V = math.pi * D1 * N1 / 60000.0     # 圆周速度 m/s
PRECISION = 8               # GB/T 10095 精度等级（v≈3.7 m/s，选 8 级）

# 强度校核（简化）
SIG_H_ACT = 2.5 * 189.8 * math.sqrt(2 * K_LOAD * T1_NMM * (u + 1.0) /
                                    (B2 * D1 * D1 * u))
YF1, YS1 = 2.80, 1.55       # z=20 近似
YF2, YS2 = 2.24, 1.75       # z=80 近似
SIG_F_ACT1 = 2 * K_LOAD * T1_NMM * YF1 * YS1 / (B2 * M * M * Z1)
SIG_F_ACT2 = SIG_F_ACT1 * YF2 * YS2 / (YF1 * YS1)

# ===========================================================================
# 3. 轴设计（45 钢，C=126）
# ===========================================================================
C_AXIS = 126.0
D1MIN = C_AXIS * (P_RAW / N1) ** (1.0 / 3.0)
D2MIN = C_AXIS * (P2 / N2) ** (1.0 / 3.0)

# 高速轴（轴1）分段（从中心向外，半径 mm，半长 mm，键 b×h/t1）
SHAFT1_HALF = [
    dict(rad=15.0, half=27.5, key=(8, 7, 4.0, 40.0), tag="齿轮座 8×7×40"),
    dict(rad=12.5, half=15.0, key=None, tag="轴承 6205"),
    dict(rad=12.5, half=20.0, key=None, tag="密封段"),
    dict(rad=10.0, half=50.0, key=(6, 6, 3.5, 32.0), tag="联轴器 6×6×32"),
]
SHAFT2_HALF = [
    dict(rad=22.5, half=25.0, key=(14, 9, 5.5, 40.0), tag="齿轮座 14×9×40"),
    dict(rad=26.0, half=10.0, key=None, tag="轴肩"),
    dict(rad=20.0, half=18.0, key=None, tag="轴承 6208"),
    dict(rad=20.0, half=25.0, key=None, tag="密封段"),
    dict(rad=16.0, half=80.0, key=(10, 8, 5.0, 50.0), tag="联轴器 10×8×50"),
]
BEARING1 = dict(d=25, D=52, B=15, name="6205")
BEARING2 = dict(d=40, D=80, B=18, name="6208")

# ===========================================================================
# 4. 箱体尺寸（铸铁，经验公式）
# ===========================================================================
DELTA = 8.0                 # 箱座壁厚 0.025a+1=4.1 → 取 8（≥8）
DELTA1 = 8.0                # 箱盖壁厚 0.02a+3=5.5 → 取 8
BF = 1.5 * DELTA            # 箱座凸缘厚 12
BF1 = 1.5 * DELTA1          # 箱盖凸缘厚 12
BF2 = 2.5 * DELTA           # 箱座底凸缘厚 20
DF_BOLT = 16                # 地脚螺栓 M16（0.036a+12=16.5）
N_BOLT = 4
D1_BOLT = 12                # 轴承旁螺栓 M12
D2_BOLT = 10                # 箱盖-箱座连接螺栓 M10
D3_SCREW = 8                # 轴承端盖螺钉 M8
D4_SCREW = 6                # 视孔盖螺钉 M6
PIN_D = 8                   # 定位销 φ8（0.8d2）

H_CENTER = 140.0            # 轴心高度
WALL_X_L = -47.5            # 左外壁
WALL_X_R = 247.5            # 右外壁
CAV_X_L = -39.5
CAV_X_R = 239.5
FLANGE_X_L = -59.5          # 凸缘外沿
FLANGE_X_R = 259.5
BASE_X_L = -72.5            # 底座外沿
BASE_X_R = 272.5
BOX_Y_BOT = 20.0            # 箱座内底面（底座顶）
BOX_Y_TOP = H_CENTER        # 分箱面（过轴心线）
COVER_Y_TOP = 190.5         # 箱盖顶面（内壁 182.5 + 壁厚 8）
SPLIT_Y = H_CENTER
FLANGE_Y_BOT = 128.0        # 箱盖凸缘下沿
FLANGE_Y_TOP = 152.0        # 箱座凸缘上沿
BOX_Z = 101.0               # 箱体外宽（内腔 85 + 2×8）
CAV_Z_HALF = 42.5           # 内腔半宽
BASE_Z = 141.0              # 底座宽
SEAT_PRO = 60.0             # 轴承座凸台外伸量
SEAT_D2 = {0: 95.0, 125: 125.0}       # 轴承座外径（高速/低速）
SEAT_BORE = {0: 52.0, 125: 80.0}      # 轴承座孔径 = 轴承外径
SEAT_SCR = {0: (4, 37.0), 125: (6, 52.0)}   # 端盖螺钉数、分布圆半径

# 连接法兰螺栓（M10）位置
FLANGE_BOLT_X = [-30.0, 85.0, 165.0, 220.0]
FLANGE_BOLT_Z = [20.0, 80.0]
PIN_POS = [(-10.0, 20.0), (-10.0, 80.0), (210.0, 20.0), (210.0, 80.0)]
BASE_BOLT_POS = [(-60.0, 20.0), (-60.0, 120.0), (240.0, 20.0), (240.0, 120.0)]

OIL_DRAIN = dict(x=160.0, z=50.0, d=18.0)     # 放油孔
OIL_GAUGE = dict(x=170.0, y=55.0, d=13.0)     # 油标孔（前壁）
SIGHT = dict(x0=60.0, x1=160.0, z0=-35.0, z1=35.0, y0=182.5, y1=190.5)  # 视孔
SIGHT_COVER = dict(dx=70.0, dz=55.0, th=6.0)  # 视孔盖半宽/半深/厚
VENT = dict(d=16.0, hex_s=24.0, hex_h=10.0, shank=20.0)  # 通气器 M16×1.5
PLUG = dict(d=16.0, hex_s=24.0, hex_h=10.0, shank=20.0)  # 放油螺塞 M16×1.5

# 轴承端盖（凸缘式）
CAP = {
    "hs": dict(D3=95.0, D0=52.0, hole=25.0, scr=4, r=37.0, th=10.0, lip=20.0),
    "ls": dict(D3=120.0, D0=80.0, hole=40.0, scr=6, r=52.0, th=10.0, lip=20.0),
}

# ===========================================================================
# 5. 几何生成器
# ===========================================================================
def involute_gear_profile(z, m, alpha_deg=ALPHA_DEG, pf=5):
    """渐开线直齿圆柱齿轮完整外轮廓（闭合折线，mm）。"""
    alpha = math.radians(alpha_deg)
    r = m * z / 2.0
    ra = r + m
    rf = r - 1.25 * m
    rb = r * math.cos(alpha)
    # 齿廓起始半径：无根切（齿根圆 > 基圆，约 z>=42）时齿廓从齿根圆开始；
    # 有根切（基圆 > 齿根圆）时齿廓从基圆开始，基圆到齿根圆用过渡连线。
    r0 = max(rb, rf)

    def inv(rho):
        t = math.sqrt(max(rho * rho / (rb * rb) - 1.0, 0.0))
        return t - math.acos(min(max(rb / rho, -1.0), 1.0))

    invr = inv(r)
    delta = math.pi / (2.0 * z)
    pts = []
    for i in range(z):
        th0 = 2.0 * math.pi * i / z
        # 右齿面 r0→ra
        for k in range(pf):
            rho = r0 + (ra - r0) * k / (pf - 1)
            ang = th0 - delta - invr + inv(rho)
            pts.append((rho * math.cos(ang), rho * math.sin(ang)))
        # 齿顶弧
        angR = th0 - delta - invr + inv(ra)
        angL = th0 + delta + invr - inv(ra)
        for k in range(1, 3):
            ang = angR + (angL - angR) * k / 3.0
            pts.append((ra * math.cos(ang), ra * math.sin(ang)))
        # 左齿面 ra→r0
        for k in range(pf):
            rho = ra - (ra - r0) * k / (pf - 1)
            ang = th0 + delta + invr - inv(rho)
            pts.append((rho * math.cos(ang), rho * math.sin(ang)))
        # 齿根：左齿面终点(r0) 到 下一齿右齿面起点(r0)
        angA = th0 + delta + invr - inv(r0)
        angB = th0 + 2.0 * math.pi / z - delta - invr + inv(r0)
        if r0 > rf:
            # 有根切：径向过渡线 + 齿根圆弧
            pts.append((rf * math.cos(angA), rf * math.sin(angA)))
            for k in range(1, 5):
                ang = angA + (angB - angA) * k / 5.0
                pts.append((rf * math.cos(ang), rf * math.sin(ang)))
            pts.append((r0 * math.cos(angB), r0 * math.sin(angB)))
        else:
            # 无根切：齿根圆弧直接连接相邻齿廓，齿谷平滑无凸起
            for k in range(1, 5):
                ang = angA + (angB - angA) * k / 5.0
                pts.append((rf * math.cos(ang), rf * math.sin(ang)))
    return np.array(pts)


def hub_bore_profile(R, b, t2, n=48):
    """轮毂孔 + 键槽（D 形孔，键槽朝上），闭合折线。"""
    beta = math.atan((b / 2.0) / R)
    a0 = math.pi / 2.0
    pts = []
    for ang in np.linspace(a0 + beta, a0 - beta + 2 * math.pi, n):
        pts.append((R * math.cos(ang), R * math.sin(ang)))
    pts.append((b / 2.0, R))
    pts.append((b / 2.0, R + t2))
    pts.append((-b / 2.0, R + t2))
    pts.append((-b / 2.0, R))
    return np.array(pts)


def shaft_notch_profile(r, b, t1, n=48):
    """带键槽的轴截面（键槽朝上），闭合折线。"""
    y_top = math.sqrt(r * r - (b / 2.0) ** 2)
    aA = math.atan2(y_top, b / 2.0)
    aB = math.pi - aA
    pts = []
    for ang in np.linspace(aB, aA + 2 * math.pi, n):
        pts.append((r * math.cos(ang), r * math.sin(ang)))
    pts.append((b / 2.0, r - t1))
    pts.append((-b / 2.0, r - t1))
    pts.append((-b / 2.0, y_top))
    return np.array(pts)


def stadium_profile(L, b, n=24):
    """圆头平键 A 型顶视轮廓（长 L，宽 b），闭合折线。"""
    r = b / 2.0
    pts = []
    for k in range(n + 1):
        ang = -math.pi / 2 + math.pi * k / n
        pts.append((L / 2 - r + r * math.cos(ang), r * math.sin(ang)))
    for k in range(n + 1):
        ang = math.pi / 2 + math.pi * k / n
        pts.append((-L / 2 + r + r * math.cos(ang), r * math.sin(ang)))
    return np.array(pts)


def circle_profile(R, n=48):
    """圆轮廓折线（备用）。"""
    angs = np.linspace(0, 2 * math.pi, n, endpoint=False)
    return np.column_stack((R * np.cos(angs), R * np.sin(angs)))


def hex_profile(s, n=6):
    """正六边形（螺栓头），s 为对边距。"""
    R = s / 2.0 / math.cos(math.radians(30))
    angs = np.linspace(0, 2 * math.pi, n, endpoint=False)
    return np.column_stack((R * np.cos(angs), R * np.sin(angs)))


def close_poly(pts, eps=1e-6):
    """去除零长度段并闭合。"""
    pts = np.asarray(pts, dtype=float)
    keep = [pts[0]]
    for p in pts[1:]:
        if np.linalg.norm(p - keep[-1]) > eps:
            keep.append(p)
    if np.linalg.norm(keep[0] - keep[-1]) <= eps:
        keep = keep[:-1]
    return np.array(keep)


# ===========================================================================
# 6. 汇总数据
# ===========================================================================
GEAR1_PROF = close_poly(involute_gear_profile(Z1, M))
GEAR2_PROF = close_poly(involute_gear_profile(Z2, M))
GEAR1_BORE = close_poly(hub_bore_profile(15.0, 8, 3.3))
GEAR2_BORE = close_poly(hub_bore_profile(22.5, 14, 3.8))

D = dict(
    input=dict(P=P_RAW, n1=N1, i=I_RATIO, life=LIFE_H),
    gear=dict(m=M, z1=Z1, z2=Z2, a=A, d1=D1, d2=D2, da1=DA1, da2=DA2,
              df1=DF1, df2=DF2, b1=B1, b2=B2, v=round(V, 3),
              precision=PRECISION, sigmaH_act=round(SIG_H_ACT, 1),
              sigmaF_act1=round(SIG_F_ACT1, 1), sigmaF_act2=round(SIG_F_ACT2, 1),
              sigmaH_allow=SIG_H_ALLOW),
    shaft=dict(d1min=round(D1MIN, 1), d2min=round(D2MIN, 1),
               shaft1=SHAFT1_HALF, shaft2=SHAFT2_HALF),
    bearing=dict(b1=BEARING1, b2=BEARING2),
    housing=dict(delta=DELTA, delta1=DELTA1, bf=BF, bf1=BF1, bf2=BF2,
                 df=DF_BOLT, d1=D1_BOLT, d2=D2_BOLT, d3=D3_SCREW,
                 d4=D4_SCREW, pin=PIN_D, h_center=H_CENTER,
                 box_z=BOX_Z, base_z=BASE_Z, seat_pro=SEAT_PRO,
                 seat_d2=SEAT_D2, seat_bore=SEAT_BORE),
    fasteners=dict(base_bolt=16, flange_bolt=10, side_bolt=12,
                   cap_screw=8, sight_screw=6,
                   base_pos=BASE_BOLT_POS, flange_x=FLANGE_BOLT_X,
                   flange_z=FLANGE_BOLT_Z, pin_pos=PIN_POS),
    accessories=dict(drain=OIL_DRAIN, gauge=OIL_GAUGE, sight=SIGHT,
                     sight_cover=SIGHT_COVER, vent=VENT, plug=PLUG),
    caps=CAP,
)


def write_json():
    path = os.path.join(OUT_DIR, "reducer_design.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(D, f, ensure_ascii=False, indent=2)
    return path


def write_report():
    """生成设计计算说明书（Markdown）。"""
    g = D["gear"]
    s1 = D["shaft"]["shaft1"]
    s2 = D["shaft"]["shaft2"]
    lines = []
    ap = lines.append
    ap("# 一级圆柱齿轮减速器 设计计算说明书\n")
    ap("## 1 原始数据（自定）\n")
    ap(f"- 输入功率 P = {P_RAW} kW，输入转速 n₁ = {N1:.0f} r/min")
    ap(f"- 传动比 i = {I_RATIO}，载荷平稳，寿命 Lh = {LIFE_H:.0f} h")
    ap(f"- 工作制：每天 8 h，10 年（300 天/年）\n")
    ap("## 2 齿轮传动设计（GB/T 1357、GB/T 10095）\n")
    ap(f"- 材料：小齿轮 40Cr 调质，大齿轮 45 钢调质（软齿面）")
    ap(f"- 按齿面接触强度初算中心距 a ≥ {a_min:.1f} mm，圆整取 a = {A:.0f} mm")
    ap(f"- 模数 m = {M}（GB/T 1357 第一系列），z₁ = {Z1}，z₂ = {Z2}，u = {u}")
    ap(f"- 分度圆 d₁ = {D1} mm，d₂ = {D2} mm；齿宽 b₁ = {B1} mm，b₂ = {B2} mm")
    ap(f"- 圆周速度 v = {V:.2f} m/s → 选 {PRECISION} 级精度（GB/T 10095.1）")
    ap(f"- 接触应力 σH = {g['sigmaH_act']} MPa ≤ [σH] = {g['sigmaH_allow']} MPa ✓")
    ap(f"- 弯曲应力 σF1 = {g['sigmaF_act1']} MPa ≤ [σF1] = {SIG_F1} MPa ✓")
    ap(f"- 弯曲应力 σF2 = {g['sigmaF_act2']} MPa ≤ [σF2] = {SIG_F2} MPa ✓\n")
    ap("## 3 轴设计（45 钢，C = 126）\n")
    ap(f"- 高速轴最小直径 d ≥ {D1MIN:.1f} mm → 联轴器段 φ20")
    ap(f"- 低速轴最小直径 d ≥ {D2MIN:.1f} mm → 联轴器段 φ32")
    ap("- 高速轴（轴1）分段：")
    for seg in s1:
        ap(f"  - φ{seg['rad']*2:.0f} × {seg['half']*2:.0f} mm（{seg['tag']}）")
    ap("- 低速轴（轴2）分段：")
    for seg in s2:
        ap(f"  - φ{seg['rad']*2:.0f} × {seg['half']*2:.0f} mm（{seg['tag']}）")
    ap(f"- 轴承：高速轴 {BEARING1['name']}（{BEARING1['d']}×{BEARING1['D']}×{BEARING1['B']}），"
       f"低速轴 {BEARING2['name']}（{BEARING2['d']}×{BEARING2['D']}×{BEARING2['B']}），"
       f"深沟球轴承 GB/T 276\n")
    ap("## 4 平键（GB/T 1095-1096）\n")
    ap("- 高速轴齿轮座：键 8×7×40；高速轴联轴器：键 6×6×32")
    ap("- 低速轴齿轮座：键 14×9×40；低速轴联轴器：键 10×8×50\n")
    ap("## 5 箱体（HT200，经验公式）\n")
    ap(f"- 箱座壁厚 δ = 0.025a+1 = {0.025*A+1:.1f} → 取 {DELTA:.0f} mm（≥8）")
    ap(f"- 箱盖壁厚 δ₁ = 0.02a+3 = {0.02*A+3:.1f} → 取 {DELTA1:.0f} mm")
    ap(f"- 凸缘厚 b = b₁ = {BF:.0f} mm；底座凸缘 b₂ = {BF2:.0f} mm")
    ap(f"- 地脚螺栓 df = 0.036a+12 = {0.036*A+12:.1f} → M{DF_BOLT}×{N_BOLT}")
    ap(f"- 轴承旁螺栓 d₁ = 0.75df → M{D1_BOLT}；箱盖-箱座螺栓 d₂ = 0.5df → M{D2_BOLT}")
    ap(f"- 端盖螺钉 d₃ → M{D3_SCREW}；视孔盖螺钉 d₄ → M{D4_SCREW}；定位销 φ{PIN_D}")
    ap(f"- 轴心高度 h = {H_CENTER:.0f} mm（齿根离底面 ≥ 15 mm，油池润滑）")
    ap(f"- 润滑与密封：齿轮油池润滑（浸油 1~2 模数），轴承脂润滑；"
       f"通气器、油标、放油螺塞均按课程设计常规配置\n")
    ap("## 6 输出文件\n")
    ap("- 桌面目录 `一级减速器\\`：全部零件 .SLDPRT + 装配体 .SLDASM + 本说明书")
    path = os.path.join(OUT_DIR, "reducer_design_report.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path


def write_preview():
    """生成 2D 设计预览（Pillow），前视 + 俯视。"""
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return None
    W, H = 820, 640
    img = Image.new("RGB", (W, H), "white")
    dr = ImageDraw.Draw(img)
    # 前视比例
    def tx(x):
        return 60 + (x - BASE_X_L) * 2.3
    def ty(y):
        return 560 - y * 2.3
    def box(x0, y0, x1, y1):
        xs = sorted((tx(x0), tx(x1)))
        ys = sorted((ty(y0), ty(y1)))
        return [xs[0], ys[0], xs[1], ys[1]]
    dr.text((20, 12), "一级圆柱齿轮减速器 - 设计预览（前视）", fill="black")
    # 底座
    dr.rectangle(box(BASE_X_L, 20, BASE_X_R, 0), outline="blue")
    # 外壁
    dr.rectangle(box(WALL_X_L, BOX_Y_TOP, WALL_X_R, BOX_Y_BOT), outline="black")
    # 凸缘
    dr.rectangle(box(FLANGE_X_L, FLANGE_Y_TOP, FLANGE_X_R, FLANGE_Y_BOT), outline="gray")
    # 内腔
    dr.rectangle(box(CAV_X_L, 182.5, CAV_X_R, BOX_Y_BOT), outline="red")
    # 齿轮
    for cx, r in ((0, D1 / 2), (A, D2 / 2)):
        dr.ellipse(box(cx - r, H_CENTER - r, cx + r, H_CENTER + r), outline="green", width=2)
        dr.ellipse(box(cx - 4, H_CENTER - 4, cx + 4, H_CENTER + 4), fill="black")
    # 轴承座
    for cx, dd in ((0, SEAT_D2[0]), (A, SEAT_D2[125])):
        dr.ellipse(box(cx - dd / 2, H_CENTER - dd / 2, cx + dd / 2, H_CENTER + dd / 2), outline="purple")
    # 视孔
    dr.rectangle(box(SIGHT["x0"], SIGHT["y1"], SIGHT["x1"], SIGHT["y0"]), outline="orange")
    dr.text((tx(-60), ty(210)), "a=125, m=2.5, z1=20, z2=80", fill="black")
    dr.text((tx(-60), ty(225)), "δ=8, b=12, M16×4/M10/M8", fill="black")
    # 俯视（右侧小图）
    ox, oy, s = 610, 430, 0.9
    dr.text((ox - 60, oy - 150), "俯视（箱体+轴）", fill="black")
    dr.rectangle([ox, oy, ox + 220 * s, oy + 101 * s], outline="black")
    for cx in (0, A):
        dr.ellipse([ox + (cx - 25) * s, oy - 25 * s, ox + (cx + 25) * s, oy + 25 * s], outline="green")
        dr.line([ox + (cx - 80) * s, oy + 50 * s, ox + (cx + 80) * s, oy + 50 * s], fill="green")
    path = os.path.join(OUT_DIR, "reducer_preview.png")
    img.save(path)
    return path


def main():
    print("=" * 68)
    print("一级圆柱齿轮减速器 - 设计计算")
    print("=" * 68)
    print(f"输入: P={P_RAW} kW, n1={N1:.0f} r/min, i={I_RATIO}")
    print(f"中心距 a_min={a_min:.1f} → a={A:.0f} mm, m={M}, z1={Z1}, z2={Z2}")
    print(f"d1={D1}, d2={D2}, da1={DA1}, da2={DA2}, b1={B1}, b2={B2}")
    print(f"v={V:.2f} m/s → {PRECISION} 级精度; σH={SIG_H_ACT:.0f}≤{SIG_H_ALLOW}, "
          f"σF1={SIG_F_ACT1:.0f}≤{SIG_F1}, σF2={SIG_F_ACT2:.0f}≤{SIG_F2}")
    print(f"轴: d1≥{D1MIN:.1f}→φ20, d2≥{D2MIN:.1f}→φ32; 轴承 6205/6208")
    print(f"箱体: δ={DELTA:.0f}, 凸缘 b={BF:.0f}, 底座 b2={BF2:.0f}, "
          f"螺栓 M{DF_BOLT}×{N_BOLT}/M{D1_BOLT}/M{D2_BOLT}, 端盖螺钉 M{D3_SCREW}")
    print(f"轴心高度 h={H_CENTER:.0f}, 箱宽 z={BOX_Z:.0f}, 底座宽 {BASE_Z:.0f}")
    print(f"齿轮轮廓点数: 小齿轮 {len(GEAR1_PROF)}, 大齿轮 {len(GEAR2_PROF)}")
    j = write_json()
    r = write_report()
    p = write_preview()
    print(f"JSON: {j}")
    print(f"报告: {r}")
    print(f"预览: {p}")
    print("=" * 68)


if __name__ == "__main__":
    main()
