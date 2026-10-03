# -*- coding: utf-8 -*-
"""开发用：渲染主截面预览图（不调用 SolidWorks）。"""
import os
import importlib.util

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    "gripper", os.path.join(_here, "build_umi_fin_gripper.py"))
g = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(g)

from PIL import Image, ImageDraw

dense, quads, holes, stats = g.build_geometry()
scale = 8  # px/mm
W, H = int(40 * scale), int(124 * scale)
img = Image.new("RGB", (W, H), "white")
d = ImageDraw.Draw(img)

def mm2px(x, y):
    return (x * scale, (124.0 - y) * scale)

d.polygon([mm2px(x, y) for x, y in dense], fill=(205, 225, 240), outline=(20, 20, 20))
for q in quads:
    d.polygon([mm2px(x, y) for x, y in q], fill=(120, 160, 205), outline=(20, 20, 20))
for hx, hy in holes:
    cx, cy = mm2px(hx, hy)
    r = 2.0 * scale
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill="white", outline=(20, 20, 20))

out = os.path.join(_here, "profile_preview.png")
os.makedirs(os.path.dirname(out), exist_ok=True)
img.save(out)
print("saved:", out)
