"""定位套：准备含退刀槽的轴向截面，回转由正式MCP执行。"""
from datetime import datetime
import ctypes
import json
import math
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[2]/"year08/connector"))
from build_connector import (member, mm, sketch, sketch_line, new_document,
                             dimension_existing_relations, save_document)
from sw_connect import create_empty_dispatch_variant
from sw_sketch_quality import inspect_model_sketches
import win32com.client

out = Path(sys.argv[1]).resolve()
dest = out/"定位套_LJT03.28.SLDPRT"
assert out.is_dir() and not dest.exists()
sw = win32com.client.Dispatch("SldWorks.Application.32")
assert int(member(sw, "GetProcessID")) == 34092
sw.Visible = True
assert ctypes.windll.user32.IsWindowVisible(69222)
if ctypes.windll.user32.IsIconic(69222):
    ctypes.windll.user32.ShowWindow(69222, 9)
m = new_document(sw, "part")
# 坐标为(半径,轴向)。60°相对径向；退刀槽由双向0.6、4及R1构成。
cone_start = 150-5/math.tan(math.radians(30))
thread_r = (80-1.0825317547305484*2)/2
points = [(37,0),(48,0),(50,2),(50,15),(95,15),(95,48),
          (59,48),(58.4,47.4),(55.4,47.4),(54.4,48.4),
          (54.4,51.4),(55,52),(55,93),(50,93),(50,115),
          (55,115),(55,cone_start),(50,150),(50,188),(48,190),
          (thread_r+2,190),(thread_r,188),(thread_r,168),(41,168),(41,160),(37,160)]
with sketch(m, "Front Plane") as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0,0,0,0,mm(190),0)
        edges = []
        for i,(a,b) in enumerate(zip(points,points[1:]+points[:1])):
            if i == 8:
                edges.append(sm.CreateArc(mm(55.4),mm(48.4),0,
                    mm(a[0]),mm(a[1]),0,mm(b[0]),mm(b[1]),0,-1))
            else:
                edges.append(sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1])))
    finally:
        sm.AddToDB = old
    assert axis is not None and all(e is not None for e in edges)
    sd = member(m.SelectionManager,"CreateSelectData")
    def relation(items, kind):
        m.ClearSelection2(True)
        for i,item in enumerate(items):
            assert item.Select4(i>0,sd)
        m.SketchAddConstraints(kind)
    for a,b in zip(edges,edges[1:]+edges[:1]):
        relation([member(a,"GetEndPoint2"),member(b,"GetStartPoint2")],"sgCOINCIDENT")
    for i,(e,a,b) in enumerate(zip(edges,points,points[1:]+points[:1])):
        if i != 8 and (a[0]==b[0] or a[1]==b[1]):
            relation([e],"sgVERTICAL" if a[0]==b[0] else "sgHORIZONTAL")
    relation([edges[7],edges[8]],"sgTANGENT")
    relation([edges[8],edges[9]],"sgTANGENT")
    relation([axis],"sgVERTICAL")
    m.ClearSelection2(True)
    assert member(axis,"GetStartPoint2").Select4(False,sd)
    assert m.Extension.SelectByID2("Point1@原点","EXTSKETCHPOINT",0,0,0,True,0,
                                  create_empty_dispatch_variant(),0)
    m.SketchAddConstraints("sgCOINCIDENT")
    relation([member(edges[0],"GetStartPoint2"),member(axis,"GetStartPoint2")],
             "sgHORIZONTALPOINTS2D")
    # 回转轴终点和右端面共享轴向位置，避免零差尺寸被自动补尺寸略过。
    relation([member(axis,"GetEndPoint2"),member(edges[19],"GetStartPoint2")],
             "sgHORIZONTALPOINTS2D")
    relation([edges[2],edges[13]],"sgCOLINEAR")
    dimension_existing_relations(m)
    for e,p in zip(edges,points):
        actual = member(e,"GetStartPoint2")
        assert math.dist([member(actual,"X")*1000,member(actual,"Y")*1000],p)<1e-5
    contours = list(member(member(sm,"ActiveSketch"),"GetSketchContours") or [])
    assert len(contours)==1 and member(contours[0],"IsClosed")
m.FeatureByName(str(ref)).Name="轴向截面"
assert save_document(m,str(dest))
quality=inspect_model_sketches(m)
assert len(quality["sketches"])==1
assert all(s["constrained_status"]==3 and s["fix_count"]==0 for s in quality["sketches"])
m.ShowNamedView2("*Front",1)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
result={"path":str(dest),"points_mm":points,"quality":quality,
        "thread_limit":"用户指出右端螺纹段由30-8=22得出；小径采用公制基本牙型D1=80-1.0825317547*2，后续添加原生装饰螺纹，不建立实体牙型。",
        "finished":datetime.now().astimezone().isoformat()}
(out/"profile.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps({"path":str(dest),"constrained":quality["sketches"][0]["constrained_status"]},ensure_ascii=False))
