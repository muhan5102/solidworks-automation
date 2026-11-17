"""MCP创建圆角后的局部组边补充；固定本题几何位置，不用边序号。"""
import json
import math
from pathlib import Path
import sys
import pythoncom
import win32com.client

sys.path.insert(0,str(Path(__file__).resolve().parents[3]/"practice"))
from shared.cad import member as g,save_document
from sw_connect import create_empty_dispatch_variant
from sw_edge_select import iter_model_edges

out=Path(sys.argv[1]).resolve()
target=out/"踏架_LJT05.19.SLDPRT"
sw=win32com.client.Dispatch("SldWorks.Application.32");m=g(sw,"ActiveDoc")
assert g(sw,"GetProcessID")==49384 and Path(g(m,"GetPathName")).resolve()==target
f=m.FeatureByName("圆角3");d=g(f,"GetDefinition")
assert d.AccessSelections(m,create_empty_dispatch_variant())
selected=list(g(d,"Edges"));edges,_=iter_model_edges(m)
added=[]
for e in edges:
    p=e.get("mid_point_mm")
    if e.get("is_line") and p and abs(p[0]-42.265763)<.001 and abs(p[1]-32.541164)<.001 and abs(abs(p[2])-15.859360)<.001:
        selected.append(e["edge"]);added.append(p)
assert len(selected)==4 and len(added)==2
d.Edges=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,selected)
assert f.ModifyDefinition(d,m,create_empty_dispatch_variant())
assert m.ForceRebuild3(False)
f.Name="翼缘边缘R2"
m.FeatureByName("阵列(线性)1").Name="两沉头孔32间距_MCP"
m.FeatureByName("阵列(线性)2").Name="两M10孔25间距_MCP"
cm=g(m.Extension,"CustomPropertyManager","")
for name,value in [("图号","LJT05.19"),("名称","踏架"),("材料","HT150"),
    ("螺纹说明","2×M10-7H，GB孔向导及装饰螺纹，近壁通入主孔"),
    ("技术要求","铸件不得有气孔、裂纹及砂眼；时效、退火消除内应力；未注圆角R2～R4；GB/T 1804—2000-c；GB/T 1184—1996-L")]:
    cm.Add3(name,30,value,2)
assert save_document(m,str(target))
m.ClearSelection2(True);m.ShowNamedView2("*Isometric",7);m.ViewZoomtofit2();m.GraphicsRedraw2()
(out/"mcp-operations.json").write_text(json.dumps({"operations":["踏板四角R5","沉头孔32间距阵列","M10孔25间距阵列","根部R3初始两边","翼缘R2初始两边"],
    "COM补充":"主体草图、Rib、斜基准、GB孔向导、真实双向贯穿、孔终止纠错、同值圆角补充组边与属性",
    "note":"主体阶段在用户提醒前走COM；用户提醒后优先MCP。脚本未从空白整段重放。"},ensure_ascii=False,indent=2),encoding="utf-8")
print("翼缘上下四条边共用R2，已保存；待导出重开验收",flush=True)
