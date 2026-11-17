"""用户纠正：未注半径仅适用于图上已画出圆弧过渡的位置。"""
import json
from pathlib import Path
import sys
import pythoncom
import win32com.client

sys.path.insert(0,str(Path(__file__).resolve().parents[3]/"practice"))
from shared.cad import member as g,save_document
from sw_connect import create_empty_dispatch_variant
from sw_edge_select import iter_model_edges
from sw_review import save_preview,collect_model_summary

out=Path(sys.argv[1]).resolve();target=out/"踏架_LJT05.19_圆角核对.SLDPRT"
sw=win32com.client.Dispatch("SldWorks.Application.32");m=g(sw,"ActiveDoc")
assert g(sw,"GetProcessID")==49384 and Path(g(m,"GetPathName")).resolve()==target
f=m.FeatureByName("图示根部R3");d=g(f,"GetDefinition")
assert d.AccessSelections(m,create_empty_dispatch_variant())
seed=list(g(d,"Edges"));assert len(seed)==9
# 用几何签名识别九条母边，不按不稳定的边序号移除。
edges,_=iter_model_edges(m)
remove=[]
for row in edges:
    p=row.get("mid_point_mm")
    if not p:continue
    if ((abs(p[0]-59)<.001 and abs(p[1]-31.997361)<.001 and abs(p[2])<.001)
            or (abs(p[0]-76)<.001 and abs(p[1]-37)<.001 and abs(p[2])<.001 and abs(row["length_mm"]-10)<.001)):
        remove.append(row)
assert len(remove)==2
kept=[e for e in seed if not any(e._oleobj_==r["edge"]._oleobj_ for r in remove)]
assert len(kept)==7
d.Edges=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,kept)
assert f.ModifyDefinition(d,m,create_empty_dispatch_variant()) and m.ForceRebuild3(False)
health=collect_model_summary(m)
assert not health["faulty_features"] and not health["warning_features"]
assert save_document(m,str(target))
for view in ["front","isometric","bottom"]:save_preview(m,str(out/f"corrected_{view}.bmp"),view)
(out/"fillet-correction.json").write_text(json.dumps({
    "removed":[{"location":"翼缘接踏板下侧台阶","reason":"原主视内轮廓画成直线折角，无圆弧过渡"},
               {"location":"腹板斜边落到踏板处","reason":"原主视斜边与踏板相交，未画圆弧过渡"}],
    "retained":{"踏板R5":"俯视明确标R5","根部R3":"原主视与T形移出断面明确画出过渡，未注R2～R4内取R3；7条母边",
                "翼缘R2":"T形移出断面明确画出翼缘端部圆弧，未注R2～R4内取R2；4条母边"},
    "prior_result":"r01数值检查通过但多余圆角漏检，不记合格首轮"},ensure_ascii=False,indent=2),encoding="utf-8")
print("已撤除两处无图示依据的R3，7条有图示依据的母边保留；重建与特征健康通过",flush=True)
