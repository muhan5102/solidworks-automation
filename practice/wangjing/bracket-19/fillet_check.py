"""原图两处直线折角检查；用实际边线识别，不以圆角特征数代替结构正确。"""
from sw_edge_select import iter_model_edges


def check_fillet_outline(model):
    edges,_=iter_model_edges(model)
    def line_at(x,y):
        return any(e.get("is_line") and e.get("mid_point_mm")
            and abs(e["mid_point_mm"][0]-x)<.001
            and abs(e["mid_point_mm"][1]-y)<.001
            and abs(e["mid_point_mm"][2])<.001
            and abs(abs(e["direction"][2])-1)<1e-7 and e["length_mm"]>10 for e in edges)
    # 前者Y由R9、5厚及踏板37高的切线关系得出；后者是原图76定位线与37顶面的交点。
    return {"翼缘与踏板下侧保留图示折角":line_at(59,31.997361),
            "腹板斜边与踏板保留图示交线":line_at(76,37)}
