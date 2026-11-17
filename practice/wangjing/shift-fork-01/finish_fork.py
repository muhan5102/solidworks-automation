"""接续用户手动现场；备份后替换未消费试验草图，完成本题筋。"""
import json
from pathlib import Path
import shutil
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
sys.path.insert(0, str(HERE.parents[1]))
from shared.single_part import g, save
from sw_review import save_preview
from rib_fork import create_rib
import win32com.client

out = Path(sys.argv[1]).resolve()
target = out/"拨叉_LJT05.01.SLDPRT"
s = win32com.client.Dispatch("SldWorks.Application.32")
m = g(s, "ActiveDoc")
assert Path(g(m, "GetPathName")).resolve() == target
assert g(s, "GetProcessID") == 12356
timer = time.perf_counter()
if g(m.SketchManager, "ActiveSketch") is not None:
    m.SketchManager.InsertSketch(True)
save(m, target)
backup = out/"拨叉_用户手动草图阶段.SLDPRT"
assert not backup.exists()
shutil.copy2(target, backup)
for name in ["草图4", "基准面1", "最外角筋轮廓"]:
    f = m.FeatureByName(name)
    if f is not None:
        m.ClearSelection2(True)
        assert f.Select2(False, 0)
        if not m.Extension.DeleteSelection2(0):
            # 已知本机残留树节点：只保存重开一次，不重复盲删。
            save(m, target)
            s.CloseDoc(g(m, "GetTitle"))
            errors = win32com.client.VARIANT(16387, 0)
            warnings = win32com.client.VARIANT(16387, 0)
            m = s.OpenDoc6(str(target), 1, 1, "", errors, warnings)
            assert m is not None and errors.value == 0
            f = m.FeatureByName(name)
            if f is not None:
                assert f.Select2(False, 0)
                assert m.Extension.DeleteSelection2(0), name
create_rib(s, m)
assert m.ForceRebuild3(False)
save(m, target)
save_preview(m, str(out/"放样曲面限定筋.bmp"), "isometric")
(out/"loft-rib-stage.json").write_text(json.dumps({
    "execution_seconds": time.perf_counter()-timer,
    "route": "surface loft then boss up to surface",
    "solid_bodies": len(g(m, "GetBodies2", 0, False)),
    "surface_bodies": len(g(m, "GetBodies2", 1, False) or []),
}, ensure_ascii=False, indent=2), encoding="utf-8")
print("放样限定筋已保存", flush=True)
