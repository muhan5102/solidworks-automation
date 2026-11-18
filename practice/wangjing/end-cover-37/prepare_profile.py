"""在本轮空白零件建立完全定义的轴向截面。主体回转和首次保存由MCP执行。"""
import sys
from pathlib import Path
import win32com.client
import win32gui
import win32process

out = Path(sys.argv[1]).resolve()
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, sketch, sketch_line, dimension_existing_relations
from shared.single_part import relation, anchor, quiet_dimensions

pid, hwnd = map(int, sys.argv[2:])
assert out.is_dir() and not (out / "端盖_LJT04.37.SLDPRT").exists()
assert win32gui.IsWindow(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1] == pid
win32gui.ShowWindow(hwnd, 9)
assert win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd)
sw = win32com.client.Dispatch('SldWorks.Application.32')
assert g(sw, 'GetProcessID') == pid
m = g(sw, 'ActiveDoc')
assert m is not None and g(m, 'GetPathName') == ''
assert g(m, 'GetType') == 1 and not g(m, 'GetBodies2', 0, False)
# Y begins on the Φ84 side. X is section radius. C1.6 belongs to Φ48 entry.
coords = [(0,42), (0,25.6), (1.6,24), (15,24), (15,13), (50,13), (50,28), (56,28), (56,43), (48,43), (48,72), (16,72), (16,42)]
with sketch(m, 'Front Plane') as ref:
    sm = m.SketchManager
    old = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0, 0, 0, 0, .056, 0)
        pairs = list(zip(coords, coords[1:] + coords[:1]))
        edges = [sketch_line(m, r/1000, y/1000, s/1000, z/1000) for (y,r),(z,s) in pairs]
    finally:
        sm.AddToDB = old
    for a,b in zip(edges,edges[1:]+edges[:1]):
        relation(m, [g(a,'GetEndPoint2'),g(b,'GetStartPoint2')], 'sgCOINCIDENT')
    for e,((y,r),(z,s)) in zip(edges,pairs):
        if y==z or r==s:
            relation(m,[e], 'sgHORIZONTAL' if y==z else 'sgVERTICAL')
    anchor(m,g(axis,'GetStartPoint2'),x=True,y=True)
    relation(m,[axis],'sgVERTICAL')
    relation(m,[g(axis,'GetStartPoint2'),g(edges[0],'GetStartPoint2')],'sgHORIZONTALPOINTS2D')
    relation(m,[g(axis,'GetEndPoint2'),g(edges[7],'GetEndPoint2')],'sgHORIZONTALPOINTS2D')
    with quiet_dimensions(sw):
        quality = dimension_existing_relations(m)
    for edge,(y,r) in zip(edges,coords):
        pt=g(edge,'GetStartPoint2')
        assert abs(g(pt,'X')-r/1000)<1e-8 and abs(g(pt,'Y')-y/1000)<1e-8
print(str(ref), quality['status'])
