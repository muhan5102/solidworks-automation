"""王静图集LJT04.21拨盘；轴向Y，背面Y=0，偏置孔在世界+Z侧。"""
import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'year08/connector'))
from build_connector import (member, mm, sketch, sketch_line, new_document,
    save_document, revolve_boss, dimension_existing_relations, offset_plane, circle_cut)
from sw_connect import create_empty_dispatch_variant, open_document
from sw_sketch_quality import inspect_model_sketches
from sw_review import collect_model_summary, save_preview
import win32com.client

parser = argparse.ArgumentParser()
parser.add_argument('output', type=Path)
parser.add_argument('--process-id', type=int, required=True)
parser.add_argument('--name', default='拨盘_LJT04.21.SLDPRT')
args = parser.parse_args()
out = args.output.resolve()
assert Path(args.name).name == args.name and Path(args.name).suffix.lower()=='.sldprt'
dest = out / args.name
assert out.is_dir() and not dest.exists()
started = time.monotonic()
sw = win32com.client.Dispatch('SldWorks.Application.32')
assert int(member(sw, 'GetProcessID')) == args.process_id
sw.Visible = True
m = new_document(sw, 'part')

# Given: D49, D29.4, D16, D10, L19, flange5.5, shoulder7.5, relief1x0.2.
# The two sloping bore-mouth lines are C0.5; outer square edges stay square.
points = [(5.5,0),(5,.5),(5,18.5),(5.5,19),(8,19),(8,13),
          (14.7,13),(14.7,6.5),(14.5,6.5),(14.5,5.5),(24.5,5.5),(24.5,0)]
with sketch(m, 'Front Plane') as ref:
    sm = m.SketchManager
    old_add = sm.AddToDB
    sm.AddToDB = True
    try:
        axis = sm.CreateCenterLine(0,0,0,0,mm(19),0)
        lines = [sketch_line(m,mm(a[0]),mm(a[1]),mm(b[0]),mm(b[1]))
                 for a,b in zip(points,points[1:]+points[:1])]
    finally:
        sm.AddToDB = old_add
    assert axis is not None and all(e is not None for e in lines)
    sd = member(m.SelectionManager,'CreateSelectData')
    for a,b in zip(lines,lines[1:]+lines[:1]):
        m.ClearSelection2(True)
        assert member(a,'GetEndPoint2').Select4(False,sd)
        assert member(b,'GetStartPoint2').Select4(True,sd)
        m.SketchAddConstraints('sgCOINCIDENT')
    for line,a,b in zip(lines,points,points[1:]+points[:1]):
        relation = 'sgVERTICAL' if a[0]==b[0] else 'sgHORIZONTAL' if a[1]==b[1] else None
        if relation:
            m.ClearSelection2(True)
            assert line.Select4(False,sd)
            m.SketchAddConstraints(relation)
    m.ClearSelection2(True)
    assert axis.Select4(False,sd)
    m.SketchAddConstraints('sgVERTICAL')
    m.ClearSelection2(True)
    assert member(axis,'GetStartPoint2').Select4(False,sd)
    assert m.Extension.SelectByID2('Point1@原点','EXTSKETCHPOINT',0,0,0,True,0,create_empty_dispatch_variant(),0)
    m.SketchAddConstraints('sgCOINCIDENT')
    # The bottom profile point needs a Y=0 relation; zero-offset dimensions are omitted.
    m.ClearSelection2(True)
    assert member(lines[0],'GetStartPoint2').Select4(False,sd)
    assert member(axis,'GetStartPoint2').Select4(True,sd)
    m.SketchAddConstraints('sgHORIZONTALPOINTS2D')
    dimension_existing_relations(m)
    for line,p in zip(lines,points):
        actual=member(line,'GetStartPoint2')
        assert math.dist([member(actual,'X')*1000,member(actual,'Y')*1000],p)<1e-5
    contours=list(member(member(sm,'ActiveSketch'),'GetSketchContours') or [])
    assert contours and all(member(c,'IsClosed') for c in contours)
feature=revolve_boss(m,ref,2*math.pi)
assert feature is not None
feature.Name='主体回转'
m.FeatureByName(str(ref)).Name='轴向截面'
assert m.ForceRebuild3(False) and save_document(m,str(dest))
m.ShowNamedView2('*Isometric',7)
m.ViewZoomtofit2()
m.GraphicsRedraw2()
print('[主体] 台阶、中心孔、1×0.2退刀槽及图示孔口C0.5已完成',flush=True)

entry=offset_plane(m,'Top Plane',13,'沉台起始面')
circle_cut(m,sw,entry,[(0,13,17.8)],11,9,-1,'偏置沉台')
circle_cut(m,sw,'Top Plane',[(0,0,17.8)],3.25,0,1,'偏置通孔')
props=m.Extension.CustomPropertyManager('')
for k,v in [('名称','拨盘'),('图号','LJT04.21'),('材料','45'),
            ('配合','Φ29.4h7；Φ10J7；Φ6.5H7，按公称尺寸建模'),
            ('技术要求','未注倒角C0.5；去毛刺、飞边；锐边倒钝'),
            ('来源','王静《新标准机械图图集》PDF82/印刷76页')]:
    props.Add3(k,30,v,2)
assert m.ForceRebuild3(False) and save_document(m)
sw.CloseDoc(member(m,'GetTitle'))
m=open_document(sw,str(dest),silent=True)
assert m is not None and m.ForceRebuild3(False)
quality=inspect_model_sketches(m)
health=collect_model_summary(m)
faces=[]
for body in member(m,'GetBodies2',0,False):
    for face in member(body,'GetFaces'):
        s=member(face,'GetSurface')
        item={'box_m':list(member(face,'GetBox')),'area_m2':member(face,'GetArea')}
        for label,kind,params in [('cylinder','IsCylinder','CylinderParams'),('cone','IsCone','ConeParams'),('plane','IsPlane','PlaneParams')]:
            if member(s,kind):
                item[label]=list(member(s,params))
        faces.append(item)
def cylinder(radius,center_z,low,high):
    return any('cylinder' in f and abs(f['cylinder'][6]*1000-radius)<1e-5 and
        abs(f['cylinder'][0])<1e-9 and abs(abs(f['cylinder'][4])-1)<1e-9 and
        abs(f['cylinder'][2]*1000-center_z)<1e-5 and
        abs(f['box_m'][1]*1000-low)<1e-5 and abs(f['box_m'][4]*1000-high)<1e-5 for f in faces)
cones=[f for f in faces if 'cone' in f]
checks={'single_solid':len(member(m,'GetBodies2',0,False))==1,
        'fully_defined_no_fix':quality['status']=='pass',
        'no_error_or_warning':not health['faulty_features'] and not health['warning_features'],
        'flange_D49_thickness5_5':cylinder(24.5,0,0,5.5),
        'hub_D16_length6':cylinder(8,0,13,19),
        'shoulder_D29_4':cylinder(14.7,0,6.5,13),
        'relief_width1_depth0_2':cylinder(14.5,0,5.5,6.5),
        'center_bore_D10':cylinder(5,0,.5,18.5),
        'offset_pocket_D22_Y4_to13':cylinder(11,17.8,4,13),
        'offset_hole_D6_5_remaining4':cylinder(3.25,17.8,0,4),
        'two_bore_chamfers_only':len(cones)==2,
        'bore_chamfers_C0_5_at_both_ends':len(cones)==2 and
            all(abs((f['box_m'][4]-f['box_m'][1])*1000-.5)<1e-5 for f in cones) and
            sorted(round(f['box_m'][1]*1000,5) for f in cones)==[0,18.5]}
report={'checks':checks,'faces':faces,'sketches':quality,'health':health,
        'execution_seconds':time.monotonic()-started,'finished':datetime.now().astimezone().isoformat(),
        'limits':'材料与公差仅属性；未生成工程图和装配；未查看参考立体答案'}
(out/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
assert all(checks.values()),checks
if member(m,'GetVisibilityOfConstructPlanes'):
    member(m,'ViewDispRefplanes')
save_preview(m,str(out/'拨盘_等轴测.bmp'),'isometric')
save_preview(m,str(out/'拨盘_正视.bmp'),'top')
save_preview(m,str(out/'拨盘_侧视.bmp'),'front')
assert save_document(m)
print(json.dumps({'checks':checks,'execution_seconds':report['execution_seconds']},ensure_ascii=False),flush=True)
