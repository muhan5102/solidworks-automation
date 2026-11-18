"""对已由MCP导出STEP的端盖保存重开，检查实体与题面，不替代目视。"""
import json
import math
from pathlib import Path
import sys
from datetime import datetime
import win32com.client
from PIL import Image

out=Path(sys.argv[1]).resolve()
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from shared.single_part import g, surfaces, save, finish_review
from sw_review import run_review, save_preview
from sw_sketch_quality import inspect_model_sketches

pid=int(sys.argv[2])
sw=win32com.client.Dispatch('SldWorks.Application.32')
m=g(sw,'ActiveDoc')
target=out/'端盖_LJT04.37.SLDPRT'
assert g(sw,'GetProcessID')==pid and Path(g(m,'GetPathName')).resolve()==target
if len(sys.argv)>3:
    finish_review(sw,target,visual_notes=sys.argv[3])
    raise SystemExit

assert not (out/'verification.json').exists()
assert target.with_suffix('.step').stat().st_size>0
save(m,target)
sw.CloseDoc(g(m,'GetTitle'))
err,warn=(win32com.client.VARIANT(16387,0) for _ in range(2))
m=sw.OpenDoc6(str(target),1,1,'',err,warn)
assert m is not None and err.value==0 and Path(g(m,'GetPathName')).resolve()==target
assert m.ForceRebuild3(False)
m.ClearSelection2(True)
m.ViewZoomtofit2()
report,_=run_review(m,str(out),basename='end_cover',views=['isometric','front','top','right','bottom'],expected_outputs=[str(target),str(target.with_suffix('.step'))])
sk=inspect_model_sketches(m)
rows=surfaces(m)
checks={}


def near(a,b,tol=.005): return abs(a-b)<tol


def by_radius(kind,r):
    idx={'Cylinder':6,'Cone':6,'Torus':7}[kind]
    return [s for s in rows if kind in s and near(s[kind][idx]*1000,r)]


def span(s,ax,l,h): return near(s['box'][ax],l) and near(s['box'][ax+3],h)


def line_through(s,kind,axis,point):
    v=s[kind]
    return near(abs(v[axis+3]),1,1e-6) and all(near(v[k]*1000,point[k]) for k in range(3) if k!=axis)


def check_axial_cylinder(r,y0,y1,x=0,z=0):
    return any(span(s,1,y0,y1) and line_through(s,'Cylinder',1,(x,None,z)) for s in by_radius('Cylinder',r))


health=report['model'] if 'model' in report else {}
if 'faulty_features' not in health:
    from sw_review import collect_model_summary
    health=collect_model_summary(m)
checks['单实体']=len(g(m,'GetBodies2',0,False))==1
checks['特征无错无警告']=not any(health.get(k) for k in ('feature_error','faulty_features','warning_features','unknown_feature_health'))
manual={'轴向截面','安装沉孔母孔孔位','M8母孔孔位','轴向交叉Φ16孔位','径向交叉孔截面'}
checks['五张设计草图完全定义无Fix并被使用']={s['name'] for s in sk['sketches'] if s['name'] in manual and s['status']=='pass' and s['consumed_by'] and not s['fix_count']}==manual
internal=[s for s in sk['sketches'] if s['name'] not in manual]
checks['三张孔向导内部轮廓已完全约束']=len(internal)==3 and all(s['constrained_status']==3 and len(s['consumed_by'])==1 for s in internal)
checks['Φ84近端凸台16']=check_axial_cylinder(42,0,16)
checks['Φ144法兰厚32及端面位置']=check_axial_cylinder(72,16,48)
checks['Φ86远端凸台8及总厚56']=check_axial_cylinder(43,48,56)
checks['Φ48入口深15及C1.6']=check_axial_cylinder(24,1.6,15) and any(span(s,1,0,1.6) and near(s['Cone'][7],math.pi/4,.001) for s in by_radius('Cone',24))
checks['Φ26中央通道']=check_axial_cylinder(13,15,50)
checks['Φ56避空深6及图示R3']=check_axial_cylinder(28,53,56) and bool(by_radius('Torus',3)) and all(near(s['Torus'][6]*1000,25) and span(s,1,50,53) for s in by_radius('Torus',3))
checks['背面入口边保留折角']=all(near(s['Torus'][7]*1000,3) and near(s['Torus'][1]*1000,53) for s in rows if 'Torus' in s)
mount=[(57*math.cos(i*math.tau/6),57*math.sin(i*math.tau/6)) for i in range(6)]
checks['六孔Φ114等分、Φ9贯穿法兰']=all(check_axial_cylinder(4.5,25,48,x,z) for x,z in mount) and len(by_radius('Cylinder',4.5))==6
checks['六孔均从Φ84侧沉Φ15深9']=all(check_axial_cylinder(7.5,16,25,x,z) for x,z in mount) and len(by_radius('Cylinder',7.5))==6
m8=[(0,-33),(33*math.sqrt(3)/2,16.5),(-33*math.sqrt(3)/2,16.5)]
def tap_bore_depth(x,z):
    bore=[s for s in by_radius('Cylinder',3.4) if line_through(s,'Cylinder',1,(x,None,z))]
    lead=[s for s in by_radius('Cone',4) if line_through(s,'Cone',1,(x,None,z)) and near(s['box'][1],0) and 0<s['box'][4]<.5]
    return len(bore)==len(lead)==1 and span(bore[0],1,lead[0]['box'][4],22)

checks['三M8位于Φ66；导入锥与Φ6.8圆柱连续至深22']=all(tap_bore_depth(x,z) for x,z in m8) and len(by_radius('Cylinder',3.4))==3
checks['三个M8钻尖在实体内终止']=all(any(line_through(s,'Cone',1,(x,None,z)) and near(s['box'][1],22) and 23<s['box'][4]<25 for s in by_radius('Cone',3.4)) for x,z in m8)
d=g(m.FeatureByName('M8母孔'),'GetDefinition')
checks['M8母孔GB、7H、装饰螺纹深20']=g(d,'Standard2')==13 and g(d,'FastenerSize')=='M8' and g(d,'ThreadClass')=='7H' and near(g(d,'ThreadDepth')*1000,20)
checks['轴向Φ16中心偏30、全径到Y40']=check_axial_cylinder(8,40,56,0,30)
checks['轴向孔锥尖与径向孔相通']=any(line_through(s,'Cone',1,(0,None,30)) and near(s['Cone'][7],math.pi/4,.001) and near(s['box'][4],40) for s in by_radius('Cone',8))
checks['径向孔在法兰中厚、Φ16止于Z20']=any(line_through(s,'Cylinder',2,(0,32,None)) and span(s,2,20,54) for s in by_radius('Cylinder',8))
checks['径向入口Φ10.7、18长及内端保留材料']=any(line_through(s,'Cylinder',2,(0,32,None)) and span(s,2,54,72) for s in by_radius('Cylinder',5.35)) and any(line_through(s,'Cone',2,(0,32,None)) and s['box'][2]>13 and near(s['box'][5],20) for s in rows if 'Cone' in s)

m.ViewDisplayWireframe()
try:
    wire=save_preview(m,str(out/'right_wire.bmp'),'right')
    with Image.open(wire) as im: im.save(out/'right_wire.png')
finally:
    m.ViewDisplayShaded()
    m.ShowNamedView2('*Isometric',7)
    m.GraphicsRedraw2()
with Image.open(out/'end_cover_isometric.bmp') as im: im.save(out/'isometric.png')
report_data={'finished':datetime.now().astimezone().isoformat(),'process_id':pid,'target':target.name,'passed':all(checks.values()),'status':'awaiting_visual_review' if all(checks.values()) else 'failed','checks':checks,'sketches':sk,'health':health,'surfaces':rows,'limits':'Rc1/4按GB底孔几何和属性表达；原图入口台阶画法矛盾。原生模型终检，STEP未重导入。'}
(out/'verification.json').write_text(json.dumps(report_data,ensure_ascii=False,indent=2),encoding='utf8')
print('checks',len(checks),'failed',[k for k,v in checks.items() if not v], 'generic',report['evaluation']['status'])
