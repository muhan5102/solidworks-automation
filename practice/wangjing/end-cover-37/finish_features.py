"""端盖LJT04.37分阶段特征，需先通过MCP回转并保存主体。"""
import json
import math
from datetime import datetime
from pathlib import Path
import sys
import win32com.client

out = Path(sys.argv[1]).resolve()
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import member as g, dimension_existing_relations, round_edges
from shared.single_part import anchor, quiet_dimensions, save
from sw_connect import create_empty_dispatch_variant
from sw_part import fillet, _get_plane_name_candidates, _select_first_candidate
from sw_pattern import circular_pattern
from shared.single_part import surfaces

pid = int(sys.argv[2])
stage = sys.argv[3]
target = out / '端盖_LJT04.37.SLDPRT'
sw = win32com.client.Dispatch('SldWorks.Application.32')
m = g(sw,'ActiveDoc')
assert g(sw,'GetProcessID') == pid and Path(g(m,'GetPathName')).resolve() == target


def checkpoint(label):
    assert m.ForceRebuild3(False)
    assert len(g(m,'GetBodies2',0,False)) == 1
    save(m,target)
    m.ClearSelection2(True)
    m.ShowNamedView2('*Isometric',7)
    m.ViewZoomtofit2()
    m.GraphicsRedraw2()
    info = {'time':datetime.now().astimezone().isoformat(),'stage':label,'file':target.name}
    (out / f'checkpoint-{stage}.json').write_text(json.dumps(info,ensure_ascii=False),encoding='utf-8')
    print(label, flush=True)


def at_face(entry, direction):
    m.ClearSelection2(True)
    o = [p-.004*n for p,n in zip(entry,direction)]
    assert m.Extension.SelectByRay(*o,*direction,1e-6,2,False,0,0)
    sm = g(m,'SelectionManager')
    hit = g(sm,'GetSelectionPoint2',1,-1)
    assert math.dist(hit,entry)<1e-6,(hit,entry)


def locate(feature, name, expected):
    assert feature is not None, name
    feature.Name = name
    sub = g(feature,'GetFirstSubFeature')
    sub.Name = name + '孔位'
    m.ClearSelection2(True)
    assert sub.Select2(False,0)
    m.EditSketch()
    try:
        sk = g(m.SketchManager,'ActiveSketch')
        p, = g(sk,'GetSketchPoints2')
        p0 = [g(p,k) for k in ('X','Y','Z')]
        if abs(p0[0])<1e-9 or abs(p0[1])<1e-9:
            anchor(m,p,x=abs(p0[0])<1e-9,y=abs(p0[1])<1e-9)
        dimension_existing_relations(m)
        a = g(g(g(sk,'ModelToSketchTransform'),'Inverse'),'ArrayData')
        q = [g(p,k) for k in ('X','Y','Z')]
        actual = [sum(q[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
        assert math.dist(actual,expected)<1e-7,(actual,expected)
    finally:
        m.SketchManager.InsertSketch(True)
    return g(feature,'GetDefinition')


with quiet_dimensions(sw):
    if stage=='base':
        assert m.FeatureByName('台阶底部R3') is None
        m.FeatureByName('旋转1').Name = '主体回转'
        m.FeatureByName('草图1').Name = '轴向截面'
        matches = [e for e,p,r in round_edges(m) if math.dist(p,[0,50,0])<.001 and abs(r-28)<.001]
        edge, = matches
        m.ClearSelection2(True)
        assert edge.Select4(False,g(m.SelectionManager,'CreateSelectData'))
        f = fillet(m,.003)
        assert f is not None
        f.Name = '台阶底部R3'
        m.ClearSelection2(True)
        assert _select_first_candidate(m.Extension,_get_plane_name_candidates('Front Plane'),'PLANE')
        assert _select_first_candidate(m.Extension,_get_plane_name_candidates('Right Plane'),'PLANE',append=True)
        assert m.InsertAxis2(True)
        feat = g(m,'FirstFeature')
        while feat is not None:
            if g(feat,'GetTypeName2')=='RefAxis':
                feat.Name='回转轴'
                break
            feat = g(feat,'GetNextFeature')
        assert m.FeatureByName('回转轴') is not None
        checkpoint('主体回转、图示R3和阵列基准轴')
    elif stage=='mount':
        assert m.FeatureByName('安装沉孔母孔') is None
        entry=(.057,.016,0)
        at_face(entry,(0,1,0))
        # GB 内六角圆柱头螺钉。使用自定义Φ9／Φ15深9；回读验证实际值。
        f=m.FeatureManager.HoleWizard5(0,13,361,'M8',1,.009,.032,-1,
            .015,.009,0,-1,-1,0,0,0,0,0,0,-1,'',False,False,True,False,False,False)
        data=locate(f,'安装沉孔母孔',entry)
        assert g(data,'Standard2')==13
        assert abs(g(data,'ThruHoleDiameter')-.009)<1e-8
        assert abs(g(data,'CounterBoreDiameter')-.015)<1e-8
        assert abs(g(data,'CounterBoreDepth')-.009)<1e-8
        checkpoint('GB安装沉孔母孔待阵列')
    elif stage=='pattern6':
        label='六安装沉孔圆周阵列'
        assert m.FeatureByName(label) is None
        f=circular_pattern(m,'安装沉孔母孔','回转轴',math.tau,6)
        f.Name=label
        checkpoint(label)
    elif stage=='axial':
        assert m.FeatureByName('轴向交叉Φ16') is None
        entry=(0,.056,.030)
        at_face(entry,(0,-1,0))
        f=m.FeatureManager.HoleWizard5(2,13,355,'Ø16',0,.016,.016,-1,
            math.pi/2,0,0,0,0,0,-1,-1,-1,-1,-1,-1,'',False,False,True,False,False,False)
        data=locate(f,'轴向交叉Φ16',entry)
        assert g(data,'Standard2')==13 and g(data,'FastenerSize')=='Ø16'
        assert data.AccessSelections(m,create_empty_dispatch_variant())
        data.DrillAngle=math.pi/2
        assert f.ModifyDefinition(data,m,create_empty_dispatch_variant())
        assert abs(g(g(f,'GetDefinition'),'DrillAngle')-math.pi/2)<1e-8
        checkpoint('GB轴向Φ16全径16和90度锥尖')
    elif stage=='radial':
        from shared.cad import sketch, sketch_line
        from shared.single_part import relation
        from sw_part import _ensure_sketch_selected
        assert m.FeatureByName('径向交叉孔') is None
        # Rc1/4 底孔采用本机GB库的Φ10.7。18为入口轴向段，52为径向Ø16全径钻深。
        # 原图剖面进口画大于后段；实作保留两处明确规格，记录为几何学习假设。
        tip = 20 - 8/math.tan(math.radians(59))
        coords = [(32,74),(37.35,74),(37.35,54),(40,54),(40,20),(32,tip)]
        with sketch(m,'Right Plane') as ref:
            sm=m.SketchManager
            sk=g(sm,'ActiveSketch')
            a=g(g(sk,'ModelToSketchTransform'),'ArrayData')
            def local(y,z):
                xyz=(0,y/1000,z/1000)
                v=[sum(xyz[j]*a[j*3+i] for j in range(3))*a[12]+a[9+i] for i in range(3)]
                assert abs(v[2])<1e-8
                return v[:2]
            pts=[local(*q) for q in coords]
            old=sm.AddToDB; sm.AddToDB=True
            try:
                edges=[sketch_line(m,*a0,*a1) for a0,a1 in zip(pts,pts[1:]+pts[:1])]
                axis=sm.CreateCenterLine(*pts[0],0,*pts[-1],0)
            finally:
                sm.AddToDB=old
            for cur,nxt in zip(edges,edges[1:]+edges[:1]):
                relation(m,[g(cur,'GetEndPoint2'),g(nxt,'GetStartPoint2')],'sgCOINCIDENT')
            for i in (0,1,2,3,5):
                a0,a1=pts[i],pts[(i+1)%len(pts)]
                relation(m,[edges[i]],'sgVERTICAL' if abs(a0[0]-a1[0])<1e-9 else 'sgHORIZONTAL')
            relation(m,[g(axis,'GetStartPoint2'),g(edges[0],'GetStartPoint2')],'sgCOINCIDENT')
            relation(m,[g(axis,'GetEndPoint2'),g(edges[4],'GetEndPoint2')],'sgCOINCIDENT')
            dimension_existing_relations(m)
            for edge,xy in zip(edges,pts):
                p=g(edge,'GetStartPoint2')
                assert math.dist([g(p,'X'),g(p,'Y')],xy)<1e-8
        m.FeatureByName(str(ref)).Name='径向交叉孔截面'
        _ensure_sketch_selected(m,'径向交叉孔截面')
        sd=g(m.SelectionManager,'CreateSelectData'); sd.Mark=16
        assert axis.Select4(True,sd)
        f=m.FeatureManager.FeatureRevolve2(True,True,False,True,False,False,0,0,math.tau,0,False,False,0,0,0,0,0,True,False,True)
        assert f is not None
        f.Name='径向交叉孔'
        props=g(m.Extension,'CustomPropertyManager','')
        for name, value in [('名称','端盖'),('图号','LJT04.37'),('材料','HT200'),('径向接口','Rc1/4，入口段18，按GB库Φ10.7表达底孔，未建牙型'),('M8深度修正','底孔22，螺纹20'),('未注公差','原图 GB/T 1804—2000-m、GB/T 1184—1996-K')]:
            props.Add3(name,30,value,2)
        checkpoint('径向孔及Rc1/4接口属性')
    elif stage=='m8':
        assert m.FeatureByName('M8母孔') is None
        entry=(0,0,-.033)
        at_face(entry,(0,1,0))
        f=m.FeatureManager.HoleWizard5(4,13,359,'M8',0,.0068,.022,-1,
            .020,.008,math.radians(118),0,0,-1,1,0,-1,-1,-1,-1,'7H',False,False,True,False,False,False)
        data=locate(f,'M8母孔',entry)
        assert g(data,'Standard2')==13 and g(data,'ThreadClass')=='7H'
        assert abs(g(data,'ThreadDepth')-.020)<1e-8 and abs(g(data,'TapDrillDepth')-.022)<1e-8
        f=circular_pattern(m,'M8母孔','回转轴',math.tau,3)
        f.Name='三M8圆周阵列'
        assert m.ForceRebuild3(False)
        con=[s for s in surfaces(m) if 'Cone' in s and abs(s['Cone'][6]-.0034)<1e-8]
        print('M8钻尖',[(s['Cone'][:3],s['box'][1:5:3]) for s in con],flush=True)
        assert len(con)==3 and all(22-1e-5<=s['box'][1] and s['box'][4]<25 for s in con)
        assert m.FeatureByName('轴向交叉Φ16') is not None and m.FeatureByName('径向交叉孔') is not None
        checkpoint('GB M8三处实际盲孔、底孔22、螺纹20')
    else:
        raise ValueError(stage)
