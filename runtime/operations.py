"""Individually callable daily-role operations; executed only by audited runtime."""
from pathlib import Path
from datetime import date,timedelta,datetime
import hashlib,json
from . import reporting as base

FILES={
'sync':{'接入清单.xlsx','字段完整率.png','每日到数.png','sync.json'},
'detect':{'异常预警清单.xlsx','双轴趋势.png','detect.json'},
'copy':{'经营日报.docx','日报数据摘要.xlsx','report_copy.json'},
'daily':{'经营日报.docx','经营日报.xlsx','双轴趋势.png','report_flow.json','detect.json'},
'collect':{'采集归集单.xlsx','collect.json','sync.json'},
'alert':{'预警推送单.xlsx','推送消息预览.md','alert_flow.json','detect.json'},
'weekly':{'周度经营建议.docx','周度汇总.xlsx','周对比图.png','weekly_flow.json','detect.json'}
}

def source_check(payload):
    source=payload.get('source')
    if not isinstance(source,dict):return {'status':'needs_input','missing_fields':['source'],'downstream_allowed':False}
    if source.get('authorized') is not True:return {'status':'blocked','reason':'未授权的源不得采集','downstream_allowed':False}
    if source.get('access') not in {'owned_synthetic','user_export','manual_entry'}:return {'status':'blocked','reason':'本技能不连接外部系统；未启用自动连接器','downstream_allowed':False}
    missing=[k for k in ['id','owner','acquired_at','scope','expected_fields'] if not source.get(k)]
    if missing:return {'status':'needs_input','missing_fields':['source.'+k for k in missing],'downstream_allowed':False}
    if source['access']=='owned_synthetic' and payload.get('data_origin')!='SYNTHETIC':return {'status':'blocked','reason':'自有仿真不得冒充客户实绩','downstream_allowed':False}
    return None

def clean(payload):
    missing=[k for k in ['csv_text','store_id','start','end'] if not payload.get(k)]
    if missing:return {'status':'needs_input','missing_fields':missing,'output_files':[]}
    if len(payload['csv_text'].encode())>1000000:raise ValueError('当前输入超过1MB')
    if date.fromisoformat(payload['start'])>date.fromisoformat(payload['end']):raise ValueError('日期区间反转')
    checked=source_check(payload)
    if checked:return {**checked,'output_files':[]}
    sync=base.clean_csv(payload['csv_text'])
    import csv,io
    expected=set(payload['source']['expected_fields']);actual=set(csv.DictReader(io.StringIO(payload['csv_text'].lstrip('\ufeff'))).fieldnames or [])
    coverage=len(expected&actual)/len(expected)
    sync['coverage']=coverage
    if coverage<.5:return {'status':'blocked','reason':'应到字段覆盖不足50%','coverage':coverage,'output_files':[]}
    if coverage<.8:sync['quality'].append({'reason':'应到字段覆盖不足80%','missing':sorted(expected-actual)})
    invalid=[v for v in sync['quality'] if 'line' in v]
    if sync['status'] in {'blocked','needs_input'} or invalid:
        return {**sync,'status':'blocked' if sync['status']=='blocked' else 'needs_input','missing_fields':['review_invalid_source_rows'] if invalid else ['core_fields'],'output_files':[]}
    rows=[v for v in sync['rows'] if v['store_id']==payload['store_id']]
    if not rows:return {'status':'needs_input','missing_fields':['current_store_rows'],'output_files':[]}
    sync['rows']=rows
    sync['source']={k:v for k,v in payload['source'].items() if k not in {'token','password','secret'}}
    sync['cleaned_sha256']=hashlib.sha256(json.dumps(rows,ensure_ascii=False,sort_keys=True).encode()).hexdigest()
    sync['residual_private_fields']=[k for v in rows for k in v if k.lower() in base.PRIVATE]
    sync['downstream_allowed']=not sync['residual_private_fields'] and payload.get('source_confirmed') is True
    sync['human_confirmation']='confirmed' if sync['downstream_allowed'] else 'pending'
    return sync

def source_files(payload,sync,action):
    folder=Path('dist');folder.mkdir(exist_ok=True)
    record={k:v for k,v in sync.items() if k!='rows'}
    records=[v for v in sync['rows'] if payload['start']<=v['date']<=payload['end']]
    if not records:return {'status':'needs_input','missing_fields':['selected_period_rows'],'output_files':[]}
    import matplotlib
    matplotlib.use('Agg')
    from matplotlib import pyplot as plt
    if action=='sync':
        base.workbook(folder/'接入清单.xlsx',[('来源',[sync['source']]),('字段完整率',[{'coverage':sync['coverage']}]),('每日到数',[{'date':v['date'],'rows':1} for v in records]),('隐私剔除',[{'field':k} for k in sync['removed_fields']]),('质量及补齐',sync['quality'])])
        for name,values in [('字段完整率',[sync['coverage']]),('每日到数',[1 for _ in records])]:
            fig,ax=plt.subplots(figsize=(8,3));ax.bar(range(len(values)),values);ax.set_title('Field coverage' if name=='字段完整率' else 'Received rows');fig.tight_layout();fig.savefig(folder/(name+'.png'),dpi=120);plt.close(fig)
    else:
        collection={'status':'draft','downstream_allowed':sync['downstream_allowed'],'human_confirmation':sync['human_confirmation'],'source_sha256':sync['source_sha256'],'cleaned_sha256':sync['cleaned_sha256'],'source':sync['source'],'residual_private_fields':sync['residual_private_fields']}
        (folder/'collect.json').write_text(json.dumps(collection,ensure_ascii=False,indent=2),'utf-8')
        base.workbook(folder/'采集归集单.xlsx',[('归集概览',[collection]),('已归集',records),('待补齐',sync['quality']),('已剔除',[{'field':k} for k in sync['removed_fields']]),('门禁',[{'downstream_allowed':sync['downstream_allowed']}])])
    (folder/'sync.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),'utf-8')
    return {'status':'draft','action':action,'sync':record,'output_files':sorted(p.name for p in folder.iterdir()),'metrics_analysis_performed':False,'downstream_allowed':sync['downstream_allowed'],'human_confirmation':sync['human_confirmation'],'data_origin':payload.get('data_origin'),'business_completed':False}

def copy_from_checked(payload):
    checked=payload.get('checked_metrics');evidence=payload.get('confirmed_anomalies')
    if not isinstance(checked,dict) or not isinstance(evidence,list):return {'status':'needs_input','missing_fields':['checked_metrics','confirmed_anomalies'],'output_files':[]}
    daily=checked.get('daily')
    if not isinstance(daily,list) or not daily or not checked.get('source_sha256'):return {'status':'needs_input','missing_fields':['checked_metrics.daily','checked_metrics.source_sha256'],'output_files':[]}
    if any(v.get('human_confirmation')!='confirmed' for v in evidence):return {'status':'blocked','reason':'文案只使用确认异常；候选不能自动升格','output_files':[]}
    required={'date','revenue','orders','visitors','refunds','refund_rate','conversion_rate','average_order_value','revenue_yoy','revenue_same_weekday','yoy_dod_conflict'}
    if any(not required<=set(v) for v in daily):return {'status':'needs_input','missing_fields':['checked_metric_fields'],'output_files':[]}
    sync={'source_sha256':checked['source_sha256'],'coverage':1.0,'quality':[],'rows':daily,'removed_fields':[]}
    result=base.export_files(sync,{'daily':daily,'alerts':evidence,'thresholds':checked.get('thresholds',{})},payload)
    result['detector_rerun']=False
    return result

def weekly(payload,sync,detected):
    days=detected['daily'];actual_days={v['date'] for v in days};start=date.fromisoformat(payload['start']);end=date.fromisoformat(payload['end'])
    if (end-start).days!=6 or len(actual_days)!=7:return {'status':'needs_input','missing_fields':['complete_seven_day_week'],'output_files':[]}
    result=base.export_files(sync,detected,payload);folder=Path('dist');record=json.loads((folder/'weekly_flow.json').read_text('utf-8'))
    previous_start=start-timedelta(days=7);previous_end=start-timedelta(days=1)
    old=[v for v in sync['rows'] if str(previous_start)<=v['date']<=str(previous_end)]
    if len({v['date'] for v in old})==7:
        totals={k:round(sum(v[k] for v in old),2) for k in ['revenue','orders','visitors','refunds']}
        totals.update(average_order_value=base.ratio(totals['revenue'],totals['orders']),conversion_rate=base.ratio(totals['orders'],totals['visitors']),refund_rate=base.ratio(totals['refunds'],totals['revenue']))
        record.update(previous_week=totals,week_change=base.change(record['totals']['revenue'],totals['revenue']),comparison_state='actual_same_store_seven_day_previous_week')
    # Only explicitly confirmed source anomalies may win the high-alert priority.
    confirmed=payload.get('confirmed_anomaly_ids',[])
    high=[v for v in detected['alerts'] if v['state']=='candidate' and v['level']=='高' and v['id'] in confirmed]
    total=record['totals']
    if high:priority='已确认高等级异常';action='值班负责人针对已确认高异常逐项核查原始明细，维持零预算边界。'
    elif total['refund_rate'] is not None and total['refund_rate']>.03:priority='退款率超过3%';action='售后负责人复核本周退款原因与订单期间，形成一项可验证的改善动作。'
    elif total['conversion_rate'] is not None and total['conversion_rate']<.02:priority='转化率低于2%';action='门店负责人核对本周访客与订单统计口径，再检查接待流程。'
    elif record['week_change'] is not None and record['week_change']<=-.10:priority='周营收下降至少10%';action='门店负责人核查本周与上周商品和渠道结构，缺效率事实不擅自调整渠道预算。'
    else:priority='维持现有节奏';action='维持现有节奏；值班负责人持续核对同口径到数，不硬给经营调整。'
    structure=[]
    for weekday in range(7):
        subset=[v for v in days if date.fromisoformat(v['date']).weekday()==weekday]
        structure.append({'weekday':weekday,'sample_days':len(subset),'average_revenue':sum(v['revenue'] for v in subset)/len(subset) if subset else None,'average_orders':sum(v['orders'] for v in subset)/len(subset) if subset else None,'single_day_limit':len(subset)==1})
    record.update(priority=priority,action=action,owner=payload.get('owner','值班负责人'),weekday_structure=structure,budget_cny=payload.get('budget_cny',0),data_origin=payload.get('data_origin'),source_sha256=sync['source_sha256'])
    (folder/'weekly_flow.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),'utf-8')
    from docx import Document
    doc=Document();doc.add_heading('周度经营建议 · 仿真验收草稿' if payload.get('data_origin')=='SYNTHETIC' else '周度经营建议 · 待确认',0)
    for title,value in [('经营汇总',total),('同口径上周',record['previous_week']),('星期结构',structure),('唯一行动',{'priority':priority,'action':action,'owner':record['owner'],'budget_cny':record['budget_cny']}),('来源及限制',{'source_sha256':sync['source_sha256'],'data_origin':payload.get('data_origin'),'human_confirmation':'pending'})]:doc.add_heading(title,1);doc.add_paragraph(json.dumps(value,ensure_ascii=False))
    doc.save(folder/'周度经营建议.docx')
    base.workbook(folder/'周度汇总.xlsx',[('本周汇总',[total]),('星期结构',structure),('重点行动',[{'priority':priority,'action':action,'budget_cny':record['budget_cny']}]),('同口径上周',[record['previous_week'] or {'state':'缺数据不估算'}]),('来源',[{'source_sha256':sync['source_sha256'],'data_origin':payload.get('data_origin')}])])
    from matplotlib import pyplot as plt
    fig,ax=plt.subplots(figsize=(7,3));labels=['Current week'];values=[total['revenue']]
    if record['previous_week'] is not None:labels.insert(0,'Previous week');values.insert(0,record['previous_week']['revenue'])
    ax.bar(labels,values);ax.set_ylabel('Revenue CNY');fig.tight_layout();fig.savefig(folder/'周对比图.png',dpi=120);plt.close(fig)
    result['weekly']=record
    return result

def execute(payload,context=None):
    if not isinstance(payload,dict) or payload.get('action') not in FILES:raise ValueError('未知已安装业务操作')
    action=payload['action']
    if action=='copy':result=copy_from_checked(payload)
    else:
        sync=clean(payload)
        if sync.get('status') in {'needs_input','blocked'}:return sync
        if action in {'sync','collect'}:return source_files(payload,sync,action)
        # Keep individual source collection usable before a report is generated.
        detected=base.detect(sync['rows'],payload['store_id'],payload['start'],payload['end'],holidays=payload.get('holidays',()))
        if not detected['daily']:return {'status':'needs_input','missing_fields':['selected_period_rows'],'output_files':[]}
        if action in {'daily','weekly','alert'} and not sync['downstream_allowed']:return {'status':'needs_input','missing_fields':['current_source_confirmation'],'source_sha256':sync['source_sha256'],'cleaned_sha256':sync['cleaned_sha256'],'output_files':[]}
        result=weekly(payload,sync,detected) if action=='weekly' else base.export_files(sync,detected,payload)
        if result['status']=='needs_input':return result
        folder=Path('dist')
        source_record={k:v for k,v in sync.items() if k!='rows'}
        (folder/'sync.json').write_text(json.dumps(source_record,ensure_ascii=False,indent=2),'utf-8')
        if action=='collect':
            collection={'status':'draft','downstream_allowed':sync['downstream_allowed'],'human_confirmation':sync['human_confirmation'],'source_sha256':sync['source_sha256'],'cleaned_sha256':sync['cleaned_sha256'],'source':sync['source'],'residual_private_fields':sync['residual_private_fields']}
            (folder/'collect.json').write_text(json.dumps(collection,ensure_ascii=False,indent=2),'utf-8')
            base.workbook(folder/'采集归集单.xlsx',[('归集概览',[collection]),('已归集',sync['rows']),('待补齐',sync['quality']),('已剔除',[{'field':f} for f in sync['removed_fields']]),('门禁',[{'downstream_allowed':sync['downstream_allowed'],'reason':'绑定本次来源核对；不授予外部连接权限'}])])
        result.update(sync=source_record,detected=detected,downstream_allowed=sync['downstream_allowed'])
    if result.get('status') in {'needs_input','blocked'}:return result
    folder=Path('dist')
    for file in folder.iterdir():
        if file.name not in FILES[action]:file.unlink()
    result.update(action=action,output_files=sorted(p.name for p in folder.iterdir()),business_completed=False,human_confirmation='pending')
    return result
