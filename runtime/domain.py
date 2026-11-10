"""Source-scoped daily operations: deterministic calculations and editable files."""
import csv,hashlib,io,json,math,statistics
from datetime import date,timedelta
from decimal import Decimal,InvalidOperation
from pathlib import Path

FIELDS=('date','revenue','orders','visitors','refunds','store_id','currency','promotion','opening_date','data_origin')
PRIVATE={'phone','mobile','email','name','customer_name','id_card','national_id','bank_account','address','license_plate',
    '手机号','手机号码','联系电话','邮箱','电子邮箱','姓名','客户姓名','身份证','身份证号','银行卡号','地址','家庭住址','车牌号'}
SHEETS=('概览','确认候选','拦截与抑制','数据质量','规则与来源')

def money(value):return Decimal(str(value)).quantize(Decimal('.01'))
def ratio(a,b):return None if not b else float(Decimal(str(a))/Decimal(str(b)))
def change(a,b):return None if b in (None,0) or a is None else float((Decimal(str(a))-Decimal(str(b)))/Decimal(str(b)))

def clean_csv(text):
    reader=csv.DictReader(io.StringIO(text.lstrip('\ufeff')));headers=reader.fieldnames or []
    removed=[field for field in headers if field.lower() in PRIVATE]
    coverage=sum(field in headers for field in FIELDS)/len(FIELDS)
    quality=[];rows=[];seen=set()
    if coverage<.5:return {'status':'blocked','coverage':coverage,'rows':[],'removed_fields':removed,'quality':[{'reason':'字段覆盖不足50%'}]}
    if coverage<.8:quality.append({'reason':'字段覆盖不足80%','missing':[f for f in FIELDS if f not in headers]})
    if not {'date','revenue','orders','visitors','refunds','store_id'}<=set(headers):
        return {'status':'needs_input','coverage':coverage,'rows':[],'removed_fields':removed,'quality':[{'reason':'核心指标缺字段','missing':sorted(set(FIELDS[:6])-set(headers))}]}
    for line,row in enumerate(reader,2):
        # Drop private fields before retaining even invalid rows or evidence.
        row={k:v for k,v in row.items() if k and k.lower() not in PRIVATE}
        try:
            day=date.fromisoformat(row['date']);key=(row['store_id'],day.isoformat())
            if key in seen:raise ValueError('duplicate_day_rows')
            values={name:Decimal(row[name]) for name in ('revenue','orders','visitors','refunds')}
            if any(not n.is_finite() or n<0 for n in values.values()):raise ValueError('negative_or_nonfinite_metric')
            if any(values[k]!=values[k].to_integral_value() for k in ('orders','visitors')):raise ValueError('noninteger_count')
            if row.get('currency','CNY')!='CNY':raise ValueError('mixed_currency_requires_mapping')
            seen.add(key)
            record={**row,'date':day.isoformat(),**{k:float(v) for k,v in values.items()}}
            record.update(average_order_value=ratio(values['revenue'],values['orders']),conversion_rate=ratio(values['orders'],values['visitors']),refund_rate=ratio(values['refunds'],values['revenue']))
            rows.append(record)
        except (ValueError,InvalidOperation,TypeError) as error:quality.append({'line':line,'reason':str(error),'safe_row':row})
    rows.sort(key=lambda r:(r['store_id'],r['date']))
    for store in sorted({r['store_id'] for r in rows}):
        daily=[date.fromisoformat(r['date']) for r in rows if r['store_id']==store]
        for previous,current in zip(daily,daily[1:]):
            if (current-previous).days>1:quality.append({'store_id':store,'reason':'missing_calendar_date','start':str(previous+timedelta(days=1)),'end':str(current-timedelta(days=1)),'days':(current-previous).days-1,'filled_with_zero':False})
    return {'status':'draft' if rows else 'needs_input','coverage':coverage,'rows':rows,'removed_fields':removed,'quality':quality,'source_sha256':hashlib.sha256(text.encode()).hexdigest()}

def detect(rows,store,start,end,*,holidays=(),thresholds=None):
    thresholds={'revenue_decline':.20,'conversion_low':.015,'refund_high':.05,**(thresholds or {})}
    by_day={r['date']:r for r in rows if r['store_id']==store};alerts=[];daily=[]
    def emit(row,rule,metric,value,reference,level,reason,check):
        day=date.fromisoformat(row['date']);protect=[]
        if row.get('promotion','').lower()=='true' or row['date'] in holidays:protect.append('特殊期保护')
        try:
            if (day-date.fromisoformat(row['opening_date'])).days<30:protect.append('新店30天保护')
        except (ValueError,KeyError):protect.append('开业日期未提供')
        if metric in ('conversion_rate','refund_rate') and row['orders']<20:protect.append('不足20单只报比率值')
        alerts.append({'id':row['date']+'-'+rule+'-'+metric,'date':row['date'],'rule':rule,'metric':metric,'actual':value,'reference':reference,
            'level':level,'state':'suppressed' if protect else 'candidate','suppression_reasons':protect,'reason':reason,'investigation':check,'source_date':row['date']})
    selected=[r for r in by_day.values() if start<=r['date']<=end]
    for row in selected:
        day=date.fromisoformat(row['date']);week=by_day.get(str(day-timedelta(days=7)));prior=by_day.get(str(day-timedelta(days=1)))
        try:yoy=by_day.get(str(day.replace(year=day.year-1)))
        except ValueError:yoy=None
        r={**row,'revenue_same_weekday':change(row['revenue'],week['revenue']) if week else None,'revenue_dod':change(row['revenue'],prior['revenue']) if prior else None,
            'revenue_yoy':change(row['revenue'],yoy['revenue']) if yoy else None,'year_reference_date':yoy['date'] if yoy else None,'weekday_reference_date':week['date'] if week else None}
        r['yoy_dod_conflict']=r['revenue_yoy'] is not None and r['revenue_dod'] is not None and r['revenue_yoy']*r['revenue_dod']<0
        daily.append(r)
        preferred=r['revenue_yoy'] if r['revenue_yoy'] is not None else r['revenue_same_weekday']
        reference=yoy if r['revenue_yoy'] is not None else week
        r['effective_comparison']='year_on_year' if r['revenue_yoy'] is not None else 'same_weekday' if week else 'unavailable'
        if preferred is not None and preferred<=-thresholds['revenue_decline']:
            emit(row,'R1','revenue',row['revenue'],{'date':reference['date'],'revenue':reference['revenue'],'change':preferred,'comparison':r['effective_comparison']},'高','营收下降至少20%，有去年同期优先同比，否则同星期对比','值班负责人核查本日到数、停业与促销口径，不先认定经营原因')
        for metric,key,test,level in [('conversion_rate','conversion_low',lambda x,t:x<t,'中'),('refund_rate','refund_high',lambda x,t:x>t,'高')]:
            value=row[metric]
            if value is not None and test(value,thresholds[key]):emit(row,'R1',metric,value,{'threshold':thresholds[key]},level,'比率阈值候选','值班负责人对照原订单/访客或退款明细，确认分母口径及退款所属期间')
        history=[by_day[str(day-timedelta(days=n))]['revenue'] for n in range(1,29) if str(day-timedelta(days=n)) in by_day]
        if len(history)>=14:
            mean=statistics.mean(history);sd=statistics.pstdev(history)
            if sd and abs(row['revenue']-mean)>2*sd:emit(row,'R2','revenue',row['revenue'],{'method':'2sigma','count':len(history),'mean':mean,'sd':sd},'中','历史分布两倍标准差外','核查历史覆盖和节假日，检查商品/渠道构成后再判断原因')
        elif len(history)>=4:
            q1,_,q3=statistics.quantiles(history,n=4,method='inclusive');iqr=q3-q1
            if iqr and not q1-1.5*iqr<=row['revenue']<=q3+1.5*iqr:emit(row,'R2','revenue',row['revenue'],{'method':'IQR','count':len(history),'q1':q1,'q3':q3},'低','不足14天回退IQR','补齐历史，并复核少样本分布；不将短期极值直接认定原因')
        sequence=[by_day.get(str(day-timedelta(days=n))) for n in range(2,-1,-1)]
        if all(sequence):
            movement=change(sequence[-1]['revenue'],sequence[0]['revenue'])
            directional=all(a['revenue']>b['revenue'] for a,b in zip(sequence,sequence[1:])) or all(a['revenue']<b['revenue'] for a,b in zip(sequence,sequence[1:]))
            if directional and movement is not None and abs(movement)>.10:
                emit(row,'R3','revenue',row['revenue'],{'start':sequence[0]['date'],'start_revenue':sequence[0]['revenue'],'change':movement},'中','连续至少三天同向且累计幅度超过10%','按日核原到数与活动，缺日断开，不补零制造连续趋势')
        if r['yoy_dod_conflict']:
            emit(row,'R4','revenue',row['revenue'],{'year_date':yoy['date'],'yoy':r['revenue_yoy'],'dod':r['revenue_dod']},'中','同比与环比方向冲突，解释优先展示同比且保留环比','核同店去年同期范围及经营日条件，不能只选择看起来更好的增长口径')
    return {'status':'draft' if daily else 'needs_input','daily':daily,'alerts':alerts,'thresholds':thresholds,'rules':'source-bound R1/R2/R3/R4; controlled internal rules, not customer facts','human_confirmation':'pending'}

def safe_cell(cell,value):
    cell.value=json.dumps(value,ensure_ascii=False) if isinstance(value,(dict,list)) else value
    if isinstance(cell.value,str) and cell.value.startswith(('=','+','-','@')):cell.data_type='s'

def workbook(path,sheets):
    from openpyxl import Workbook
    book=Workbook();book.remove(book.active)
    for name,rows in sheets:
        sheet=book.create_sheet(name[:31]);headers=list(dict.fromkeys(k for row in rows for k in row)) if rows else ['状态']
        sheet.append(headers)
        for values in rows:
            for col,key in enumerate(headers,1):safe_cell(sheet.cell(sheet.max_row+1,col) if col==1 else sheet.cell(sheet.max_row,col),values.get(key))
        sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
        for col in sheet.columns:sheet.column_dimensions[col[0].column_letter].width=min(55,max(14,len(str(col[0].value))+4))
    book.save(path)

def export_files(sync,detected,payload):
    import matplotlib
    matplotlib.use('Agg')
    from matplotlib import pyplot as plt
    from docx import Document
    from docx.shared import Inches
    out=Path('dist');out.mkdir(exist_ok=True)
    daily=detected['daily'];alerts=detected['alerts'];last=daily[-1];candidates=[a for a in alerts if a['state']=='candidate'];suppressed=[a for a in alerts if a['state']=='suppressed']
    figure,axis=plt.subplots(figsize=(10,4));other=axis.twinx();x=list(range(len(daily)))
    axis.plot(x,[r['revenue'] for r in daily],label='Revenue (CNY)',color='#1466a4');other.plot(x,[r['orders'] for r in daily],label='Orders',color='#c67b26')
    axis.set_xticks(x,[r['date'][5:] for r in daily],rotation=30);axis.set_ylabel('Revenue CNY');other.set_ylabel('Orders');figure.tight_layout();figure.savefig(out/'双轴趋势.png',dpi=130);plt.close(figure)
    for name,values in [('字段完整率',[sync['coverage']]),('每日到数',[1 for _ in daily])]:
        figure,axis=plt.subplots(figsize=(8,3));axis.bar(range(len(values)),values);axis.set_title('Field coverage' if len(values)==1 else 'Received daily rows');figure.tight_layout();figure.savefig(out/(name+'.png'),dpi=120);plt.close(figure)
    selected_numbers=[{'metric':'营收','value':last['revenue'],'reference':{'same_weekday_change':last['revenue_same_weekday'],'yoy':last['revenue_yoy']}}]
    for metric in ['refund_rate','conversion_rate','average_order_value']:
        if len(selected_numbers)<3:selected_numbers.append({'metric':metric,'value':last[metric],'reference':'无已确认健康区间；不编对比'})
    yoy='缺去年同期，不计算同比' if last['revenue_yoy'] is None else f"同比{last['revenue_yoy']:.1%}"
    conclusion=f"营收{last['revenue']:.2f}元；{yoy}。"
    if len(conclusion)>40:raise ValueError('日报结论超过40字')
    body=conclusion+('同比与环比方向相反，先核同店口径。' if last['yoy_dod_conflict'] else '')+'以下为仿真台账的辅助分析，异常候选及归因待人工复核。'
    action='值班负责人核查本店本日原始到数和去年同期口径，确认候选异常后再选经营动作。'
    copy={'status':'draft','conclusion':conclusion,'body':body,'key_numbers':selected_numbers,'action':action,'owner':payload.get('owner','值班负责人'),
        'actual_causes_proven':False,'human_confirmation':'pending','data_origin':payload.get('data_origin','user_supplied_unverified')}
    summary={'source_sha256':sync['source_sha256'],'range':[daily[0]['date'],daily[-1]['date']],'data_origin':payload.get('data_origin','user_supplied_unverified'),'selected_days':len(daily),'candidate_alerts':len(candidates),'suppressed_alerts':len(suppressed),'human_confirmation':'pending'}
    common=[('概览',[summary]),('确认候选',candidates),('拦截与抑制',suppressed),('数据质量',sync['quality']),('规则与来源',[{'rule':k,'value':v} for k,v in detected['thresholds'].items()])]
    workbook(out/'异常预警清单.xlsx',common)
    workbook(out/'日报数据摘要.xlsx',[('关键数',selected_numbers),('全部指标',daily),('异常证据',alerts),('数据质量',sync['quality']),('来源边界',[summary])])
    workbook(out/'经营日报.xlsx',[('经营概览',[summary]),('每日明细',daily),('异常证据',alerts),('质量与拦截',sync['quality']+suppressed),('来源口径',[{'source_sha256':sync['source_sha256'],'data_origin':summary['data_origin']}])])
    workbook(out/'接入清单.xlsx',[('来源',[summary]),('字段完整率',[{'coverage':sync['coverage'],'missing':[f for f in FIELDS if f not in (sync['rows'][0] if sync['rows'] else {})]}]),('每日到数',[{'date':r['date'],'rows':1} for r in daily]),('隐私剔除',[{'removed_field':f} for f in sync['removed_fields']]),('质量及补齐',sync['quality'])])
    workbook(out/'采集归集单.xlsx',[('归集概览',[summary]),('已归集',daily),('待补齐',sync['quality']),('已剔除',[{'field':f} for f in sync['removed_fields']]),('门禁',[{'downstream_allowed':False,'reason':'原始资料人工确认待完成；产物为验收草稿'}])])
    doc=Document();doc.add_heading('经营日报 · 仿真验收草稿',0)
    for title,text in [('一、核心结论',conclusion),('二、关键指标',json.dumps(selected_numbers,ensure_ascii=False)),('三、趋势与异常证据','异常为候选；抑制和缺日保留，原因未证实。'),('四、明日排查',action),('五、来源与边界',json.dumps(summary,ensure_ascii=False))]:
        doc.add_heading(title,level=1);doc.add_paragraph(text)
        if title.startswith('三'):doc.add_picture(str(out/'双轴趋势.png'),width=Inches(6))
    doc.save(out/'经营日报.docx')
    # Weekly output has its own aggregation and exactly one priority action.
    totals={k:round(sum(r[k] for r in daily),2) for k in ['revenue','orders','visitors','refunds']};totals.update(conversion_rate=ratio(totals['orders'],totals['visitors']),refund_rate=ratio(totals['refunds'],totals['revenue']),average_order_value=ratio(totals['revenue'],totals['orders']))
    week_action='维持现有经营，值班负责人优先补齐同口径上周数据。';priority='维持/缺比较'
    if any(a['level']=='高' for a in candidates):priority='高异常候选待确认';week_action='值班负责人核对本周高风险候选的原始到数；确认前不执行渠道或预算调整。'
    elif totals['refund_rate'] is not None and totals['refund_rate']>.03:priority='退款率超过3%';week_action='售后负责人复核本周退款原因与订单期间，形成一项可验证的改善动作。'
    elif totals['conversion_rate'] is not None and totals['conversion_rate']<.02:priority='转化率低于2%';week_action='门店负责人核对本周访客与订单统计口径，再检查接待流程。'
    weekly={'status':'draft','totals':totals,'previous_week':None,'week_change':None,'comparison_state':'缺同口径上周汇总，不估算',
        'weekday_structure':[{'weekday':i,'sample_days':sum(date.fromisoformat(r['date']).weekday()==i for r in daily)} for i in range(7)],
        'priority':priority,'action':week_action,'budget_cny':0,'external_actions_executed':False,'human_confirmation':'pending'}
    weekly_doc=Document();weekly_doc.add_heading('周度经营建议 · 仿真验收草稿',0);weekly_doc.add_paragraph(json.dumps(weekly,ensure_ascii=False));weekly_doc.save(out/'周度经营建议.docx')
    workbook(out/'周度汇总.xlsx',[('本周汇总',[totals]),('星期结构',weekly['weekday_structure']),('重点行动',[{'priority':priority,'action':week_action,'budget_cny':0}]),('上周缺口',[{'comparison_state':weekly['comparison_state']}]),('来源',[summary])])
    figure,axis=plt.subplots(figsize=(7,3));axis.bar(['Current week'],[totals['revenue']]);axis.set_ylabel('Revenue CNY');axis.set_title('Previous comparable week unavailable');figure.tight_layout();figure.savefig(out/'周对比图.png',dpi=120);plt.close(figure)
    previews=[];dedup={}
    for a in candidates:
        key=(a['date'],a['metric'])
        if key in dedup:
            dedup[key]['supporting_evidence'].append(a['id']);continue
        preview={'date':a['date'],'metric':a['metric'],'actual':a['actual'],'evidence':a['id'],'supporting_evidence':[a['id']],
            'status':'manual_preview_not_sent','delivery_receipt':None,'escalation':'30分钟未确认升级负责人，仅为待执行规则','quiet_hours':'静默至次日07:00'}
        dedup[key]=preview;previews.append(preview)
    workbook(out/'预警推送单.xlsx',[('候选预览',previews),('抑制台账',suppressed),('渠道配置',[{'configured':False,'automatic_send':False}]),('去重规则',[{'window_hours':12,'delivery_count':0}]),('来源',[summary])])
    (out/'推送消息预览.md').write_text('# 预警消息草稿，未发送\n\n'+'\n'.join(json.dumps(v,ensure_ascii=False) for v in previews),'utf-8')
    records={'sync.json':{k:v for k,v in sync.items() if k!='rows'},'detect.json':detected,'report_copy.json':copy,
        'collect.json':{'status':'draft','downstream_allowed':False,'human_confirmation':'pending','source_sha256':sync['source_sha256']},
        'report_flow.json':{'status':'draft','source_sha256':sync['source_sha256'],'copy':copy,'detector_is_actual_script':True},
        'weekly_flow.json':weekly,'alert_flow.json':{'status':'draft','previews':previews,'pushed_list_count':len(previews),'actually_delivered':0,'automatic_send':False}}
    for name,value in records.items():(out/name).write_text(json.dumps(value,ensure_ascii=False,indent=2,allow_nan=False),'utf-8')
    return {'status':'draft','source_sha256':sync['source_sha256'],'metrics':daily,'summary':summary,'output_files':sorted(f.name for f in out.iterdir()),'business_completed':False,'human_confirmation':'pending','actual_delivery_count':0}

def execute(payload,context=None):
    if not isinstance(payload,dict):raise ValueError('输入不是对象')
    if not payload.get('csv_text'):return {'status':'needs_input','missing_fields':['csv_text'],'output_files':[]}
    if len(payload['csv_text'])>1000000:raise ValueError('CSV超过当前1MB边界')
    for field in ['start','end','store_id']:
        if not payload.get(field):return {'status':'needs_input','missing_fields':[field],'output_files':[]}
    start=date.fromisoformat(payload['start']);end=date.fromisoformat(payload['end'])
    if start>end:raise ValueError('日期区间反转')
    sync=clean_csv(payload['csv_text'])
    if sync['status'] in ['blocked','needs_input']:return {**sync,'output_files':[]}
    invalid=[q for q in sync['quality'] if 'line' in q]
    if invalid and payload.get('allow_invalid_rows') is not True:
        return {'status':'needs_input','missing_fields':['review_invalid_source_rows'],'quality':invalid,'output_files':[],'source_sha256':sync['source_sha256']}
    detected=detect(sync['rows'],payload['store_id'],str(start),str(end),holidays=payload.get('holidays',()))
    if not detected['daily']:return {'status':'needs_input','missing_fields':['selected_store_and_period_rows'],'output_files':[]}
    return export_files(sync,detected,payload)
