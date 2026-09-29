"""Reviewable CSV project entry; no company/date/layout hardcoding or input execution."""
import csv
import json
from pathlib import Path
from copy import deepcopy
from common import ROOT, allowed_write, dec, file_hash, fingerprint, write_json, write_text, result
from ingest import extract, metric_inputs, SCALES
from financial import SPECS, CHECKS, evaluate, reconcile, bridge, cross_bucket_check
from cash import opening_cash, forecast_plan
from reporting import export_bundle, assert_fresh, dependency_signature

FILES={
 '项目设置.csv':['key','value','source'],
 '材料清单.csv':['path'],
 '科目确认.csv':['record_id','decision','standard_account','entity','scope','period','currency','version','cutoff','unit','confirmed_raw','reviewer','evidence'],
 '补充映射.csv':['file','sha256','sheet','cell','page','line','bbox','original_account','standard_account','entity','scope','period','currency','version','cutoff','unit','confirmed_raw','reviewer','evidence'],
 '现金变动.csv':['id','date','amount','status','source'],
 '持续费用.csv':['id','amount','until_milestone','start_date','proration_source','source'],
 '现金事件.csv':['id','amount','date','direction','category','timing','milestone','offset_months','condition','obligation_id','source'],
 '合同节点.csv':['contract_id','version','active','total','paid','obligation_id','amount','date','timing','milestone','offset_months','source'],
 '里程碑.csv':['id','date','source'],
 '情景设置.csv':['key','value','source'],
 '调整基数.csv':['bucket','base_account','source'],
 '调整事项.csv':['adjustment_id','economic_item_id','bucket','amount','status','approved','reviewer','reason','evidence_file','evidence_sha256','evidence_location'],
}
SETTINGS=('entity','scope','period','currency','cutoff','baseline','business_model','confirmed_by',
          'confirmation_evidence','synthetic_only','complete_bridge','coverage_source','estimated_net_change',
          'estimate_source','minimum','horizon_months','partial_month_policy')
STANDARD=set(k for req,_,_ in SPECS.values() for k in req.split()) | set(k for t in CHECKS.values() for k in t) | {'interest_bearing_debt'}


def csv_save(path, fields, rows):
    p=allowed_write(path);p.parent.mkdir(parents=True,exist_ok=True)
    with p.open('w',encoding='utf-8-sig',newline='') as f:
        def safe(v):
            if isinstance(v,str) and v.lstrip().startswith(('=','+','-','@')):
                try:dec(v)
                except ValueError:return "'"+v
            return v
        w=csv.DictWriter(f,fields,extrasaction='ignore');w.writeheader()
        w.writerows({k:safe(v) for k,v in row.items()} for row in rows)


def csv_load(path):
    with Path(path).open(encoding='utf-8-sig',newline='') as f:
        rows=list(csv.DictReader(f))
    return [{k:v for k,v in r.items() if k and v not in (None,'')} for r in rows if any(r.values())]


def keyed(path):
    out={}
    for r in csv_load(path):
        key=r['key']
        if key in out:raise ValueError('重复配置项: '+key)
        out[key]=r.get('value','')
    return out


def yes(v):
    if str(v).lower() in ('true','yes','1','是'):return True
    if str(v).lower() in ('false','no','0','否',''):return False
    raise ValueError('无效布尔值: '+str(v))


def resolve_source(folder, value):
    p=Path(value)
    return p.resolve() if p.is_absolute() else (folder/p).resolve()


def sources(folder):
    paths=[resolve_source(folder,r['path']) for r in csv_load(folder/'材料清单.csv')]
    if not paths or len(set(paths))!=len(paths):raise ValueError('材料清单为空或重复')
    return paths


def prepare(folder, paths):
    """Create a new review pack only; never overwrite the original or existing pack."""
    f=allowed_write(folder)
    if f.exists():raise FileExistsError('项目目录已存在；请使用新目录，不能覆盖确认记录')
    paths=[Path(p).resolve() for p in paths]
    artifacts=[dict(path=str(p),extracted=extract(p)) for p in paths]
    records=[r for a in artifacts for r in a['extracted'].get('records',[])]
    f.mkdir(parents=True)
    write_json(f/'材料提取.json',artifacts)
    candidates=[]
    for r in records:
        candidates.append(dict(record_id=r['record_id'],decision='pending',standard_account=r.get('standard_account'),
            **{k:r.get(k) for k in ('entity','scope','period','currency','version','cutoff')},
            unit=r.get('original_unit'),confirmed_raw='',reviewer='',evidence=''))
    params={k:'' for k in SETTINGS}
    for key in ('entity','scope','period','currency','cutoff'):
        vals={r.get(key) for r in records if r.get(key)}
        if len(vals)==1:params[key]=next(iter(vals))
    params.update(synthetic_only='yes',complete_bridge='no')
    for name,cols in FILES.items():
        rows=[]
        if name=='项目设置.csv':rows=[dict(key=k,value=v,source='提取候选，需确认') for k,v in params.items()]
        elif name=='材料清单.csv':rows=[{'path':str(p)} for p in paths]
        elif name=='科目确认.csv':rows=candidates
        elif name=='情景设置.csv':rows=[dict(key=k,value=v,source='待确认') for k,v in {'name':'现有资源','include_financing':'no','delay_months':'0','extend_horizon':'yes','conditions_met':'','include_receipts':''}.items()]
        csv_save(f/name,cols,rows)
    write_text(f/'确认说明.md','# 材料确认\n\n仅合成样本。填写项目设置中的确认人、依据、业务模式和分析日期。\n'
        '科目确认每行选择include/exclude/pending，include必须填写复核人和依据；OCR还须填写confirmed_raw。\n'
        '原始值和位置见材料提取.json。未知布局由助理根据原件补映射，不要求用户写JSON。\n'
        '金额计划表统一为项目币种的元；多币种不自动换算。修改原件后先refresh，旧记录确认不会自动沿用。\n'
        '空白不是零。调整基数明确指定科目，调整事项必须有材料清单内证据、位置、理由和采纳人。\n')
    return f


def refresh(folder):
    """Re-extract after source change; retain decisions only for identical source fingerprints."""
    f=allowed_write(folder)
    old=csv_load(f/'科目确认.csv');indexed={r['record_id']:r for r in old}
    artifacts=[dict(path=str(p),extracted=extract(p)) for p in sources(f)]
    current=[]
    for a in artifacts:
        for r in a['extracted'].get('records',[]):
            current.append(indexed.get(r['record_id'],dict(record_id=r['record_id'],decision='pending',
                standard_account=r.get('standard_account'),**{k:r.get(k) for k in ('entity','scope','period','currency','version','cutoff')},unit=r.get('original_unit'))))
    archive=f/'history'/fingerprint(old)
    write_json(archive/'科目确认.json',old)
    csv_save(f/'科目确认.csv',FILES['科目确认.csv'],current)
    write_json(f/'材料提取.json',artifacts)
    if (f/'results'/'manifest.json').exists():
        m=json.loads((f/'results'/'manifest.json').read_text());m['status']='STALE'
        write_json(f/'results'/'manifest.json',m)
    return current


def mapped_records(folder, paths, config):
    records=[r for p in paths for r in extract(p).get('records',[])]
    lookup={r['record_id']:r for r in records}
    seen=set();approved=[];issues=[]
    for m in csv_load(folder/'科目确认.csv'):
        rid=m['record_id']
        if rid in seen:raise ValueError('重复映射记录: '+rid)
        seen.add(rid)
        if rid not in lookup:
            raise ValueError('材料或版本已变化，科目确认过期；先refresh再复核')
        r=deepcopy(lookup[rid]);decision=m.get('decision','pending')
        r['extracted_metadata']={k:r.get(k) for k in ('entity','scope','period','currency','version','cutoff','original_unit')}
        if decision=='exclude':continue
        if decision not in ('pending','include'):raise ValueError('无效科目确认状态')
        if decision=='include':
            if not m.get('reviewer') or not m.get('evidence'):raise ValueError('采纳映射缺确认人或依据')
            key=m.get('standard_account')
            if key not in STANDARD:raise ValueError('未支持的标准科目: '+str(key))
            for k in ('entity','scope','period','currency','version','cutoff'):
                if not m.get(k):raise ValueError('映射元数据缺失: '+k)
                r[k]=m[k]
            # Unsupported formulas cannot be promoted merely by approving a mapping.
            val=m.get('confirmed_raw')
            if r['status'] not in ('READ','RECOMPUTED','REVIEWED') and val is None:
                issues.append({'record_id':rid,'reason':'候选值需人工确认数值'});r['value']=None
            else:
                raw=val if val is not None else r['original_value']
                unit=m.get('unit')
                if unit not in SCALES:raise ValueError('单位未支持或缺失')
                if val is None and r['status']=='RECOMPUTED':
                    r['value']=r['value']/SCALES[r['original_unit']]*SCALES[unit]
                else:r['value']=dec(raw)*SCALES[unit]
                r['status']='REVIEWED'
            r['standard_account']=key
            r['review']=dict(m)
        else:
            r['value']=None;r['status']='NEEDS_REVIEW'
            issues.append({'record_id':rid,'reason':'映射待确认'})
        approved.append(r)
    for rid,r in lookup.items():
        if rid not in seen:
            r=deepcopy(r);r['value']=None;r['status']='NEEDS_REVIEW';approved.append(r)
            issues.append({'record_id':rid,'reason':'新增记录未确认'})
    approved.extend(supplemental_records(folder,paths))
    data,ids,blocked=metric_inputs(approved,config['entity'],config['scope'],config['period'],config['currency'])
    # Prevent mixing statement dates within a chosen period.
    for r in approved:
        if r.get('standard_account') in data and r.get('record_id')==ids.get(r.get('standard_account')) and r.get('cutoff')!=config['cutoff']:
            key=r['standard_account'];data.pop(key,None);blocked[key]='截止日不同'
    return approved,data,ids,blocked,issues


def supplemental_records(folder, paths):
    """Human-confirmed fallback for unrecognized layout; evidence is tied to exact source bytes."""
    known={str(p):p for p in paths};out=[]
    for m in csv_load(folder/'补充映射.csv'):
        p=resolve_source(folder,m['file'])
        if str(p) not in known or m.get('sha256')!=file_hash(p):raise ValueError('补充映射原件不在清单或版本已变化')
        for key in ('standard_account','entity','scope','period','currency','version','cutoff','unit','confirmed_raw','reviewer','evidence','original_account'):
            if key not in m:raise ValueError('补充映射缺少: '+key)
        if m['standard_account'] not in STANDARD or m['unit'] not in SCALES:raise ValueError('补充科目或单位未支持')
        src=dict(file=str(p),sha256=file_hash(p))
        if p.suffix.lower()=='.xlsx':
            from openpyxl import load_workbook
            w=load_workbook(p,data_only=False,keep_links=False)
            try:
                raw=w[m['sheet']][m['cell']].value
                if raw is None:raise ValueError('补充位置为空')
                src.update(sheet=m['sheet'],cell=m['cell'])
            finally:w.close()
        elif p.suffix.lower()=='.pdf':
            extracted=extract(p);page=int(m['page']);line=int(m['line'])
            if page<1 or line<1:raise ValueError('页或行须为正')
            raw=extracted['raw'][page-1]['text'].splitlines()[line-1]
            src.update(physical_page=page,line=line,original_line=raw)
        elif p.suffix.lower() in ('.png','.jpg','.jpeg'):
            from PIL import Image
            vals=[int(x) for x in m.get('bbox','').split(',')]
            with Image.open(p) as im:
                if len(vals)!=4 or not (0<=vals[0]<vals[2]<=im.width and 0<=vals[1]<vals[3]<=im.height):raise ValueError('补充图像坐标无效')
            raw='人工转录，须对照原图';src['bbox']=vals
        else:raise ValueError('补充映射仍不支持该格式')
        out.append(dict(record_id=fingerprint(src)[:20],original_account=m['original_account'],original_value=raw,
            original_unit=m['unit'],value=dec(m['confirmed_raw'])*SCALES[m['unit']],status='REVIEWED',
            standard_account=m['standard_account'],source=src,review=dict(m),
            **{k:m[k] for k in ('entity','scope','period','currency','version','cutoff')}))
    return out


def adjustment_bridges(folder, data, metrics, paths):
    items=[];known={str(p.resolve()) for p in paths};seen_ids=set()
    for r in csv_load(folder/'调整事项.csv'):
        item=dict(r);e=resolve_source(folder,r.get('evidence_file',''))
        if r['adjustment_id'] in seen_ids:raise ValueError('重复调整编号')
        seen_ids.add(r['adjustment_id'])
        if r['bucket'] not in ('QoE','NWC','NetDebt'):raise ValueError('调整类别无效')
        valid=(r.get('status')=='supported' and yes(r.get('approved','')) and r.get('reviewer') and
               r.get('reason') and str(e) in known and r.get('evidence_location') and r.get('evidence_sha256')==file_hash(e))
        item['status']='supported' if valid else 'unverified'
        item['source']={'file':str(e),'sha256':file_hash(e),'location':r['evidence_location']} if str(e) in known and r.get('evidence_location') else None
        if valid:dec(item['amount'])
        items.append(item)
    accepted=[r for r in items if r['status']=='supported']
    conflict=cross_bucket_check(accepted)
    conflicts=set(conflict['value'] or [])
    conflict_buckets={r['bucket'] for r in accepted if r.get('economic_item_id',r['adjustment_id']) in conflicts}
    seen=set();bridges={}
    for row in csv_load(folder/'调整基数.csv'):
        bucket=row['bucket'];key=row['base_account']
        if bucket not in ('QoE','NWC','NetDebt') or bucket in seen:raise ValueError('调整类别无效或重复')
        seen.add(bucket)
        base=data.get(key)
        if key in metrics and metrics[key]['status']=='PASS':base=metrics[key]['value']
        if key=='net_debt' and data.get('interest_bearing_debt') is not None and data.get('book_cash') is not None and data.get('restricted_cash') is not None:
            base=data['interest_bearing_debt']-(data['book_cash']-data['restricted_cash'])
        if bucket in conflict_buckets:r=result('BLOCKED_INPUT',reason='同一经济事项跨类别采纳，需先解释并拆分，禁止重复影响')
        elif base is None or not row.get('source'):r=result('BLOCKED_INPUT',reason='缺调整基数或口径依据: '+key)
        else:r=bridge(base,[x for x in items if x['bucket']==bucket],bucket)
        bridges[bucket]=dict(r,base=base,base_account=key,base_source=row.get('source'))
    for bucket in ('QoE','NWC','NetDebt'):
        if bucket not in bridges:bridges[bucket]=result('BLOCKED_INPUT',reason='未提供调整基数及口径')
    return items,bridges,conflict


def scenario_config(folder):
    s=keyed(folder/'情景设置.csv')
    return dict(name=s.get('name','未命名情景'),include_financing=yes(s.get('include_financing','')),
                delay_months=int(s.get('delay_months') or 0),extend_horizon=yes(s.get('extend_horizon','yes')),
                conditions_met=[x.strip() for x in s.get('conditions_met','').split(';') if x.strip()],
                include_receipts=[x.strip() for x in s.get('include_receipts','').split(';') if x.strip()])


def load_project(folder):
    f=Path(folder).resolve();cfg=keyed(f/'项目设置.csv')
    for k in ('entity','scope','period','currency','cutoff','baseline','business_model','confirmed_by','confirmation_evidence'):
        if not cfg.get(k):raise ValueError('项目设置待确认: '+k)
    if not yes(cfg.get('synthetic_only','')):raise ValueError('当前原型仅获准使用合成样本')
    if cfg['currency']!='CNY':raise ValueError('v0.2仅验证人民币元，不自动转换币种')
    if cfg['business_model'] not in ('研发型','授权/合作开发型','产品销售型','混合模式'):raise ValueError('业务模式需明确确认，不根据收入自动判断')
    paths=sources(f)
    records,data,ids,blocked,issues=mapped_records(f,paths,cfg)
    metrics={k:dict(evaluate(k,data),source_record_ids={d:ids.get(d) for d in SPECS[k][0].split()}) for k in SPECS}
    checks={k:dict(reconcile(k,data),source_record_ids={d:ids.get(d) for d in CHECKS[k]}) for k in CHECKS}
    items,bridges,conflict=adjustment_bridges(f,data,metrics,paths)
    assumption=None
    if cfg.get('estimated_net_change'):
        assumption=dict(estimated_net_change=cfg['estimated_net_change'],source=cfg.get('estimate_source'))
    opening=opening_cash(data.get('book_cash'),data.get('restricted_cash'),cfg['cutoff'],cfg['baseline'],
        csv_load(f/'现金变动.csv'),complete=yes(cfg.get('complete_bridge','')),coverage_source=cfg.get('coverage_source'),assumption=assumption)
    contracts={}
    for row in csv_load(f/'合同节点.csv'):
        key=(row['contract_id'],row['version'])
        attrs={k:row[k] for k in ('contract_id','version','total','paid','source')};attrs['active']=yes(row.get('active',''))
        if key not in contracts:contracts[key]=dict(attrs,obligations=[])
        elif any(contracts[key][k]!=attrs[k] for k in ('total','paid','active')):raise ValueError('合同版本总额、已付或状态不一致')
        contracts[key]['obligations'].append({k:v for k,v in row.items() if k in ('obligation_id','amount','date','timing','milestone','offset_months','source')})
    milestones={}
    for r in csv_load(f/'里程碑.csv'):
        if r['id'] in milestones or not r.get('source'):raise ValueError('里程碑重复或缺依据')
        milestones[r['id']]=r['date']
    plan=dict(baseline=cfg['baseline'],opening=opening,currency=cfg['currency'],minimum=cfg.get('minimum') or None,
        horizon_months=int(cfg.get('horizon_months') or 12),milestones=milestones,contracts=list(contracts.values()),
        recurring=csv_load(f/'持续费用.csv'),events=csv_load(f/'现金事件.csv'),partial_month_policy=cfg.get('partial_month_policy'))
    scenario=scenario_config(f)
    report=dict(name=cfg['entity'],cutoff=cfg['cutoff'],baseline=cfg['baseline'],business_model=cfg['business_model'],
        opening=opening,metrics=metrics,checks=checks,cash=forecast_plan(plan,scenario),records=records,blocked=blocked,
        mapping_issues=issues,adjustments=items,bridges=bridges,adjustment_conflicts=conflict,synthetic_only=True)
    from findings import build_findings
    report['findings']=build_findings(report)
    return report,paths+[f/n for n in FILES],scenario


def analyze(folder):
    f=allowed_write(folder);out=f/'results'
    # Invalidate BEFORE parsing: a broken new input must not leave old results looking READY.
    if (out/'manifest.json').exists():
        manifest=json.loads((out/'manifest.json').read_text());manifest['status']='RECOMPUTING'
        write_json(out/'manifest.json',manifest)
    try:
        before=dependency_signature(sources(f)+[f/n for n in FILES],scenario_config(f))
        report,paths,scenario=load_project(f)
        export_bundle(out,report,paths,scenario,expected_signature=before)
    except Exception as exc:
        write_json(out/'manifest.json',{'status':'FAILED','reason':str(exc),'old_outputs':'不可引用；未完成新计算'})
        raise
    return report


def read_current(folder, mode='work'):
    """Guarded output reader; no background monitor and no silent auto-recalculation."""
    f=allowed_write(folder);out=f/'results'
    try:
        assert_fresh(out,sources(f)+[f/n for n in FILES],scenario_config(f))
    except Exception as exc:
        if (out/'manifest.json').exists():
            m=json.loads((out/'manifest.json').read_text());m['status']='STALE';m['reason']=str(exc)
            write_json(out/'manifest.json',m)
        raise ValueError('STALE：来源、配置、代码或输出已变化；先重新分析，禁止引用旧摘要') from exc
    return (out/('学习说明.md' if mode=='learning' else '工作摘要.md')).read_text()
