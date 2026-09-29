"""Runnable local prototype: extract raw artifacts or run both synthetic end-to-end cases."""
import argparse
import json
from copy import deepcopy
from pathlib import Path
from common import ROOT, dec, serial, write_json, write_text
from cash import opening_cash, forecast_plan
from financial import SPECS, CHECKS, evaluate, reconcile, bridge
from ingest import extract, metric_inputs
from reporting import export_bundle, assert_fresh


def table(ws):
    values=list(ws.values)
    if not values:return []
    return [{k:v for k,v in zip(values[0],row) if k and v is not None} for row in values[1:] if any(v is not None for v in row)]


def load_case(kind, source=None):
    from openpyxl import load_workbook
    source=Path(source) if source else ROOT/'examples'/kind/'合成原始材料.xlsx'
    extracted=extract(source)
    name='合成研发药企' if kind=='biotech' else '合成医疗器械公司'
    data,sources,blocked=metric_inputs(extracted['records'],name,'合并','2026-H1')
    wb=load_workbook(source,data_only=True,keep_links=False)
    params={r['参数']:r['值'] for r in table(wb['分析参数'])}
    changes=table(wb['期间现金变动'])
    opening=opening_cash(data.get('book_cash'),data.get('restricted_cash'),'2026-06-30',params['baseline'],changes,
                         complete=params['complete_bridge'],coverage_source='合成期间流水完整性声明：分析参数!C3')
    contracts={}
    for row in table(wb['合同节点']):
        key=row['contract_id']
        if key not in contracts:contracts[key]=dict(contract_id=key,version=row['version'],total=row['total'],paid=row['paid'],active=True,source=row['source'],obligations=[])
        contracts[key]['obligations'].append({k:v for k,v in row.items() if k in ('obligation_id','amount','date','timing','milestone','source')})
    plan=dict(baseline=params['baseline'],opening=opening,currency='CNY',minimum=params['minimum'],horizon_months=params['horizon_months'],
              milestones={'readout':params['milestone']},recurring=table(wb['持续预算']),events=table(wb['现金事件']),contracts=list(contracts.values()))
    wb.close()
    return source,extracted,data,sources,blocked,plan,params


def make_report(kind, scenario, source=None):
    source,extracted,data,sources,blocked,plan,params=load_case(kind,source)
    metrics={k:evaluate(k,data) for k in SPECS}
    for k,res in metrics.items():res['source_record_ids']={key:sources.get(key) for key in SPECS[k][0].split()}
    checks={k:reconcile(k,data) for k in CHECKS}
    cash=forecast_plan(plan,scenario)
    return dict(name='合成研发药企' if kind=='biotech' else '合成医疗器械公司',cutoff='2026-06-30',baseline=plan['baseline'],
                business_model=params['business_model'],opening=plan['opening'],metrics=metrics,checks=checks,cash=cash,
                records=extracted['records'],blocked=blocked,
                adjustments=[{'adjustment_id':'EXAMPLE-ONLY','bucket':'QoE','status':'unverified','amount':0,'source':'无真实证据','reason':'合成演示不自动调整利润；需核实明细'}]),source,plan


def run_examples():
    done=[]
    for kind in ('biotech','sales'):
        scenarios=[dict(name='现有资源',include_financing=False,include_receipts=[],conditions_met=[],delay_months=0)]
        if kind=='biotech':scenarios += [dict(name='临床延迟三个月',include_financing=False,include_receipts=[],conditions_met=[],delay_months=3),
                                        dict(name='条件融资演示',include_financing=True,include_receipts=[],conditions_met=['data_available'],delay_months=3)]
        for i,scenario in enumerate(scenarios):
            report,source,plan=make_report(kind,scenario)
            folder=ROOT/'outputs'/f'{kind}-{i+1}'
            export_bundle(folder,report,[source],scenario)
            done.append({'case':kind,'scenario':scenario,'folder':str(folder),'cash_status':report['cash']['status']})
    write_json(ROOT/'outputs'/'示例清单.json',done)
    return done


def main():
    parser=argparse.ArgumentParser(description='本地财务原型；不得据此认定资料真实或项目可投')
    sub=parser.add_subparsers(dest='command',required=True)
    sub.add_parser('examples')
    e=sub.add_parser('extract');e.add_argument('input');e.add_argument('--output',required=True)
    f=sub.add_parser('freshness');f.add_argument('folder')
    p=sub.add_parser('prepare');p.add_argument('folder');p.add_argument('inputs',nargs='+')
    a=sub.add_parser('analyze');a.add_argument('folder')
    r=sub.add_parser('refresh');r.add_argument('folder')
    v=sub.add_parser('view');v.add_argument('folder');v.add_argument('--mode',choices=['work','learning'],default='work')
    args=parser.parse_args()
    if args.command=='examples':print(json.dumps(serial(run_examples()),ensure_ascii=False,indent=2))
    elif args.command=='extract':
        write_json(args.output,extract(args.input));print('已提取；识别不等于真实性核验，请检查状态与来源。')
    elif args.command in ('prepare','analyze','refresh','view'):
        from project import prepare,analyze,refresh,read_current
        if args.command=='prepare':print(prepare(args.folder,args.inputs))
        elif args.command=='analyze':
            report=analyze(args.folder);print('已生成本次快照；现金状态：'+report['cash']['status'])
        elif args.command=='refresh':print('刷新候选记录数：'+str(len(refresh(args.folder))))
        else:print(read_current(args.folder,args.mode))
    else:
        folder=Path(args.folder);m=json.loads((folder/'manifest.json').read_text())
        assert_fresh(folder,list(m['inputs']),m['assumptions']);print('FRESH：输入、代码、假设快照及输出指纹一致；非审计结论')


if __name__=='__main__':main()
