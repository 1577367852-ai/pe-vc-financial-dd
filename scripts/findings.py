"""Evidence-linked questions from observed calculation results, not investment decisions."""
from common import dec


def build_findings(report):
    rows=[]
    def add(key, fact, amount, source, impact, request, kind='analysis'):
        rows.append(dict(finding_id=key,kind=kind,fact=fact,amount=amount,source=source,impact=impact,request=request))
    cash=report.get('cash',{})
    if cash.get('status')=='PASS':
        if dec(cash['max_zero_gap'])>0:
            add('cash-gap',f'情景“{cash["scenario"]["name"]}”首次负余额为{cash["first_negative"]}，最大资金缺口如下。',
                cash['max_zero_gap'],{'event_ids':[e['id'] for e in cash['schedule']['events']], 'opening_as_of':str(report['opening'].get('cash_as_of'))},
                '需要在缺口发生前安排到账；期末余额不能替代期间资金需求。',
                '取得首次缺口日前的付款清单、融资交割先决条件及可验证到账时间。')
        for row in cash['rows']:
            if row['financing_inflow']>0 and row['operating_net_burn']>0:
                add('financing-burn:'+row['date'],f'{row["date"]}期间收到增资{row["financing_inflow"]}元，但经营仍净消耗{row["operating_net_burn"]}元。',
                    row['operating_net_burn'],{'event_ids':row['event_ids']},'融资改善流动性，不代表经营盈亏改善。',
                    '核对新增资金用途、经营预算和下一轮融资所需研发节点。')
        for e in cash['schedule'].get('excluded',[]):
            if e.get('category') in ('primary_financing','licensing','grant'):
                add('excluded:'+e['id'],f'未计入现金：{e["id"]}；{e["exclusion"]}。',e.get('amount'),e.get('source'),
                    '潜在资金不能替代已到账现金；可能改变融资规模和付款安排。',
                    f'核对{e["id"]}的合同、付款条件、金额和最晚到账日期。','unverified')
    else:
        add('cash-blocked','无法生成当前口径现金预测：'+cash.get('reason','缺输入'),None,report['opening'],
            '目前不能据此判断当前现金续航或融资缺口。','补齐截止日至基准日现金流水、受限资金及完整性说明。','unverified')
    for k,r in report['checks'].items():
        if r['status']=='FAIL_RECONCILIATION':
            add('reconciliation:'+k,f'{k}勾稽差额超出容差。',r['value'],r.get('source_record_ids',{}),
                '差异可能来自范围、期间或会计记录；不能通过无依据调整配平。',
                f'提供{k}的科目明细、调整分录及差额解释。','fact')
    for bucket,r in report.get('bridges',{}).items():
        if r['status']=='PASS':
            delta=r['value']-dec(r['base'])
            if delta:
                impact={'QoE':'影响所选利润口径，不自动折算估值；需区分一次性和持续性事项。',
                        'NWC':'影响正常经营资金占用及交割营运资金讨论；不自动当成新增现金支出。',
                        'NetDebt':'影响净债务口径及企业价值到股权价值的桥接；不自动改变估值。'}[bucket]
                add('bridge:'+bucket,f'{bucket}由{r["base"]}元调整为{r["value"]}元。',delta,{'base_source':r.get('base_source'),'adjustment_ids':[x['adjustment_id'] for x in r['inputs']['items']]},impact,
                    '复核已采纳调整的会计依据、经济事项编号及是否与其他价格调整重复。')
        elif r['status']=='BLOCKED_INPUT':
            add('bridge-blocked:'+bucket,f'{bucket}调整桥未完成：'+r['reason'],None,r.get('base_source'),
                '不能将未完成口径用于盈利质量或交割价格判断。','补充调整基数、口径和重复事项解释。','unverified')
    for item in report.get('adjustments',[]):
        if item.get('status')!='supported':
            add('adjustment-pending:'+item['adjustment_id'],f'调整{item["adjustment_id"]}未采纳。',item.get('amount'),item.get('source'),
                '仍保留在待核实清单，不影响已计算的调整后金额。',
                '提供原始凭证位置、调整理由和采纳人确认。','unverified')
    for key,reason in report.get('blocked',{}).items():
        add('input:'+key,f'{key}不可用：{reason}。',None,{'account':key},
            '仅依赖该输入的指标无法计算，其他指标保持有效。',f'复核{key}的单位、期间、版本及原始数值。','unverified')
    return rows
