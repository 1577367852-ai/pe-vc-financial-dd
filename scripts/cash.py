"""Dated cash bridge, payment identities, staged financing and milestone-linked scenarios."""
from copy import deepcopy
import calendar
from datetime import date, timedelta
from common import add_months, dec, month_ends, require_source, result


def forecast_ends(baseline, months, partial_policy=None):
    """N statement periods: a midmonth start includes its remaining calendar month."""
    base = date.fromisoformat(baseline)
    last = calendar.monthrange(base.year, base.month)[1]
    if base.day == last:
        return list(month_ends(baseline, months))
    if partial_policy != 'actual_days_month_end':
        raise ValueError('月中基准须明确actual_days_month_end口径及持续费用分摊依据')
    first = base.replace(day=last)
    return [add_months(first, i) for i in range(months)]


def opening_cash(book_cash, restricted_cash, cutoff, baseline, movements=None,
                 complete=False, coverage_source=None, assumption=None):
    cutoff, baseline = date.fromisoformat(cutoff), date.fromisoformat(baseline)
    if baseline < cutoff:
        return result('BLOCKED_INPUT', reason='分析基准日早于财务截止日，不能逆推')
    try:
        book, restricted = dec(book_cash), dec(restricted_cash)
        if restricted < 0 or restricted > book:
            raise ValueError('受限现金范围无效')
        initial = book - restricted
        if baseline == cutoff:
            return result('PASS', initial, 'book_cash - restricted_cash',
                          {'book_cash': book, 'restricted_cash': restricted}, cash_as_of=cutoff, label='截止日可支配现金')
        if not complete or not coverage_source:
            if assumption is None:
                return result('HISTORICAL_ONLY', initial, 'book_cash - restricted_cash',
                              cash_as_of=cutoff, analysis_date=baseline,
                              reason='期间资料不完整，仅为历史现金，不是当前现金')
            require_source(assumption)
            delta = dec(assumption['estimated_net_change'])
            return result('ESTIMATED', initial + delta, 'historical_available_cash + assumed_net_change',
                          {'historical': initial, 'assumption': assumption}, cash_as_of=baseline,
                          reason='假设估算，未完成期间实际现金核对')
        ids, delta, rows = set(), dec(0), []
        for m in movements or []:
            require_source(m)
            if m['id'] in ids: raise ValueError('重复现金变动编号')
            ids.add(m['id'])
            d = date.fromisoformat(m['date'])
            if not cutoff < d <= baseline: raise ValueError('现金更新记录超出截止日到基准日期间')
            if m.get('status') != 'actual': raise ValueError('现金更新必须为已发生记录')
            delta += dec(m['amount'])
            rows.append(m)
        return result('PASS', initial + delta, 'historical_available_cash + actual_net_changes',
                      {'historical': initial, 'movements': rows}, cash_as_of=baseline,
                      coverage_source=coverage_source, label='期间资料更新后的基准日可支配现金；真实性未审计')
    except (ValueError, KeyError) as exc:
        return result('BLOCKED_INPUT', reason=str(exc))


def reconcile_contract(contract):
    """Accounting payable/prepaid/budget records are views, not additional cash events."""
    try:
        require_source(contract)
        total, paid = dec(contract['total']), dec(contract['paid'])
        obligations = contract['obligations']
        ids = [o['obligation_id'] for o in obligations]
        if len(ids) != len(set(ids)): raise ValueError('重复付款节点编号')
        remaining = sum((dec(o['amount']) for o in obligations), dec(0))
        if any(dec(o['amount']) < 0 for o in obligations) or paid < 0 or total < 0:
            raise ValueError('合同付款不得为负')
        delta = total - paid - remaining
        if abs(delta) > dec('.01'):
            return result('FAIL_RECONCILIATION', delta, 'total - paid - unpaid_obligations', reason='合同节点与已付款不匹配')
        return result('PASS', remaining, 'total - paid', {'contract_id': contract['contract_id'], 'version': contract['version'], 'total': total, 'paid': paid},
                      obligations=obligations, accounting_views=contract.get('accounting_views', []))
    except (KeyError, ValueError) as exc:
        return result('BLOCKED_INPUT', reason=str(exc))


def select_contracts(contracts):
    active = {}
    for c in contracts:
        if not c.get('active'): continue
        if c['contract_id'] in active: raise ValueError('同一合同存在多个有效版本')
        active[c['contract_id']] = c
    return list(active.values())


def licensing_record(nominal, received, recognized, unconditional_unpaid, contingent, source):
    vals = [dec(v) for v in (nominal, received, recognized, unconditional_unpaid, contingent)]
    if not source or any(v < 0 for v in vals): raise ValueError('授权口径缺依据或金额无效')
    return dict(zip(['nominal_maximum', 'received_cash', 'recognized_revenue', 'unconditional_unpaid', 'contingent_future'], vals), source=source,
                note='独立口径，不自动把名义总额转成应收款或现金')


def build_events(plan, scenario):
    """One source of scheduled cash events. Never infer cash from accounting expense twice."""
    p = deepcopy(plan)
    base = date.fromisoformat(p['baseline'])
    delay = int(scenario.get('delay_months', 0))
    if delay < 0: raise ValueError('延迟月数不得为负')
    milestones = {k: add_months(v, delay) for k, v in p.get('milestones', {}).items()}
    horizon = int(p['horizon_months']) + (delay if scenario.get('extend_horizon', True) else 0)
    if horizon < 1 or horizon > 120: raise ValueError('原型预测期限须为1—120个月')
    boundaries = forecast_ends(p['baseline'], horizon, p.get('partial_month_policy'))
    end = boundaries[-1]
    events, excluded, covered_ids = [], [], set()

    def add(item, origin):
        item = deepcopy(item)
        require_source(item)
        if item.get('currency', p.get('currency', 'CNY')) != p.get('currency', 'CNY'):
            raise ValueError('不自动转换币种')
        if item.get('category') == 'secondary_sale':
            excluded.append(dict(item, exclusion='老股转让款不属于公司'))
            return
        if item.get('condition') and item['condition'] not in scenario.get('conditions_met', []):
            excluded.append(dict(item, exclusion='该情景未假定条件满足'))
            return
        if item.get('category') == 'primary_financing' and not scenario.get('include_financing', False):
            excluded.append(dict(item, exclusion='无新增融资情景'))
            return
        if item.get('category') in ('licensing', 'grant') and item['id'] not in scenario.get('include_receipts', []):
            excluded.append(dict(item, exclusion='未在本情景明确纳入授权/补助'))
            return
        if item.get('timing') == 'milestone':
            if item.get('milestone') not in milestones: raise ValueError('付款依赖里程碑缺失')
            when = milestones[item['milestone']]
            when = add_months(when, int(item.get('offset_months', 0)))
        else:
            when = date.fromisoformat(item['date'])
        if when <= base:
            excluded.append(dict(item, exclusion='不在未来期间，防止截止日前已到账重复计入'))
            return
        if when > end:
            excluded.append(dict(item, exclusion='超出当前预测期限', scheduled_date=str(when)))
            return
        amount = dec(item['amount'])
        if amount < 0: raise ValueError('事件金额为正；direction决定流入流出')
        if item['direction'] not in ('in', 'out'): raise ValueError('无效收付方向')
        events.append(dict(item, date=str(when), amount=amount, origin=origin))

    for contract in select_contracts(p.get('contracts', [])):
        check = reconcile_contract(contract)
        if check['status'] != 'PASS': raise ValueError('合同金额核对未通过: ' + contract['contract_id'])
        for o in contract['obligations']:
            oid = o['obligation_id']
            if oid in covered_ids: raise ValueError('跨合同重复付款节点编号')
            covered_ids.add(oid)
            add(dict(o, id=oid, direction='out', category=o.get('category', 'opex'),
                     source=o.get('source', contract['source']), contract_id=contract['contract_id'], contract_version=contract['version']), 'contract')
    for item in p.get('events', []):
        if item.get('obligation_id') in covered_ids:
            excluded.append(dict(item, exclusion='已关联合同节点，预算/应付记录不重复计现'))
            continue
        add(item, 'event')
    for recurring in p.get('recurring', []):
        require_source(recurring)
        if recurring.get('covered_obligation_ids'):
            raise ValueError('持续费用与合同节点重叠，须先拆分预算')
        until = milestones[recurring['until_milestone']] if recurring.get('until_milestone') else end
        if (until.day != calendar.monthrange(until.year, until.month)[1] or base.day != calendar.monthrange(base.year, base.month)[1]) and p.get('partial_month_policy') != 'actual_days_month_end':
            raise ValueError('持续月度预算结束于月中，v0.1需拆为明确日期付款，不能漏算或擅自按天分摊')
        for d in boundaries:
            start = max(base + timedelta(days=1), d.replace(day=1))
            stop = min(d, until)
            if recurring.get('start_date'): start = max(start, date.fromisoformat(recurring['start_date']))
            if stop < start: continue
            days, full_days = (stop-start).days+1, calendar.monthrange(d.year,d.month)[1]
            amount = dec(recurring['amount'])
            if days != full_days:
                if p.get('partial_month_policy') != 'actual_days_month_end' or not recurring.get('proration_source'):
                    raise ValueError('非完整月持续费用缺少明确分摊及月末付款依据')
                amount = (amount * days / full_days).quantize(dec('.01'))
            add(dict(recurring, id=recurring['id'] + ':' + str(d), date=str(d), timing='fixed', amount=amount,
                     accrual_start=str(start), accrual_end=str(stop), active_days=days, month_days=full_days,
                     monthly_amount=recurring['amount'], direction='out', category='opex'), 'recurring')
    ids = [e['id'] for e in events]
    if len(ids) != len(set(ids)): raise ValueError('重复现金事件ID')
    # No same-day ordering claim. Group by date later; intraday timing is outside v0.1.
    return dict(events=sorted(events, key=lambda e: (e['date'], e['id'])), excluded=excluded,
                horizon=horizon, end=str(end), milestones={k: str(v) for k, v in milestones.items()})


def cash_forecast(initial, baseline, events, months, minimum=None, partial_policy=None):
    base_date = date.fromisoformat(baseline)
    if base_date.day != calendar.monthrange(base_date.year, base_date.month)[1] and partial_policy != 'actual_days_month_end':
        raise ValueError('v0.1月度预测起点须为月末；月中起点需先提供不完整月的明确收付安排，不自动当成整月')
    if not isinstance(months, int) or months < 1 or months > 120:
        raise ValueError('原型预测期限须为1—120个整月')
    c = dec(initial)
    floor = dec(minimum) if minimum is not None else None
    if floor is not None and floor < 0: raise ValueError('最低现金要求不能为负')
    rows, daily, seen = [], [], set()
    depletion = negative = warning = touch = None
    max_zero_gap, max_floor_gap = max(dec(0), -c), max(dec(0), (floor or dec(0)) - c)
    def observe(day, balance):
        nonlocal depletion, negative, warning, touch, max_zero_gap, max_floor_gap
        if balance <= 0 and depletion is None: depletion = day
        if balance < 0 and negative is None: negative = day
        if floor is not None:
            if balance < floor and warning is None: warning = day
            if balance == floor and touch is None: touch = day
            max_floor_gap = max(max_floor_gap, floor - balance)
        max_zero_gap = max(max_zero_gap, -balance)
    observe(baseline, c)
    grouped = {}
    boundaries = forecast_ends(baseline, months, partial_policy)
    horizon_end = boundaries[-1]
    for e in events:
        require_source(e)
        if e['id'] in seen: raise ValueError('重复事件ID')
        seen.add(e['id'])
        d = date.fromisoformat(e['date'])
        if not date.fromisoformat(baseline) < d <= horizon_end: raise ValueError('事件超出预测区间')
        if e['direction'] not in ('in', 'out') or dec(e['amount']) < 0: raise ValueError('无效收付款')
        grouped.setdefault(str(d), []).append(e)
    prior = baseline
    for end in boundaries:
        opening = c
        inflows = outflows = op_in = op_out = finance_in = dec(0)
        ids = []
        for day in sorted(d for d in grouped if prior < d <= str(end)):
            for e in grouped[day]:
                n = dec(e['amount'])
                if e['direction'] == 'in': inflows += n; c += n
                else: outflows += n; c -= n
                if e.get('category') == 'primary_financing' and e['direction'] == 'in': finance_in += n
                if e.get('category') in ('opex', 'operating_receipt'):
                    if e['direction'] == 'in': op_in += n
                    else: op_out += n
                ids.append(e['id'])
            observe(day, c)
            daily.append({'date': day, 'cash': c})
        observe(str(end), c)
        rows.append(dict(date=str(end), opening=opening, inflows=inflows, outflows=outflows,
                         closing=c, minimum=floor, gap_zero=max(dec(0), -c),
                         gap_floor=max(dec(0), floor - c) if floor is not None else None,
                         operating_net_burn=op_out-op_in, financing_inflow=finance_in, event_ids=ids))
        prior = str(end)
    return dict(status='PASS', rows=rows, daily=daily, depletion=depletion, first_negative=negative,
                first_warning=warning, threshold_touch=touch, max_zero_gap=max_zero_gap,
                max_floor_gap=max_floor_gap if floor is not None else None,
                timing_note='按输入日期、同日净额检查；月末假设不证明月内安全；负值为未覆盖需求，不是获准透支')


def forecast_plan(plan, scenario):
    opening = plan['opening']
    if opening['status'] not in ('PASS', 'ESTIMATED'):
        return result('BLOCKED_INPUT', reason='历史现金未更新至基准日；不生成当前口径预测', opening=opening)
    if str(opening['cash_as_of']) != plan['baseline']:
        return result('BLOCKED_INPUT', reason='期初现金日期与分析基准日不一致')
    try:
        built = build_events(plan, scenario)
        forecast = cash_forecast(opening['value'], plan['baseline'], built['events'], built['horizon'], plan.get('minimum'), plan.get('partial_month_policy'))
        return dict(forecast, scenario=scenario, opening_status=opening['status'], schedule=built)
    except (KeyError, ValueError, ArithmeticError) as exc:
        return result('BLOCKED_INPUT', reason='现金调度所需输入有误: '+str(exc), scenario=scenario)
