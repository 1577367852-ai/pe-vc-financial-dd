"""Explicit minimum algorithm registry, Decimal arithmetic and local dependency gates."""
from common import dec, result

# Inputs use one currency and explicit source IDs; units normalized to currency units.
SPECS = {
    'gross_profit': ('revenue cogs', 'revenue - cogs', '收入、成本同期间；不推断收入真实性'),
    'gross_margin': ('revenue cogs', '(revenue - cogs) / revenue', 'revenue > 0'),
    'net_margin': ('net_profit revenue', 'net_profit / revenue', 'revenue > 0'),
    'revenue_growth': ('revenue revenue_prior', 'revenue / revenue_prior - 1', '前后等长可比期间，prior > 0'),
    'current_ratio': ('current_assets current_liabilities', 'current_assets / current_liabilities', '期末同主体，分母 > 0'),
    'debt_to_assets': ('liabilities assets', 'liabilities / assets', '总负债/总资产，不等于有息负债率；assets > 0'),
    'rd_expense_share': ('rd_expense total_opex', 'rd_expense / total_opex', '总经营费用包括研发；不含营业成本；同口径，分母 > 0'),
    'ocf_margin': ('ocf revenue', 'ocf / revenue', 'revenue > 0；不是销售收现率'),
    'free_cash_flow': ('ocf cash_capex', 'ocf - cash_capex', '简化CFO后资本支出口径，不是UFCF；capex为正支出'),
    'operating_nwc': ('ar inventory other_operating_ca ap other_operating_cl',
                      'ar + inventory + other_operating_ca - ap - other_operating_cl', '须显式定义经营性范围；现金和有息债务不纳入'),
    'dso': ('ar_open ar credit_revenue period_days', '(ar_open + ar) / 2 / credit_revenue * period_days', '赊销收入 > 0；无赊销收入时不擅自用总收入'),
    'dio': ('inventory_open inventory cogs period_days', '(inventory_open + inventory) / 2 / cogs * period_days', 'cogs > 0；期初期末平均值'),
    'dpo': ('ap_open ap credit_purchases period_days', '(ap_open + ap) / 2 / credit_purchases * period_days', '赊购额 > 0，不默认用成本代替'),
    'roe': ('equity_open equity net_profit', 'net_profit / ((equity_open + equity) / 2)', '平均权益 > 0，不自动年化'),
    'dscr': ('cash_available_debt principal_due interest_due', 'cash_available_debt / (principal_due + interest_due)', '明确可供偿债现金与同期间本金利息；分母 > 0'),
    'gross_operating_burn': ('operating_outflows months', 'operating_outflows / months', '正值经营支出；排除融资和Capex；完整期间'),
    'net_operating_burn': ('operating_outflows operating_inflows months', '(operating_outflows - operating_inflows) / months', '融资不改变经营Burn；负值表示经营净流入'),
    'pre_financing_burn': ('operating_outflows operating_inflows cash_capex months', '(operating_outflows + cash_capex - operating_inflows) / months', '含现金Capex、排除融资与还本；债务支付另列'),
    'simple_runway': ('available_cash net_burn', 'available_cash / net_burn', '现金日期明确，net_burn > 0，仅稳定净消耗假设估算'),
}


def evaluate(name, data, comparable=True):
    req, formula, applicability = SPECS[name]
    keys = req.split()
    used = {k: data.get(k) for k in keys}
    if not comparable:
        return result('NOT_COMPARABLE', formula=formula, inputs=used, reason='主体、期间或币种不可比')
    missing = [k for k in keys if data.get(k) is None]
    if missing:
        return result('BLOCKED_INPUT', formula=formula, inputs=used, reason='缺失: ' + ', '.join(missing))
    try:
        x = {k: dec(data[k]) for k in keys}
        if any(x[k] <= 0 for k in ('months', 'period_days') if k in x):
            raise ValueError('期间长度须为正')
        if name in ('gross_operating_burn', 'net_operating_burn', 'pre_financing_burn', 'free_cash_flow'):
            if any(x[k] < 0 for k in ('operating_outflows', 'operating_inflows', 'cash_capex') if k in x):
                raise ValueError('流入、流出、现金资本支出须按正值提供')
        numer = denom = None
        if name == 'gross_profit': value = x['revenue'] - x['cogs']
        elif name == 'gross_margin': numer, denom = x['revenue'] - x['cogs'], x['revenue']
        elif name == 'net_margin': numer, denom = x['net_profit'], x['revenue']
        elif name == 'revenue_growth': numer, denom = x['revenue'], x['revenue_prior']
        elif name == 'current_ratio': numer, denom = x['current_assets'], x['current_liabilities']
        elif name == 'debt_to_assets': numer, denom = x['liabilities'], x['assets']
        elif name == 'rd_expense_share': numer, denom = x['rd_expense'], x['total_opex']
        elif name == 'ocf_margin': numer, denom = x['ocf'], x['revenue']
        elif name == 'free_cash_flow': value = x['ocf'] - x['cash_capex']
        elif name == 'operating_nwc': value = x['ar'] + x['inventory'] + x['other_operating_ca'] - x['ap'] - x['other_operating_cl']
        elif name == 'dso': numer, denom = (x['ar_open'] + x['ar']) / 2 * x['period_days'], x['credit_revenue']
        elif name == 'dio': numer, denom = (x['inventory_open'] + x['inventory']) / 2 * x['period_days'], x['cogs']
        elif name == 'dpo': numer, denom = (x['ap_open'] + x['ap']) / 2 * x['period_days'], x['credit_purchases']
        elif name == 'roe': numer, denom = x['net_profit'], (x['equity_open'] + x['equity']) / 2
        elif name == 'dscr': numer, denom = x['cash_available_debt'], x['principal_due'] + x['interest_due']
        elif name == 'gross_operating_burn': numer, denom = x['operating_outflows'], x['months']
        elif name == 'net_operating_burn': numer, denom = x['operating_outflows'] - x['operating_inflows'], x['months']
        elif name == 'pre_financing_burn': numer, denom = x['operating_outflows'] + x['cash_capex'] - x['operating_inflows'], x['months']
        elif name == 'simple_runway': numer, denom = x['available_cash'], x['net_burn']
        if denom is not None:
            if denom <= 0:
                return result('NOT_APPLICABLE', formula=formula, inputs=used, reason=applicability)
            value = numer / denom
            if name == 'revenue_growth': value -= 1
        return result('PASS', value, formula, used, applicability)
    except (ValueError, ArithmeticError) as exc:
        return result('BLOCKED_INPUT', formula=formula, inputs=used, reason=str(exc))


CHECKS = {
    'balance_sheet': {'assets': 1, 'liabilities': -1, 'equity': -1},
    'cash_rollforward': {'cash_close': 1, 'cash_open': -1, 'ocf': -1, 'icf': -1, 'fcf_financing': -1, 'fx_cash': -1},
    'cash_scope': {'book_cash': 1, 'cfs_cash': -1, 'restricted_cash': -1, 'other_cash_scope_difference': -1},
    'retained_earnings': {'retained_close': 1, 'retained_open': -1, 'attributable_profit': -1, 'dividends': 1, 'reserve_transfer': 1, 'other_retained_changes': -1},
    'debt_rollforward': {'debt_close': 1, 'debt_open': -1, 'borrowings': -1, 'principal_paid': 1, 'noncash_debt_changes': -1},
    'ppe_rollforward': {'ppe_close': 1, 'ppe_open': -1, 'ppe_additions': -1, 'depreciation': 1, 'ppe_disposal_nbv': 1, 'impairment': 1, 'ppe_other_changes': -1},
}


def reconcile(name, data, rounding=None, comparable=True):
    terms = CHECKS[name]
    formula = ' + '.join(f'{a}*{k}' for k, a in terms.items()) + ' = 0'
    if not comparable:
        return result('NOT_COMPARABLE', formula=formula, reason='主体、期间或币种未统一')
    try:
        vals = {k: dec(data.get(k)) for k in terms}
        difference = sum((vals[k] * a for k, a in terms.items()), dec(0))
        if rounding is not None:
            tolerance = sum((abs(a) * dec(rounding[k]) / 2 for k, a in terms.items()), dec(0))
            if any(dec(rounding[k]) <= 0 for k in terms): raise ValueError('无效显示精度')
            status = 'ROUNDING_ONLY' if abs(difference) <= tolerance else 'FAIL_RECONCILIATION'
        else:
            tolerance = dec('0.01')
            status = 'PASS' if abs(difference) <= tolerance else 'FAIL_RECONCILIATION'
        return result(status, difference, formula, vals, tolerance=tolerance)
    except (ValueError, KeyError) as exc:
        return result('BLOCKED_INPUT', formula=formula, inputs=data, reason='缺失/无效输入或精度: ' + str(exc))


def bridge(base, items, target):
    """Explicit signed adjustments. Approved/supported only; no automatic accounting judgments."""
    total = dec(base)
    seen = set()
    included, excluded = [], []
    for item in items:
        key = item['adjustment_id']
        if key in seen:
            return result('BLOCKED_INPUT', reason='重复调整编号: ' + key)
        seen.add(key)
        if item.get('bucket') != target or item.get('status') != 'supported' or not item.get('source'):
            excluded.append(item)
            continue
        total += dec(item['amount'])
        included.append(item)
    return result('PASS', total, 'base + sum(supported signed adjustments in selected bucket)',
                  {'base': base, 'items': included}, excluded=excluded)


def cross_bucket_check(items):
    used = {}
    conflicts = []
    for item in items:
        # Same economic item in different price-adjustment buckets requires explicit resolution.
        key = item.get('economic_item_id', item['adjustment_id'])
        bucket = item['bucket']
        if key in used and used[key] != bucket:
            conflicts.append(key)
        used[key] = bucket
    return result('BLOCKED_INPUT' if conflicts else 'PASS', conflicts, reason='跨QoE/NWC/净债务重复影响需人工解释' if conflicts else '')


def variance(actual, budget):
    a, b = dec(actual), dec(budget)
    diff = a - b
    return dict(amount=diff, pct=diff / abs(b) if b != 0 else None,
                pct_status='PASS' if b != 0 else 'NOT_APPLICABLE',
                direction='neutral' if diff == 0 else ('above_budget' if diff > 0 else 'below_budget'),
                business_judgment='待结合业务进度解释；不是自动利好/利空')
