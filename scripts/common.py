"""Local-only shared primitives. No network, shell, eval, dependency installation."""
from __future__ import annotations
import calendar
import hashlib
import json
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
QA = Path('/private/tmp/pe-vc-financial-dd-v0.2-qa')
VERSION = '0.2.0'


def dec(value):
    if value is None or isinstance(value, bool):
        raise ValueError('Missing or invalid numeric input')
    text = str(value).strip().replace(',', '').replace('，', '').replace('−', '-')
    if text in ('', '-', '—', 'N/A', 'n.a.', 'None'):
        raise ValueError('Missing numeric input')
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    try:
        n = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError('Invalid numeric input') from exc
    if not n.is_finite():
        raise ValueError('Non-finite numeric input')
    return n


def serial(value):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, Path)):
        return str(value)
    if isinstance(value, dict):
        return {str(k): serial(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [serial(v) for v in value]
    return value


def fingerprint(value):
    raw = json.dumps(serial(value), ensure_ascii=False, sort_keys=True, separators=(',', ':'))
    return hashlib.sha256(raw.encode()).hexdigest()


def file_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def code_hash():
    return fingerprint({p.name: file_hash(p) for p in sorted((ROOT / 'scripts').glob('*.py'))})


def allowed_write(path):
    p = Path(path).resolve()
    if not any(p.is_relative_to(base.resolve()) for base in (ROOT, QA)):
        raise PermissionError('Write outside approved prototype/QA scope')
    return p


def write_json(path, payload):
    p = allowed_write(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(serial(payload), ensure_ascii=False, indent=2), encoding='utf-8')


def write_text(path, text):
    p = allowed_write(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding='utf-8')


def result(status, value=None, formula='', inputs=None, reason='', **extra):
    return dict(status=status, value=value, formula=formula, inputs=inputs or {}, reason=reason, **extra)


def add_months(day, months):
    d = date.fromisoformat(day) if isinstance(day, str) else day
    m = d.year * 12 + d.month - 1 + months
    y, mi = divmod(m, 12)
    # Preserve explicitly month-end-anchored dates when shifting monthly budgets.
    target_last = calendar.monthrange(y, mi + 1)[1]
    was_month_end = d.day == calendar.monthrange(d.year, d.month)[1]
    return date(y, mi + 1, target_last if was_month_end else min(d.day, target_last))


def month_ends(base, months):
    d = date.fromisoformat(base)
    for i in range(1, months + 1):
        x = add_months(d, i)
        yield date(x.year, x.month, calendar.monthrange(x.year, x.month)[1])


def require_source(item):
    if not item.get('source'):
        raise ValueError('Missing evidence/source')


def safe_text(value):
    # Untrusted text is never exported as an active spreadsheet formula.
    if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@')):
        return "'" + value
    return value
