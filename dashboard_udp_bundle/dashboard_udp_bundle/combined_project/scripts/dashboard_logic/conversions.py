"""Хелперы расчёта конверсий.

Этот модуль централизует все расчёты конверсии, используемые в дашборде:
- REGISTERED: чистая/обычная конверсия + разрезы exec/primary
- CLOSED: конверсия закрытых (connected / closed)
"""


def pct(a, b):
    return round(100.0 * a / b, 2) if b else 0.0


def apply_conv_exec(m, connected):
    """Счётчики REGISTERED-конверсии в разрезе `exec_name`."""
    m['n_conv_exec_total'] += 1
    if connected:
        m['n_conv_exec_connected'] += 1


def apply_conv_primary(m, connected):
    """Счётчики REGISTERED-конверсии в разрезе `primary_name`."""
    m['n_conv_primary_total'] += 1
    if connected:
        m['n_conv_primary_connected'] += 1


def reg_conversion_fields(m):
    """Возвращает поля конверсии для словаря метрик REGISTERED."""
    n_exec_total = m.get('n_conv_exec_total', 0)
    n_exec_conn = m.get('n_conv_exec_connected', 0)
    n_prim_total = m.get('n_conv_primary_total', 0)
    n_prim_conn = m.get('n_conv_primary_connected', 0)
    n_clean_conn = m.get('n_conv_clean_connected', 0)

    # Rule:
    # - conv_clean   uses PRIMARY FIO axis
    # - conv_regular uses EXEC (assigned-to) FIO axis
    # Numerator/denominator differ ONLY by connected status (no extra filters).
    # For month slices we may have month-window counters (final_dt in reg_month or next_month).
    # Same logic for weeks and months; period aggregation is handled by the bucket selection.
    conv_clean = pct(n_clean_conn, n_prim_total)
    conv_regular = pct(n_exec_conn, n_exec_total)
    return {
        'conv_clean': conv_clean,
        'conv_regular': conv_regular,
        'n_conv_exec_total': n_exec_total,
        'n_conv_exec_connected': n_exec_conn,
        'conv_exec_pct': pct(n_exec_conn, n_exec_total),
        'n_conv_primary_total': n_prim_total,
        'n_conv_primary_connected': n_prim_conn,
        'conv_primary_pct': pct(n_prim_conn, n_prim_total),
    }


def closed_conversion_fields(m):
    """Возвращает поля конверсии для словаря метрик CLOSED."""
    n_cl = m['n_closed']
    n_conn = m['n_connected']
    return {
        'close_conv': pct(n_conn, n_cl),
    }
