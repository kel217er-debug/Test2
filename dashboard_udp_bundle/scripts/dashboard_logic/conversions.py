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

    # New rule: conversions differ only by FIO source:
    # - clean   = primary FIO
    # - regular = exec FIO (assigned-to)
    #
    # For month slices we may have month-window counters (final_dt in reg_month or next_month).
    n_m_prim_total = m.get('n_month_conv_primary_total', 0)
    n_m_exec_total = m.get('n_month_conv_exec_total', 0)
    if n_m_prim_total or n_m_exec_total:
        # Variant 2: numerator by exec, denominator by primary.
        conv_clean = pct(m.get('n_month_conv_exec_connected', 0), n_m_prim_total)
        conv_regular = pct(m.get('n_month_conv_exec_connected', 0), n_m_exec_total)
    else:
        # Variant 2: numerator by exec, denominator by primary.
        conv_clean = pct(n_exec_conn, n_prim_total)
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
