"""Metric calculation and output conversion helpers."""

ULTRA_CHEAP_RUB = 400.0


def pct(a, b):
    return round(100.0 * a / b, 2) if b else 0.0


def apply_reg_metrics(m, in_base, stayed, transferred, connected, sla_acc_viol, sla_cont_viol, nd_val):
    m['n_total'] += 1
    if in_base:
        m['n_base_all'] += 1
        if stayed:
            m['n_stayed_base'] += 1
            if connected:
                m['n_connected_stayed_from_period'] += 1
                m['nd_sum_from_period'] += nd_val
    if transferred:
        m['n_transferred'] += 1
    if sla_acc_viol:
        m['n_sla_acc_viol'] += 1
    if sla_cont_viol:
        m['n_sla_cont_viol'] += 1


def apply_closed_metrics(m, connected, with_equip, nd_val, service_raw, cross_eligible, internet_service, inn_val, monthly_amount):
    m['n_closed'] += 1
    if connected:
        m['n_connected'] += 1
        m['by_service'][service_raw] += 1
        if with_equip:
            m['n_with_equip'] += 1
            m['by_service_equip'][service_raw] += 1
        m['nd_sum'] += nd_val
        if cross_eligible:
            m['inn_services'][inn_val].add(service_raw)
            if not internet_service:
                m['secondaries'].append((service_raw, monthly_amount, inn_val))


def apply_open_metrics(rec, wkey_reg, is_stale):
    rec['n_open'] += 1
    rec['by_week'][wkey_reg] += 1
    if is_stale:
        rec['n_stale30'] += 1
        rec['stale_by_week'][wkey_reg] += 1


def reg_dict(m):
    n_t = m['n_total']
    return {
        'n_total': n_t,
        'n_base_all': m['n_base_all'],
        'n_stayed_base': m['n_stayed_base'],
        'n_transferred': m['n_transferred'],
        'n_sla_acc_viol': m['n_sla_acc_viol'],
        'n_sla_cont_viol': m['n_sla_cont_viol'],
        'transfer_rate': pct(m['n_transferred'], n_t),
        'sla_acc_rate': pct(m['n_sla_acc_viol'], n_t),
        'sla_cont_rate': pct(m['n_sla_cont_viol'], n_t),
        'n_connected_stayed_from_period': m['n_connected_stayed_from_period'],
        'conv_clean': pct(m['n_connected_stayed_from_period'], m['n_base_all']),
        'conv_regular': pct(m['n_connected_stayed_from_period'], m['n_stayed_base']),
        'nd_sum_from_period': round(m['nd_sum_from_period'], 2),
    }


def closed_dict(m):
    n_cl = m['n_closed']
    n_conn = m['n_connected']
    n_inn_connected = len(m['inn_services'])
    n_inn_multi = sum(1 for s in m['inn_services'].values() if len(s) >= 2)
    secondaries = m['secondaries']
    n_secondary = len(secondaries)
    n_ultra_cheap = sum(1 for _, price, _ in secondaries if price <= ULTRA_CHEAP_RUB)
    return {
        'n_closed': n_cl,
        'n_connected': n_conn,
        'n_with_equip': m['n_with_equip'],
        'nd_sum': round(m['nd_sum'], 2),
        'close_conv': pct(n_conn, n_cl),
        'equip_share': pct(m['n_with_equip'], n_conn) if n_conn else 0.0,
        'by_service': dict(m['by_service']),
        'by_service_equip': dict(m['by_service_equip']),
        'n_inn_connected': n_inn_connected,
        'n_inn_multi': n_inn_multi,
        'crossell_pct': pct(n_inn_multi, n_inn_connected),
        'n_secondary': n_secondary,
        'n_ultra_cheap': n_ultra_cheap,
        'fraud_pct': pct(n_ultra_cheap, n_secondary),
    }


def open_to_dict(r):
    return {
        'n_open': r['n_open'],
        'n_stale30': r['n_stale30'],
        'stale_rate': pct(r['n_stale30'], r['n_open']),
        'by_week': dict(r['by_week']),
        'stale_by_week': dict(r['stale_by_week']),
    }
