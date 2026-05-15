"""Aggregation bucket factories and serializers."""

from collections import Counter, defaultdict


def new_reg_metrics():
    return {
        'n_total': 0, 'n_base_all': 0, 'n_stayed_base': 0,
        'n_transferred': 0, 'n_sla_acc_viol': 0, 'n_sla_cont_viol': 0,
        'n_repeat': 0,
        'n_connected_stayed_from_period': 0,
        'nd_sum_from_period': 0.0,
        # Month-only conversion counters (reg_month -> final in reg_month or next_month)
        # Clean  = primary FIO
        # Regular = exec FIO (assigned-to)
        'n_month_conv_primary_total': 0,
        'n_month_conv_primary_connected': 0,
        'n_month_conv_exec_total': 0,
        'n_month_conv_exec_connected': 0,
        # Conversion (requested): by "exec_name" and by "primary_name"
        'n_conv_exec_total': 0,
        'n_conv_exec_connected': 0,
        'n_conv_primary_total': 0,
        'n_conv_primary_connected': 0,
        # Clean-conversion numerator (our channels; requires both primary_name and exec_name)
        'n_conv_clean_exec_connected': 0,
        'n_month_conv_clean_exec_connected': 0,
        # Clean-conversion denominator across ALL channels (registered totals by primary FIO)
        'n_conv_primary_total_all': 0,
        'n_month_conv_primary_total_all': 0,
    }


def new_closed_metrics():
    return {
        'n_closed': 0,
        'n_connected': 0,
        'n_with_equip': 0,
        'nd_sum': 0.0,
        'by_service': defaultdict(int),
        'by_service_equip': defaultdict(int),
        'inn_services': defaultdict(set),
        'secondaries': [],
    }


def new_open_rec():
    return {'n_open': 0, 'n_stale30': 0, 'by_week': Counter(), 'stale_by_week': Counter()}


def new_metric_scope(factory):
    return {
        'total': defaultdict(factory),
        'direction': defaultdict(factory),
        'director': defaultdict(factory),
        'teamlead': defaultdict(factory),
        'employee': defaultdict(factory),
        # Separate axis for exec_name ("на кого назначено") to avoid mixing with primary employee buckets.
        'employee_exec': defaultdict(factory),
    }


def serialize_per(bucket_dict, to_dict):
    out = {}
    for per, wd in bucket_dict.items():
        out[per] = {
            'total': {k: to_dict(v) for k, v in wd['total'].items()},
            'direction': {k: to_dict(v) for k, v in wd['direction'].items()},
            'director': {k: to_dict(v) for k, v in wd['director'].items()},
            'teamlead': {k: to_dict(v) for k, v in wd['teamlead'].items()},
            'employee': {k: to_dict(v) for k, v in wd['employee'].items()},
            'employee_exec': {k: to_dict(v) for k, v in wd['employee_exec'].items()},
        }
    return out
