"""Aggregation bucket factories and serializers."""

from collections import Counter, defaultdict


def new_reg_metrics():
    return {
        'n_total': 0, 'n_base_all': 0, 'n_stayed_base': 0,
        'n_transferred': 0, 'n_sla_acc_viol': 0, 'n_sla_cont_viol': 0,
        'n_connected_stayed_from_period': 0,
        'nd_sum_from_period': 0.0,
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
        }
    return out
