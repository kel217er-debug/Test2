#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETL v2 для ежедневного дашборда тимлидов — проект "Заявки+обращения".

Это отдельный ETL, чтобы не менять основной "МУЗ"-проект:
- вход: `new/Заявки+обращения_объединено.xlsx`
- выход: `data_combined/processed/daily_dashboard_data.json`
"""

import os, sys, re, json, time
from datetime import datetime
from collections import defaultdict
from pathlib import Path

import pandas as pd

from dashboard_logic.aggregations import (
    new_closed_metrics,
    new_metric_scope,
    new_open_rec,
    new_reg_metrics,
    serialize_per,
)
from dashboard_logic.excel_filters_combined import (
    COL,
    direction_for_channel,
    is_base_service,
    is_connected,
    is_internet_service,
    is_our_channel,
    is_transferred,
)
from dashboard_logic.metrics import (
    apply_closed_metrics,
    apply_open_metrics,
    apply_reg_metrics,
    closed_dict,
    open_to_dict,
    reg_dict,
)
from dashboard_logic.conversions import apply_conv_exec, apply_conv_primary
from dashboard_logic.periods import (
    build_months_info,
    build_timeline,
    build_weeks_info,
    iso_week_key,
    month_bounds,
    month_key,
    split_timeline_periods,
)
from dashboard_logic.sla_rules import is_sla_violated
from dashboard_logic.repeat_metrics import is_repeat_by_comment, normalize_phone

from calc_combined_mapping_fields import _read_mapping, build_fields  # type: ignore

STALE_HOURS = 720  # 30 дней

# -------- helpers --------
_DT_RE = re.compile(r'^\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})\s*$')

def parse_dt(v):
    if v is None: return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, datetime): return v
    if hasattr(v, 'to_pydatetime'):
        try: return v.to_pydatetime()
        except Exception: pass
    if isinstance(v, str):
        s = v.strip()
        if not s: return None
        m = _DT_RE.match(s)
        if m:
            d, mo, y, H, M, S = map(int, m.groups())
            try: return datetime(y, mo, d, H, M, S)
            except ValueError: return None
    return None


def next_month_key(mkey: str) -> str:
    """Return YYYY-MM for the month after `mkey` (YYYY-MM)."""
    try:
        y_s, m_s = mkey.split('-', 1)
        y = int(y_s)
        mo = int(m_s)
    except Exception:
        return mkey
    if mo >= 12:
        return f"{y + 1}-01"
    return f"{y}-{mo + 1:02d}"


def month_connected_in_window(reg_mkey: str, final_dt: datetime) -> bool:
    """True if final_dt month is reg_mkey or the next month after reg_mkey."""
    mkey_final = month_key(final_dt)
    if mkey_final == reg_mkey:
        return True
    return mkey_final == next_month_key(reg_mkey)

def parse_num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace('\xa0','').replace(' ','').replace(',','.')
    if not s: return 0.0
    try: return float(s)
    except: return 0.0

def norm_str(v):
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except Exception:
        pass
    s = str(v).strip()
    if not s:
        return ''
    if s.lower() in ('nan', '<na>', 'none'):
        return ''
    return s

def clean_name(v):
    s = norm_str(v)
    if not s:
        return ''

    # Some exports may append emails after FIO.
    if '@' in s:
        parts = s.split()
        clean = []
        for p in parts:
            if '@' in p:
                break
            clean.append(p)
        s = ' '.join(clean)

    tokens = re.findall(r"[A-Za-zА-Яа-яЁёЀ-ӿ-]+", s)

    # Prefer Cyrillic-only sequence when present.
    cyr = [t for t in tokens if re.search(r"[А-Яа-яЁёЀ-ӿ]", t)]
    if len(cyr) >= 2:
        return ' '.join(cyr[:3]).strip()

    return ' '.join(tokens).strip()


# -------- иерархия (копия из базового ETL) --------

def _truthy(v):
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    s = str(v).strip().lower()
    if not s:
        return False
    return s in ('1', 'true', 'yes', 'y', 'да', 'д', 'активен', 'active')


def _load_hierarchy_from_json(hierarchy_path):
    with open(hierarchy_path, 'r', encoding='utf-8') as f:
        raw = json.load(f)

    employees = raw.get('employees', raw if isinstance(raw, list) else [])
    out = {}
    for e in employees:
        name = clean_name(e.get('name', ''))
        if not name:
            continue
        out[name] = {
            'teamlead': clean_name(e.get('teamlead', '')),
            'director': clean_name(e.get('director', '')),
            'direction': norm_str(e.get('direction', '')),
            'mrf': norm_str(e.get('mrf', '')),
            'is_active': _truthy(e.get('is_active', True)),
        }
    return out


def _norm_header(v):
    if v is None:
        return ''
    s = str(v).strip().lower()
    s = re.sub(r'\s+', ' ', s)
    s = s.replace('ё', 'е')
    return s


def _load_hierarchy_from_xlsx(hierarchy_path):
    df = pd.read_excel(hierarchy_path, dtype=object)
    if df.empty:
        return {}

    headers = [_norm_header(v) for v in df.columns]

    def col_index(*needles):
        for name in needles:
            name_n = _norm_header(name)
            for i, h in enumerate(headers):
                if h == name_n:
                    return i
        return None

    idx_name = col_index('name', 'employee', 'fio', 'фио', 'сотрудник', 'фамилия имя отчество')
    idx_teamlead = col_index('teamlead', 'tl', 'тимлид', 'тим лид', 'тим-лид', 'тимлидер', 'тим лидер')
    idx_director = col_index('director', 'manager', 'руководитель', 'директор')
    idx_direction = col_index('direction', 'направление')
    idx_mrf = col_index('mrf', 'мрф')
    idx_is_active = col_index('is_active', 'active', 'активен', 'активность', 'действующий', 'действующий да/нет', 'действующий да / нет')

    if idx_name is None:
        raise ValueError('Hierarchy Excel: required column "name" (or "ФИО") not found in header row')

    out = {}
    for _, rr in df.iterrows():
        row = list(rr.values.tolist())
        name = clean_name(row[idx_name] if idx_name is not None and idx_name < len(row) else '')
        if not name:
            continue
        out[name] = {
            'teamlead': clean_name(row[idx_teamlead]) if idx_teamlead is not None and idx_teamlead < len(row) else '',
            'director': clean_name(row[idx_director]) if idx_director is not None and idx_director < len(row) else '',
            'direction': norm_str(row[idx_direction]) if idx_direction is not None and idx_direction < len(row) else '',
            'mrf': norm_str(row[idx_mrf]) if idx_mrf is not None and idx_mrf < len(row) else '',
            'is_active': _truthy(row[idx_is_active]) if idx_is_active is not None and idx_is_active < len(row) else True,
        }
    return out


def load_hierarchy(hierarchy_path):
    if hierarchy_path.lower().endswith('.xlsx'):
        return _load_hierarchy_from_xlsx(hierarchy_path)
    return _load_hierarchy_from_json(hierarchy_path)


def _load_filter_people_config(cfg_path, all_teamleads, all_directors):
    cfg = {'teamleads': {}, 'directors': {}}
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                if isinstance(raw.get('teamleads'), dict):
                    cfg['teamleads'] = {clean_name(k): bool(v) for k, v in raw['teamleads'].items() if clean_name(k)}
                if isinstance(raw.get('directors'), dict):
                    cfg['directors'] = {clean_name(k): bool(v) for k, v in raw['directors'].items() if clean_name(k)}
        except Exception as e:
            print(f"[ETL COMBINED] WARN: failed to read filter config {cfg_path}: {e}", file=sys.stderr)

    changed = False
    for n in sorted(all_teamleads):
        if n not in cfg['teamleads']:
            cfg['teamleads'][n] = True
            changed = True
    for n in sorted(all_directors):
        if n not in cfg['directors']:
            cfg['directors'][n] = True
            changed = True

    cfg['teamleads'] = {k: cfg['teamleads'][k] for k in sorted(cfg['teamleads'])}
    cfg['directors'] = {k: cfg['directors'][k] for k in sorted(cfg['directors'])}

    if changed:
        os.makedirs(os.path.dirname(cfg_path), exist_ok=True)
        with open(cfg_path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print(f"[ETL COMBINED] Filter config updated: {cfg_path}", flush=True)

    return cfg


def _build_rows_from_combined(input_xlsx: str, mapping_txt: str):
    df_in = pd.read_excel(input_xlsx, sheet_name=0, dtype=object)
    mapping_rows = _read_mapping(Path(mapping_txt))
    fields = build_fields(df_in, mapping_rows, now_dt=datetime.now())

    # For repeat metric: take the first available phone/comment columns from the combined export.
    phone_candidates = (
        'Номер телефона',
        'Номер телефона для связи',
        'Обращение.Номер телефона',
        'Обращение.Номер телефона для связи',
    )
    comment_candidates = (
        'Комментарий к статусу',
        'Обращение.Комментарий к статусу',
        'Обращение.Комментарий регистратора',
    )
    phone_series = None
    for c in phone_candidates:
        if c in df_in.columns:
            phone_series = df_in[c]
            break
    comment_series = None
    for c in comment_candidates:
        if c in df_in.columns:
            comment_series = df_in[c]
            break

    max_idx = max(COL.values())
    header_row = [''] * (max_idx + 1)
    for k, idx in COL.items():
        header_row[idx] = k

    rows = []
    phones = []
    comments = []
    for _, r in fields.iterrows():
        row = [None] * (max_idx + 1)
        for k, idx in COL.items():
            if k in fields.columns:
                row[idx] = r.get(k)
        rows.append(row)
        i = len(phones)
        phones.append(phone_series.iloc[i] if phone_series is not None else None)
        comments.append(comment_series.iloc[i] if comment_series is not None else None)

    return header_row, iter(rows), phones, comments


# -------- основной проход (адаптация базового ETL) --------

def _mrf_selected(mrf, include_mrfs=None, exclude_mrfs=None):
    if include_mrfs is not None and mrf not in include_mrfs:
        return False
    if exclude_mrfs is not None and mrf in exclude_mrfs:
        return False
    return True


def _env_list(name):
    raw = os.environ.get(name, '')
    vals = [v.strip() for v in raw.split('|') if v.strip()]
    return set(vals) if vals else None


def run_etl_combined(input_xlsx: str, mapping_txt: str, hierarchy_path: str, out_path: str, include_mrfs=None, exclude_mrfs=None):
    t0 = time.time()
    print(f"[ETL COMBINED] Start. input={input_xlsx}")

    emp_map = load_hierarchy(hierarchy_path)
    print(f"[ETL COMBINED] Hierarchy: {len(emp_map)} employees from {hierarchy_path}")

    header_row, rows_iter, repeat_phones, repeat_comments = _build_rows_from_combined(input_xlsx, mapping_txt)

    # repeat columns are not available in synthetic export
    idx_repeat_comment = None
    idx_phone = None

    header_len = len(header_row)
    raw_export = {'headers': list(header_row), 'rows': []}

    weekly_reg = defaultdict(lambda: new_metric_scope(new_reg_metrics))
    weekly_closed = defaultdict(lambda: new_metric_scope(new_closed_metrics))
    monthly_reg = defaultdict(lambda: new_metric_scope(new_reg_metrics))
    monthly_closed = defaultdict(lambda: new_metric_scope(new_closed_metrics))
    cohort = defaultdict(lambda: defaultdict(lambda: {'n_connected': 0, 'nd_sum': 0.0, 'n_with_equip': 0}))

    open_snap = {
        'total': new_open_rec(),
        'direction': defaultdict(new_open_rec),
        'director': defaultdict(new_open_rec),
        'teamlead': defaultdict(new_open_rec),
        'employee': defaultdict(new_open_rec),
    }
    daily_reg = defaultdict(lambda: {'n_total': 0})
    daily_closed = defaultdict(lambda: {'n_connected': 0, 'nd_sum': 0.0})

    weeks_meta = {}
    months_meta = {}

    n_read = 0; n_selected = 0; n_our = 0; n_skip_no_date = 0
    dates_min = None; dates_max = None
    close_min = None; close_max = None

    phone_min_info = {}
    repeat_rows = []

    def _apply_clean_same_emp_connected_weekly(m, same_emp: bool, connected: bool):
        # Clean conversion numerator: rows where primary_name == exec_name AND Status == "Услуга подключена".
        if same_emp and connected:
            m['n_conv_clean_connected'] += 1

    def _apply_clean_same_emp_connected_monthly(m, same_emp: bool, connected: bool):
        # Month-only numerator for clean conversion (same condition as weekly).
        if same_emp and connected:
            # Keep non-month counters in monthly buckets too, so month conversion can reuse week logic.
            m['n_conv_clean_connected'] += 1
            m['n_month_conv_clean_connected'] += 1

    t_iter = time.time()
    for row in rows_iter:
        n_read += 1
        if n_read % 20000 == 0:
            print(f"  ... {n_read} rows, {time.time()-t_iter:.1f}s")

        podr_prim = norm_str(row[COL['primary_podr']])
        reg_dt = parse_dt(row[COL['reg_dt']])
        if reg_dt is None:
            n_skip_no_date += 1
            continue

        # ------ Метрика повторов (комментарий + телефон) ------
        # Повтор определяем по ключевым словам в комментарии, группируем по телефону,
        # атрибуция по самой ранней reg_dt на телефон.
        try:
            phone_digits = normalize_phone(repeat_phones[n_read - 1])
        except Exception:
            phone_digits = ''
        comment_val = None
        try:
            comment_val = repeat_comments[n_read - 1]
        except Exception:
            comment_val = None
        direction = direction_for_channel(podr_prim)
        service = norm_str(row[COL['service']]).lower()
        service_raw = norm_str(row[COL['service']]) or '—'
        current_podr = norm_str(row[COL['current_podr']])
        _stayed = is_our_channel(current_podr)
        _transferred = is_transferred(current_podr)
        final_status = norm_str(row[COL['final_status']])
        _connected = is_connected(final_status)
        _has_final = bool(final_status)
        final_dt = parse_dt(row[COL['final_dt']]) if _has_final else None
        wkey_cl = None
        mkey_cl = None
        if _has_final and final_dt is not None:
            wkey_cl, _, _ = iso_week_key(final_dt)
            mkey_cl = month_key(final_dt)
        _in_conv_base = is_base_service(service)
        sla_acc = norm_str(row[COL['sla_acc']]).lower()
        sla_cont = norm_str(row[COL['sla_cont']]).lower()
        _sla_acc_viol = is_sla_violated(sla_acc)
        _sla_cont_viol = is_sla_violated(sla_cont)
        equip_price = parse_num(row[COL['equip_price']])
        _with_equip = _connected and (equip_price > 0)

        ND_OUTLIER_CAP = 1_000_000_000.0
        inst = parse_num(row[COL['install_amount']])
        mon_ = parse_num(row[COL['monthly_amount']])
        nd_val = inst + mon_
        if nd_val > ND_OUTLIER_CAP:
            nd_val = 0.0

        inn_val = norm_str(row[COL['inn']])
        _cross_eligible = bool(inn_val) and _connected
        _is_internet = is_internet_service(service)

        primary_name = clean_name(row[COL['primary_name']])
        exec_name = clean_name(row[COL['exec_name']])
        same_emp = bool(primary_name) and (primary_name == exec_name)

        wkey_reg, mon_r, sun_r = iso_week_key(reg_dt)
        mkey_reg = month_key(reg_dt)
        if wkey_reg not in weeks_meta: weeks_meta[wkey_reg] = (mon_r, sun_r)
        if mkey_reg not in months_meta: months_meta[mkey_reg] = month_bounds(mkey_reg)

        # Primary hierarchy
        info = emp_map.get(primary_name) if primary_name else None
        director = info.get('director', '') if info else ''
        teamlead = info.get('teamlead', '') if info else ''
        mrf = info.get('mrf', '') if info else ''

        if not _mrf_selected(mrf, include_mrfs, exclude_mrfs):
            continue
        n_selected += 1

        is_our = is_our_channel(podr_prim)
        if is_our:
            n_our += 1

        if dates_min is None or reg_dt < dates_min: dates_min = reg_dt
        if dates_max is None or reg_dt > dates_max: dates_max = reg_dt

        if phone_digits:
            prev = phone_min_info.get(phone_digits)
            if (prev is None) or (reg_dt < prev['dt']):
                phone_min_info[phone_digits] = {'dt': reg_dt, 'director': '', 'teamlead': '', 'employee': ''}
            if is_repeat_by_comment(comment_val):
                repeat_rows.append(phone_digits)

        # Fill repeat attribution dims (minimum reg_dt per phone)
        if phone_digits:
            prev = phone_min_info.get(phone_digits)
            if prev and prev.get('dt') == reg_dt:
                prev['director'] = director
                prev['teamlead'] = teamlead
                prev['employee'] = primary_name

        # Exec hierarchy (separate axis)
        info_exec = emp_map.get(exec_name) if exec_name else None
        exec_director = info_exec.get('director', '') if info_exec else ''
        exec_teamlead = info_exec.get('teamlead', '') if info_exec else ''
        exec_mrf = info_exec.get('mrf', '') if info_exec else ''

        # Direction for exec is not defined for combined (leave empty)
        exec_direction = ''

        # ------ REGISTERED агрегация (по дате регистрации) ------
        def apply_reg_weekly(m):
            apply_reg_metrics(m, _in_conv_base, _stayed, _transferred, _connected, _sla_acc_viol, _sla_cont_viol, nd_val)

        def apply_reg_monthly(m):
            apply_reg_metrics(m, _in_conv_base, _stayed, _transferred, _connected, _sla_acc_viol, _sla_cont_viol, nd_val)
            # Month-only conversion counters:
            # numerator differs from denominator ONLY by connected status.
            m['n_month_conv_primary_total'] += 1
            m['n_month_conv_exec_total'] += 1
            if _connected:
                m['n_month_conv_primary_connected'] += 1
                m['n_month_conv_exec_connected'] += 1

        # total/all
        _apply_clean_same_emp_connected_weekly(weekly_reg[wkey_reg]['total']['all'], same_emp, _connected)
        _apply_clean_same_emp_connected_monthly(monthly_reg[mkey_reg]['total']['all'], same_emp, _connected)

        apply_reg_weekly(weekly_reg[wkey_reg]['total']['all'])
        apply_reg_monthly(monthly_reg[mkey_reg]['total']['all'])
        apply_conv_exec(weekly_reg[wkey_reg]['total']['all'], _connected)
        apply_conv_primary(weekly_reg[wkey_reg]['total']['all'], _connected)

        apply_conv_exec(monthly_reg[mkey_reg]['total']['all'], _connected)
        apply_conv_primary(monthly_reg[mkey_reg]['total']['all'], _connected)

        # buckets: director/teamlead/employee/mrf
        if director:
            _apply_clean_same_emp_connected_weekly(weekly_reg[wkey_reg]['director'][director], same_emp, _connected)
            _apply_clean_same_emp_connected_monthly(monthly_reg[mkey_reg]['director'][director], same_emp, _connected)
            apply_reg_weekly(weekly_reg[wkey_reg]['director'][director])
            apply_reg_monthly(monthly_reg[mkey_reg]['director'][director])
            apply_conv_primary(weekly_reg[wkey_reg]['director'][director], _connected)
            apply_conv_primary(monthly_reg[mkey_reg]['director'][director], _connected)
        if teamlead:
            _apply_clean_same_emp_connected_weekly(weekly_reg[wkey_reg]['teamlead'][teamlead], same_emp, _connected)
            _apply_clean_same_emp_connected_monthly(monthly_reg[mkey_reg]['teamlead'][teamlead], same_emp, _connected)
            apply_reg_weekly(weekly_reg[wkey_reg]['teamlead'][teamlead])
            apply_reg_monthly(monthly_reg[mkey_reg]['teamlead'][teamlead])
            apply_conv_primary(weekly_reg[wkey_reg]['teamlead'][teamlead], _connected)
            apply_conv_primary(monthly_reg[mkey_reg]['teamlead'][teamlead], _connected)
        if primary_name:
            _apply_clean_same_emp_connected_weekly(weekly_reg[wkey_reg]['employee'][primary_name], same_emp, _connected)
            _apply_clean_same_emp_connected_monthly(monthly_reg[mkey_reg]['employee'][primary_name], same_emp, _connected)
            apply_reg_weekly(weekly_reg[wkey_reg]['employee'][primary_name])
            apply_reg_monthly(monthly_reg[mkey_reg]['employee'][primary_name])
            apply_conv_primary(weekly_reg[wkey_reg]['employee'][primary_name], _connected)
            apply_conv_primary(monthly_reg[mkey_reg]['employee'][primary_name], _connected)

        # Exec axis (for conv_exec)
        if exec_director:
            apply_conv_exec(weekly_reg[wkey_reg]['director'][exec_director], _connected)
            apply_conv_exec(monthly_reg[mkey_reg]['director'][exec_director], _connected)
        if exec_teamlead:
            apply_conv_exec(weekly_reg[wkey_reg]['teamlead'][exec_teamlead], _connected)
            apply_conv_exec(monthly_reg[mkey_reg]['teamlead'][exec_teamlead], _connected)
        if exec_name:
            apply_conv_exec(weekly_reg[wkey_reg]['employee_exec'][exec_name], _connected)
            apply_conv_exec(monthly_reg[mkey_reg]['employee_exec'][exec_name], _connected)
            # Month-window conversion for exec employee bucket (needed for month conv_* in UI).
            monthly_exec_bucket = monthly_reg[mkey_reg]['employee_exec'][exec_name]
            monthly_exec_bucket['n_month_conv_exec_total'] += 1
            if _connected:
                monthly_exec_bucket['n_month_conv_exec_connected'] += 1
            if same_emp:
                if _connected:
                    weekly_reg[wkey_reg]['employee_exec'][exec_name]['n_conv_clean_connected'] += 1
                    monthly_reg[mkey_reg]['employee_exec'][exec_name]['n_conv_clean_connected'] += 1
                    monthly_exec_bucket['n_month_conv_clean_connected'] += 1

        daily_reg[reg_dt.strftime('%Y-%m-%d')]['n_total'] += 1

        # ------ CLOSED агрегация (по дате итогового статуса) ------
        if _has_final:
            if final_dt is not None:
                if close_min is None or final_dt < close_min: close_min = final_dt
                if close_max is None or final_dt > close_max: close_max = final_dt
                wkey_cl, mon_c, sun_c = iso_week_key(final_dt)
                mkey_cl = month_key(final_dt)
                if wkey_cl not in weeks_meta: weeks_meta[wkey_cl] = (mon_c, sun_c)
                if mkey_cl not in months_meta: months_meta[mkey_cl] = month_bounds(mkey_cl)

                def apply_closed(m):
                    apply_closed_metrics(
                        m,
                        _connected,
                        _with_equip,
                        nd_val,
                        service_raw,
                        _cross_eligible,
                        _is_internet,
                        inn_val,
                        mon_,
                    )

                apply_closed(weekly_closed[wkey_cl]['total']['all'])
                apply_closed(monthly_closed[mkey_cl]['total']['all'])
                if director:
                    apply_closed(weekly_closed[wkey_cl]['director'][director])
                    apply_closed(monthly_closed[mkey_cl]['director'][director])
                if teamlead:
                    apply_closed(weekly_closed[wkey_cl]['teamlead'][teamlead])
                    apply_closed(monthly_closed[mkey_cl]['teamlead'][teamlead])
                if primary_name:
                    apply_closed(weekly_closed[wkey_cl]['employee'][primary_name])
                    apply_closed(monthly_closed[mkey_cl]['employee'][primary_name])

                if _connected:
                    c = cohort[mkey_cl][mkey_reg]
                    c['n_connected'] += 1
                    c['nd_sum'] += nd_val
                    if _with_equip: c['n_with_equip'] += 1

                if _connected:
                    d = daily_closed[final_dt.strftime('%Y-%m-%d')]
                    d['n_connected'] += 1
                    d['nd_sum'] += nd_val

        # ------ OPEN snapshot ------
        if not _has_final:
            hrs_in_status = parse_num(row[COL['hrs_in_status']])
            is_stale = hrs_in_status > STALE_HOURS
            def apply_open(rec):
                apply_open_metrics(rec, wkey_reg, is_stale)
            apply_open(open_snap['total'])
            if director: apply_open(open_snap['director'][director])
            if teamlead: apply_open(open_snap['teamlead'][teamlead])
            if primary_name: apply_open(open_snap['employee'][primary_name])

        # ------ RAW rows export (for UI) ------
        try:
            def _json_cell(v):
                if v is None:
                    return None
                try:
                    if pd.isna(v):
                        return None
                except Exception:
                    pass
                if isinstance(v, datetime):
                    return v.strftime('%Y-%m-%d %H:%M:%S')
                if hasattr(v, 'to_pydatetime'):
                    try:
                        vv = v.to_pydatetime()
                        if isinstance(vv, datetime):
                            return vv.strftime('%Y-%m-%d %H:%M:%S')
                    except Exception:
                        pass
                return v

            raw_export['rows'].append({
                'v': [_json_cell(v) for v in list(row[:header_len])],
                'reg_w': wkey_reg,
                'reg_m': mkey_reg,
                'cl_w': wkey_cl,
                'cl_m': mkey_cl,
                'direction': direction,
                'mrf': mrf,
                'director': director,
                'teamlead': teamlead,
                'employee': primary_name,
                'exec_direction': exec_direction,
                'exec_director': exec_director,
                'exec_teamlead': exec_teamlead,
                'exec_mrf': exec_mrf,
                'employee_exec': exec_name,
                'in_base': bool(_in_conv_base),
                'stayed': bool(_stayed),
                'transferred': bool(_transferred),
                'connected': bool(_connected),
                'has_final': bool(wkey_cl and mkey_cl),
                'with_equip': bool(_with_equip),
                'sla_acc_viol': bool(_sla_acc_viol),
                'sla_cont_viol': bool(_sla_cont_viol),
                'is_stale30': bool((not _has_final) and (parse_num(row[COL['hrs_in_status']]) > STALE_HOURS)),
            })
        except Exception:
            pass

    print(f"[ETL COMBINED] Read {n_read}, our={n_our}, skip_no_date={n_skip_no_date}, {time.time()-t0:.1f}s")
    print(f"[ETL COMBINED] Reg dates: {dates_min} — {dates_max}")
    print(f"[ETL COMBINED] Close dates: {close_min} — {close_max}")

    # -------- Аккумуляция repeat по минимальной дате телефона --------
    repeat_attributed = 0
    repeat_total = len(repeat_rows)
    repeat_attribution_missing_min_dt = 0

    def _bump_repeat(m):
        m['n_repeat'] = (m.get('n_repeat') or 0) + 1

    for phone_digits in repeat_rows:
        info = phone_min_info.get(phone_digits)
        if not info:
            repeat_attribution_missing_min_dt += 1
            continue
        dt0 = info.get('dt')
        if not isinstance(dt0, datetime):
            repeat_attribution_missing_min_dt += 1
            continue
        wkey, _, _ = iso_week_key(dt0)
        mkey = month_key(dt0)
        _bump_repeat(weekly_reg[wkey]['total']['all'])
        _bump_repeat(monthly_reg[mkey]['total']['all'])
        if info.get('director'):
            _bump_repeat(weekly_reg[wkey]['director'][info['director']])
            _bump_repeat(monthly_reg[mkey]['director'][info['director']])
        if info.get('teamlead'):
            _bump_repeat(weekly_reg[wkey]['teamlead'][info['teamlead']])
            _bump_repeat(monthly_reg[mkey]['teamlead'][info['teamlead']])
        if info.get('employee'):
            _bump_repeat(weekly_reg[wkey]['employee'][info['employee']])
            _bump_repeat(monthly_reg[mkey]['employee'][info['employee']])
        repeat_attributed += 1

    if dates_max is None:
        raise RuntimeError("No rows with reg_dt found.")

    weeks_sorted = sorted(weeks_meta.keys())
    months_sorted = sorted(months_meta.keys())

    current_week, current_month, closed_months, open_month, open_weeks = split_timeline_periods(
        dates_max, weeks_sorted, weeks_meta, months_sorted, months_meta
    )
    weeks_info = build_weeks_info(weeks_sorted, weeks_meta)
    months_info = build_months_info(months_sorted, months_meta)

    teamleads_all = sorted({v.get('teamlead','') for v in emp_map.values() if v.get('teamlead')})
    directors_all = sorted({v.get('director','') for v in emp_map.values() if v.get('director')})

    filter_cfg_path = os.path.join(os.path.dirname(hierarchy_path), 'filter_people.json')
    filter_cfg = _load_filter_people_config(filter_cfg_path, teamleads_all, directors_all)
    teamleads = sorted([n for n in teamleads_all if filter_cfg['teamleads'].get(n, True)])
    directors = sorted([n for n in directors_all if filter_cfg['directors'].get(n, True)])

    timeline = build_timeline(closed_months, open_month, open_weeks, months_meta, weeks_meta, current_week)

    # Employees list for selectors
    seen_names = set()
    for w in weeks_sorted:
        seen_names.update(weekly_reg[w]['employee'].keys())
    emp_list = []
    for name, info in emp_map.items():
        emp_list.append({
            'name': name,
            'teamlead': info['teamlead'],
            'director': info['director'],
            'direction': info.get('direction',''),
            'mrf': info.get('mrf',''),
            'is_active': info['is_active'],
            'in_data': name in seen_names,
        })
    for name in seen_names:
        if name and name not in emp_map:
            emp_list.append({
                'name': name, 'teamlead': '', 'director': '',
                'direction': '', 'mrf': '', 'is_active': True, 'in_data': True,
            })

    directions = sorted({e['direction'] for e in emp_list if e.get('direction')})
    mrfs = sorted({e['mrf'] for e in emp_list if e.get('mrf')})

    registered = {
        'by_week': serialize_per(weekly_reg, reg_dict),
        'by_month': serialize_per(monthly_reg, reg_dict),
    }
    closed = {
        'by_week': serialize_per(weekly_closed, closed_dict),
        'by_month': serialize_per(monthly_closed, closed_dict),
    }

    cohort_out = {
        'by_month': {cm: {rm: dict(v) for rm, v in rs.items()} for cm, rs in cohort.items()}
    }
    for cm, rs in cohort_out['by_month'].items():
        for rm, v in rs.items():
            v['nd_sum'] = round(v.get('nd_sum', 0.0), 2)

    all_days = sorted(set(daily_reg.keys()) | set(daily_closed.keys()))
    daily_out = {}
    for d in all_days:
        r = daily_reg.get(d, {})
        c = daily_closed.get(d, {})
        daily_out[d] = {
            'n_total': r.get('n_total', 0),
            'n_connected': c.get('n_connected', 0),
            'nd_sum': round(c.get('nd_sum', 0.0), 2),
        }

    out = {
        'meta': {
            'generated_at': datetime.now().isoformat(timespec='seconds'),
            'source_kind': 'combined_requests_incidents',
            'source_file': os.path.relpath(input_xlsx, os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            'data_from': dates_min.strftime('%Y-%m-%d %H:%M:%S') if dates_min else None,
            'data_to': dates_max.strftime('%Y-%m-%d %H:%M:%S') if dates_max else None,
            'close_from': close_min.strftime('%Y-%m-%d %H:%M:%S') if close_min else None,
            'close_to': close_max.strftime('%Y-%m-%d %H:%M:%S') if close_max else None,
            'n_rows_total': n_selected if (include_mrfs is not None or exclude_mrfs is not None) else n_read,
            'n_rows_our': n_our,
            'current_week': current_week,
            'current_month': current_month,
            'closed_months': closed_months,
            'open_month': open_month,
            'open_weeks': open_weeks,
        },
        'weeks': weeks_info,
        'months': months_info,
        'timeline': timeline,
        'registered': registered,
        'closed': closed,
        'cohort': cohort_out,
        'open': {
            'total': open_to_dict(open_snap['total']),
            'direction': {k: open_to_dict(v) for k, v in open_snap['direction'].items()},
            'director': {k: open_to_dict(v) for k, v in open_snap['director'].items()},
            'teamlead': {k: open_to_dict(v) for k, v in open_snap['teamlead'].items()},
            'employee': {k: open_to_dict(v) for k, v in open_snap['employee'].items()},
        },
        'daily': daily_out,
        'employees': emp_list,
        'teamleads': teamleads,
        'directors': directors,
        'directions': directions,
        'mrfs': mrfs,
        'repeat_summary': {
            'n_total_our': n_our,
            'n_repeat': repeat_attributed,
            'repeat_pct': round((100.0 * repeat_attributed / n_our), 2) if n_our else 0.0,
            'columns_found': {
                'comment': True,
                'phone': True,
            },
            'debug': {
                'n_repeat_rows_raw': repeat_total,
                'n_missing_min_dt': repeat_attribution_missing_min_dt,
            },
        },
        'raw': raw_export,
    }

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',',':'))
    sz = os.path.getsize(out_path)
    print(f"[ETL COMBINED] Saved {out_path}, {sz:,} bytes, {time.time()-t0:.1f}s")


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_xlsx = os.path.join(root, 'new', 'Заявки+обращения_объединено.xlsx')
    mapping_txt = os.path.join(root, 'muz_to_combined_mapping.txt')

    hierarchy_xlsx = os.path.join(root, 'config', 'employee_hierarchy.xlsx')
    hierarchy_json = os.path.join(root, 'config', 'employee_hierarchy.json')
    hierarchy = hierarchy_xlsx if os.path.exists(hierarchy_xlsx) else hierarchy_json

    if not os.path.exists(input_xlsx) or not os.path.exists(mapping_txt) or not os.path.exists(hierarchy):
        print(f"ERROR: input={input_xlsx}, mapping={mapping_txt}, hierarchy={hierarchy}", file=sys.stderr)
        sys.exit(1)
    out = os.environ.get('DASH_OUT_PATH') or os.path.join(root, 'data', 'processed', 'daily_dashboard_data.json')
    run_etl_combined(
        input_xlsx,
        mapping_txt,
        hierarchy,
        out,
        include_mrfs=_env_list('DASH_INCLUDE_MRFS'),
        exclude_mrfs=_env_list('DASH_EXCLUDE_MRFS'),
    )
    return


if __name__ == '__main__':
    main()
