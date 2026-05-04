#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETL v2 для ежедневного дашборда тимлидов.

Ключевые изменения v2:
- Две независимые агрегации:
  * REGISTERED (по дате регистрации) — leading metrics:
    заявки, SLA, передачи, нагрузка, открытые/застрявшие
  * CLOSED (по дате перевода в итоговый статус, col 127) — lagging metrics:
    подключения, НД, услуги, оборудование
- Периоды:
  * Закрытые месяцы — один тотал на месяц
  * Текущий (открытый) месяц — понедельная разбивка
- Cohort-матрица (close_month × reg_month) — видно, из каких недель
  регистрации закрываются заявки в данном периоде
"""

import os, sys, re, json, time
from datetime import datetime
from collections import defaultdict

from dashboard_logic.aggregations import (
    new_closed_metrics,
    new_metric_scope,
    new_open_rec,
    new_reg_metrics,
    serialize_per,
)
from dashboard_logic.excel_filters import (
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

STALE_HOURS = 720  # 30 дней

# -------- helpers --------
_DT_RE = re.compile(r'^\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})\s*$')

def parse_dt(v):
    if v is None: return None
    if isinstance(v, datetime): return v
    if isinstance(v, str):
        s = v.strip()
        if not s: return None
        m = _DT_RE.match(s)
        if m:
            d, mo, y, H, M, S = map(int, m.groups())
            try: return datetime(y, mo, d, H, M, S)
            except ValueError: return None
    return None

def parse_num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace('\xa0','').replace(' ','').replace(',','.')
    if not s: return 0.0
    try: return float(s)
    except: return 0.0

def norm_str(v):
    return '' if v is None else str(v).strip()

def clean_name(v):
    s = norm_str(v)
    if not s: return ''
    if '@' in s:
        parts = s.split()
        clean = []
        for p in parts:
            if '@' in p: break
            clean.append(p)
        s = ' '.join(clean)
    return s.strip()

# -------- иерархия --------

def load_hierarchy(hierarchy_path):
    with open(hierarchy_path, 'r', encoding='utf-8') as f:
        raw = json.load(f)

    employees = raw.get('employees', raw if isinstance(raw, list) else [])
    emp_map = {}
    for info in employees:
        name = (info.get('name') or info.get('manager') or '').strip()
        if not name:
            continue
        emp_map[name] = {
            'teamlead': info.get('teamlead') or '',
            'director': info.get('director') or '',
            'direction': info.get('direction') or '',
            'mrf': info.get('mrf') or '',
            'is_active': info.get('is_active', True),
        }
    return emp_map

# -------- основной проход --------

def run_etl(muz_path, hierarchy_path, out_path):
    t0 = time.time()
    print(f"[ETL v2] Start. muz={muz_path}")

    emp_map = load_hierarchy(hierarchy_path)
    print(f"[ETL v2] Hierarchy: {len(emp_map)} employees from {hierarchy_path}")

    # Calamine
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(muz_path)
    sheet = wb.get_sheet_by_name('Sheet1')
    rows_iter = iter(sheet.to_python())
    _header = next(rows_iter)

    # --- структуры ---
    # Для каждой корзины и периода — метрики
    # weekly (registered) для графиков-sparklines — по неделям, 5 срезов:
    weekly_reg = defaultdict(lambda: new_metric_scope(new_reg_metrics))
    # weekly (closed) — по неделям
    weekly_closed = defaultdict(lambda: new_metric_scope(new_closed_metrics))
    # monthly — то же по месяцам
    monthly_reg = defaultdict(lambda: new_metric_scope(new_reg_metrics))
    monthly_closed = defaultdict(lambda: new_metric_scope(new_closed_metrics))
    # Cohort: close_month → reg_month → {n_connected, nd_sum, n_with_equip}
    cohort = defaultdict(lambda: defaultdict(lambda: {'n_connected': 0, 'nd_sum': 0.0, 'n_with_equip': 0}))

    # Снэпшот открытых
    open_snap = {
        'total': new_open_rec(),
        'direction': defaultdict(new_open_rec),
        'director': defaultdict(new_open_rec),
        'teamlead': defaultdict(new_open_rec),
        'employee': defaultdict(new_open_rec),
    }
    # Дневная динамика (для sparkline главной) — 2 среза: registered и closed
    daily_reg = defaultdict(lambda: {'n_total': 0})
    daily_closed = defaultdict(lambda: {'n_connected': 0, 'nd_sum': 0.0})

    weeks_meta = {}      # wkey -> (monday, sunday)
    months_meta = {}     # mkey -> (start, end)

    n_read = 0; n_our = 0; n_skip_no_date = 0
    dates_min = None; dates_max = None
    close_min = None; close_max = None

    t_iter = time.time()
    for row in rows_iter:
        n_read += 1
        if n_read % 20000 == 0:
            print(f"  ... {n_read} rows, {time.time()-t_iter:.1f}s")

        podr_prim = norm_str(row[COL['primary_podr']])
        if not is_our_channel(podr_prim):
            continue
        n_our += 1

        reg_dt = parse_dt(row[COL['reg_dt']])
        if reg_dt is None:
            n_skip_no_date += 1
            continue

        if dates_min is None or reg_dt < dates_min: dates_min = reg_dt
        if dates_max is None or reg_dt > dates_max: dates_max = reg_dt

        direction = direction_for_channel(podr_prim)
        service = norm_str(row[COL['service']]).lower()
        service_raw = norm_str(row[COL['service']]) or '—'
        current_podr = norm_str(row[COL['current_podr']])
        _stayed = is_our_channel(current_podr)
        _transferred = is_transferred(current_podr)
        final_status = norm_str(row[COL['final_status']])
        _connected = is_connected(final_status)
        _has_final = bool(final_status)
        _in_conv_base = is_base_service(service)
        sla_acc = norm_str(row[COL['sla_acc']]).lower()
        sla_cont = norm_str(row[COL['sla_cont']]).lower()
        _sla_acc_viol = is_sla_violated(sla_acc)
        _sla_cont_viol = is_sla_violated(sla_cont)
        equip_price = parse_num(row[COL['equip_price']])
        _with_equip = _connected and (equip_price > 0)
        # Защита от outliers: отдельные записи содержат явные ошибки ввода (НД > 1 млрд руб/мес)
        ND_OUTLIER_CAP = 1_000_000_000  # 1 млрд руб — выше этого считаем ошибкой ввода и отбрасываем
        if _connected:
            inst = parse_num(row[COL['install_amount']])
            mon_ = parse_num(row[COL['monthly_amount']])
            if inst > ND_OUTLIER_CAP: inst = 0.0
            if mon_ > ND_OUTLIER_CAP: mon_ = 0.0
            nd_val = inst + mon_
        else:
            inst = 0.0; mon_ = 0.0
            nd_val = 0.0

        # Данные для cross-sell / фрод-индикатора
        inn_val = norm_str(row[COL['inn']])
        _service_lower = service  # уже нижний регистр
        _is_internet = is_internet_service(_service_lower)
        # В фрод-корзину попадают только подключённые услуги, не относящиеся к EXCL_SERVICES
        _cross_eligible = _connected and _in_conv_base and bool(inn_val)

        primary_name = clean_name(row[COL['primary_name']])
        emp_info = emp_map.get(primary_name, {})
        teamlead = emp_info.get('teamlead', '')
        director = emp_info.get('director', '')

        # Периоды регистрации
        wkey_reg, mon_r, sun_r = iso_week_key(reg_dt)
        mkey_reg = month_key(reg_dt)
        if wkey_reg not in weeks_meta: weeks_meta[wkey_reg] = (mon_r, sun_r)
        if mkey_reg not in months_meta: months_meta[mkey_reg] = month_bounds(mkey_reg)

        # ------ REGISTERED агрегация ------
        def apply_reg(m):
            apply_reg_metrics(
                m,
                _in_conv_base,
                _stayed,
                _transferred,
                _connected,
                _sla_acc_viol,
                _sla_cont_viol,
                nd_val,
            )

        for ag, key, name in [
            (weekly_reg[wkey_reg], 'total', 'all'),
            (monthly_reg[mkey_reg], 'total', 'all'),
        ]:
            apply_reg(ag[key][name])
        if direction:
            apply_reg(weekly_reg[wkey_reg]['direction'][direction])
            apply_reg(monthly_reg[mkey_reg]['direction'][direction])
        if director:
            apply_reg(weekly_reg[wkey_reg]['director'][director])
            apply_reg(monthly_reg[mkey_reg]['director'][director])
        if teamlead:
            apply_reg(weekly_reg[wkey_reg]['teamlead'][teamlead])
            apply_reg(monthly_reg[mkey_reg]['teamlead'][teamlead])
        if primary_name:
            apply_reg(weekly_reg[wkey_reg]['employee'][primary_name])
            apply_reg(monthly_reg[mkey_reg]['employee'][primary_name])

        daily_reg[reg_dt.strftime('%Y-%m-%d')]['n_total'] += 1

        # ------ CLOSED агрегация (по дате итогового статуса) ------
        if _has_final:
            final_dt = parse_dt(row[COL['final_dt']])
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
                if direction:
                    apply_closed(weekly_closed[wkey_cl]['direction'][direction])
                    apply_closed(monthly_closed[mkey_cl]['direction'][direction])
                if director:
                    apply_closed(weekly_closed[wkey_cl]['director'][director])
                    apply_closed(monthly_closed[mkey_cl]['director'][director])
                if teamlead:
                    apply_closed(weekly_closed[wkey_cl]['teamlead'][teamlead])
                    apply_closed(monthly_closed[mkey_cl]['teamlead'][teamlead])
                if primary_name:
                    apply_closed(weekly_closed[wkey_cl]['employee'][primary_name])
                    apply_closed(monthly_closed[mkey_cl]['employee'][primary_name])

                # Cohort — только для подключений, в разрезе месяцев
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
            if direction: apply_open(open_snap['direction'][direction])
            if director: apply_open(open_snap['director'][director])
            if teamlead: apply_open(open_snap['teamlead'][teamlead])
            if primary_name: apply_open(open_snap['employee'][primary_name])

    print(f"[ETL v2] Read {n_read}, our={n_our}, skip_no_date={n_skip_no_date}, {time.time()-t0:.1f}s")
    print(f"[ETL v2] Reg dates: {dates_min} — {dates_max}")
    print(f"[ETL v2] Close dates: {close_min} — {close_max}")

    # --------- сериализация ---------

    weeks_sorted = sorted(weeks_meta.keys())
    months_sorted = sorted(months_meta.keys())

    weeks_info = build_weeks_info(weeks_sorted, weeks_meta)
    months_info = build_months_info(months_sorted, months_meta)

    # Определения текущего/открытого
    current_week, current_month, closed_months, open_month, open_weeks = split_timeline_periods(
        dates_max,
        weeks_sorted,
        weeks_meta,
        months_sorted,
        months_meta,
    )

    registered = {
        'by_week': serialize_per(weekly_reg, reg_dict),
        'by_month': serialize_per(monthly_reg, reg_dict),
    }
    closed = {
        'by_week': serialize_per(weekly_closed, closed_dict),
        'by_month': serialize_per(monthly_closed, closed_dict),
    }

    # Cohort: список записей {close_period, reg_period, n_connected, nd_sum, n_with_equip}
    # + агрегированная матрица only for months
    cohort_out = {
        'by_month': {cm: {rm: dict(v) for rm, v in rs.items()} for cm, rs in cohort.items()}
    }
    # Округлим nd_sum в cohort
    for cm, rs in cohort_out['by_month'].items():
        for rm, v in rs.items():
            v['nd_sum'] = round(v['nd_sum'], 2)

    open_out = {
        'total': open_to_dict(open_snap['total']),
        'direction': {k:open_to_dict(v) for k,v in open_snap['direction'].items()},
        'director': {k:open_to_dict(v) for k,v in open_snap['director'].items()},
        'teamlead': {k:open_to_dict(v) for k,v in open_snap['teamlead'].items()},
        'employee': {k:open_to_dict(v) for k,v in open_snap['employee'].items()},
    }

    # Daily (для sparklines главной): объединим reg + closed в одну структуру
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

    # Сотрудники
    seen_names = set()
    for w in weeks_sorted:
        seen_names.update(weekly_reg[w]['employee'].keys())
    emp_list = []
    for name, info in emp_map.items():
        emp_list.append({
            'name': name,
            'teamlead': info['teamlead'],
            'director': info['director'],
            'direction': info['direction'],
            'mrf': info['mrf'],
            'is_active': info['is_active'],
            'in_data': name in seen_names,
        })
    for name in seen_names:
        if name and name not in emp_map:
            emp_list.append({
                'name': name, 'teamlead': '', 'director': '',
                'direction': '', 'mrf': '', 'is_active': True, 'in_data': True,
            })

    teamleads = sorted({e['teamlead'] for e in emp_list if e['teamlead']})
    directors = sorted({e['director'] for e in emp_list if e['director']})
    directions = sorted({e['direction'] for e in emp_list if e['direction']})
    mrfs = sorted({e['mrf'] for e in emp_list if e['mrf']})

    timeline = build_timeline(closed_months, open_weeks, months_meta, weeks_meta, current_week)

    out = {
        'meta': {
            'generated_at': datetime.now().isoformat(timespec='seconds'),
            'data_from': dates_min.strftime('%Y-%m-%d %H:%M:%S') if dates_min else None,
            'data_to': dates_max.strftime('%Y-%m-%d %H:%M:%S') if dates_max else None,
            'close_from': close_min.strftime('%Y-%m-%d %H:%M:%S') if close_min else None,
            'close_to': close_max.strftime('%Y-%m-%d %H:%M:%S') if close_max else None,
            'n_rows_total': n_read,
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
        'open': open_out,
        'daily': daily_out,
        'employees': emp_list,
        'teamleads': teamleads,
        'directors': directors,
        'directions': directions,
        'mrfs': mrfs,
    }

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',',':'))
    sz = os.path.getsize(out_path)
    print(f"[ETL v2] Saved {out_path}, {sz:,} bytes, {time.time()-t0:.1f}s")
    print(f"[ETL v2] Months: {len(months_info)} (closed={len(closed_months)}, open={open_month})")
    print(f"[ETL v2] Open weeks in current month: {open_weeks}")
    print(f"[ETL v2] Teamleads={len(teamleads)} Directors={len(directors)}")


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    source_dir = os.path.join(root, 'data', 'source')
    muz = None
    if os.path.isdir(source_dir):
        for f in os.listdir(source_dir):
            lower_name = f.lower()
            if lower_name.endswith('.xlsx') and ('muz' in lower_name or 'муз' in lower_name):
                muz = os.path.join(source_dir, f)
                break

    hierarchy = os.path.join(root, 'config', 'employee_hierarchy.json')
    if not muz or not os.path.exists(hierarchy):
        print(f"ERROR: muz={muz}, hierarchy={hierarchy}", file=sys.stderr)
        sys.exit(1)
    out = os.path.join(root, 'data', 'processed', 'daily_dashboard_data.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    run_etl(muz, hierarchy, out)
    return


if __name__ == '__main__':
    main()
