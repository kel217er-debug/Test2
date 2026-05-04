"""Period keys, labels, and timeline construction."""

from datetime import datetime, timedelta


def iso_week_key(dt):
    y, w, _ = dt.isocalendar()
    monday = dt - timedelta(days=dt.isoweekday() - 1)
    monday = datetime(monday.year, monday.month, monday.day)
    sunday = monday + timedelta(days=6, hours=23, minutes=59, seconds=59)
    return f"{y}-W{w:02d}", monday, sunday


def month_key(dt):
    return f"{dt.year}-{dt.month:02d}"


def month_label(mkey):
    y, m = map(int, mkey.split('-'))
    names = ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
             'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    return names[m] if 1 <= m <= 12 else mkey


def month_bounds(mkey):
    y, m = map(int, mkey.split('-'))
    start = datetime(y, m, 1)
    if m == 12:
        end = datetime(y + 1, 1, 1) - timedelta(seconds=1)
    else:
        end = datetime(y, m + 1, 1) - timedelta(seconds=1)
    return start, end


def build_weeks_info(weeks_sorted, weeks_meta):
    return [{
        'key': w,
        'start': weeks_meta[w][0].strftime('%Y-%m-%d'),
        'end': weeks_meta[w][1].strftime('%Y-%m-%d'),
        'label': f"{weeks_meta[w][0].strftime('%d.%m')}–{weeks_meta[w][1].strftime('%d.%m')}",
    } for w in weeks_sorted]


def build_months_info(months_sorted, months_meta):
    out = []
    for mk in months_sorted:
        s, e = months_meta[mk]
        out.append({
            'key': mk,
            'start': s.strftime('%Y-%m-%d'),
            'end': e.strftime('%Y-%m-%d'),
            'label': month_label(mk),
            'label_full': f"{month_label(mk)} {mk.split('-')[0]}",
        })
    return out


def split_timeline_periods(dates_max, weeks_sorted, weeks_meta, months_sorted, months_meta):
    current_week, _, _ = iso_week_key(dates_max)
    current_month = month_key(dates_max)
    closed_months = [mk for mk in months_sorted if mk < current_month]
    open_month = current_month
    s_om, e_om = months_meta[open_month]
    open_weeks = [w for w in weeks_sorted if weeks_meta[w][0] <= e_om and weeks_meta[w][1] >= s_om]
    return current_week, current_month, closed_months, open_month, open_weeks


def build_timeline(closed_months, open_weeks, months_meta, weeks_meta, current_week):
    timeline = []
    for mk in closed_months:
        s, e = months_meta[mk]
        timeline.append({
            'key': mk, 'kind': 'month',
            'label': month_label(mk),
            'label_full': f"{month_label(mk)} {mk.split('-')[0]}",
            'start': s.strftime('%Y-%m-%d'), 'end': e.strftime('%Y-%m-%d'),
            'closed': True,
        })
    for w in open_weeks:
        ws, we = weeks_meta[w]
        timeline.append({
            'key': w, 'kind': 'week',
            'label': f"{ws.strftime('%d.%m')}–{we.strftime('%d.%m')}",
            'label_full': f"{w} · {ws.strftime('%d.%m')}–{we.strftime('%d.%m')}",
            'start': ws.strftime('%Y-%m-%d'), 'end': we.strftime('%Y-%m-%d'),
            'closed': (w != current_week),
        })
    return timeline
