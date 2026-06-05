#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds one dashboard dataset: MUZ for all MRFs except RD Center, combined source for RD Center."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from dashboard_logic.periods import (
    build_months_info,
    build_timeline,
    build_weeks_info,
    month_bounds,
    split_timeline_periods,
)
from prepare_upsell_data import build_upsell_dataset


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RD_CENTER = "\u0420\u0414 \u0426\u0435\u043d\u0442\u0440"
ROMASHOV_TEAMLEAD = "\u0420\u043e\u043c\u0430\u0448\u043e\u0432 \u041f\u0430\u0432\u0435\u043b \u0410\u043b\u0435\u043a\u0441\u0430\u043d\u0434\u0440\u043e\u0432\u0438\u0447"

OUT_PATH = PROJECT_ROOT / "data" / "processed" / "daily_dashboard_data.json"
MUZ_PART_PATH = PROJECT_ROOT / "data" / "processed" / "hybrid_muz_without_rd_center.json"
RD_CENTER_PART_PATH = PROJECT_ROOT / "lm" / "rd_center_daily_dashboard_data.json"


def pct(a, b):
    return round(100.0 * a / b, 2) if b else 0.0


def _run_part(label: str, script: Path, cwd: Path, out_path: Path, *, include=None, exclude=None, include_teamleads=None, exclude_teamleads=None):
    env = os.environ.copy()
    env["DASH_OUT_PATH"] = str(out_path)
    if include:
        env["DASH_INCLUDE_MRFS"] = "|".join(include)
    if exclude:
        env["DASH_EXCLUDE_MRFS"] = "|".join(exclude)
    if include_teamleads:
        env["DASH_INCLUDE_TEAMLEADS"] = "|".join(include_teamleads)
    if exclude_teamleads:
        env["DASH_EXCLUDE_TEAMLEADS"] = "|".join(exclude_teamleads)

    print(f"[HYBRID] {label}: start {script}", flush=True)
    result = subprocess.run([sys.executable, "-u", str(script)], cwd=str(cwd), env=env)
    if result.returncode != 0:
        raise RuntimeError(f"{label} failed with code {result.returncode}")
    if not out_path.exists():
        raise FileNotFoundError(out_path)
    print(f"[HYBRID] {label}: OK {out_path} ({out_path.stat().st_size:,} bytes)", flush=True)


def _load(path: Path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _sum_dicts(a, b):
    out = dict(a or {})
    for k, v in (b or {}).items():
        out[k] = out.get(k, 0) + v
    return out


REG_COUNTERS = (
    "n_total",
    "n_base_all",
    "n_stayed_base",
    "n_transferred",
    "n_sla_acc_viol",
    "n_sla_cont_viol",
    "n_repeat",
    "n_connected_stayed_from_period",
    "nd_sum_from_period",
    "nd_sum_created",
    "n_month_conv_primary_total",
    "n_month_conv_primary_connected",
    "n_month_conv_exec_total",
    "n_month_conv_exec_connected",
    "n_conv_exec_total",
    "n_conv_exec_connected",
    "n_conv_primary_total",
    "n_conv_primary_connected",
    "n_emp_tab_total",
    "n_emp_tab_connected",
    "n_conv_clean_total",
    "n_conv_clean_connected",
    "n_month_conv_clean_total",
    "n_month_conv_clean_connected",
    "n_conv_primary_total_all",
    "n_month_conv_primary_total_all",
    "n_conv_clean_exec_connected",
    "n_month_conv_clean_exec_connected",
)


def _clean_counter(metric, clean_key, legacy_key):
    metric = metric or {}
    value = metric.get(clean_key, 0)
    if value:
        return value
    return metric.get(legacy_key, 0)


def _merge_reg_metric(a, b):
    m = {k: (a or {}).get(k, 0) + (b or {}).get(k, 0) for k in REG_COUNTERS}
    m["n_conv_clean_connected"] = _clean_counter(a, "n_conv_clean_connected", "n_conv_clean_exec_connected") + _clean_counter(
        b, "n_conv_clean_connected", "n_conv_clean_exec_connected"
    )
    m["n_month_conv_clean_connected"] = _clean_counter(
        a, "n_month_conv_clean_connected", "n_month_conv_clean_exec_connected"
    ) + _clean_counter(b, "n_month_conv_clean_connected", "n_month_conv_clean_exec_connected")
    m["n_conv_clean_total"] = m["n_conv_primary_total"]
    m["n_month_conv_clean_total"] = m["n_month_conv_primary_total"]
    n_t = m["n_total"]
    n_exec_total = m["n_conv_exec_total"]
    n_exec_conn = m["n_conv_exec_connected"]
    n_prim_total = m["n_conv_primary_total"]
    n_prim_conn = m["n_conv_primary_connected"]
    n_m_prim_total = m["n_month_conv_primary_total"]
    n_m_exec_total = m["n_month_conv_exec_total"]
    if n_m_prim_total or n_m_exec_total:
        conv_clean = pct(m["n_month_conv_clean_connected"], n_m_prim_total)
        conv_regular = pct(m["n_month_conv_exec_connected"], n_m_exec_total)
    else:
        conv_clean = pct(m["n_conv_clean_connected"], n_prim_total)
        conv_regular = pct(n_exec_conn, n_exec_total)
    m.update(
        {
            "transfer_rate": pct(m["n_transferred"], n_t),
            "sla_acc_rate": pct(m["n_sla_acc_viol"], n_t),
            "sla_cont_rate": pct(m["n_sla_cont_viol"], n_t),
            "repeat_pct": pct(m["n_repeat"], n_t),
            "conv_clean": conv_clean,
            "conv_regular": conv_regular,
            "conv_exec_pct": pct(n_exec_conn, n_exec_total),
            "conv_primary_pct": conv_clean,
            "nd_sum_from_period": round(m["nd_sum_from_period"], 2),
            "nd_sum_created": round(m.get("nd_sum_created", 0.0), 2),
        }
    )
    return m


CLOSED_COUNTERS = ("n_closed", "n_connected", "n_with_equip", "nd_sum", "n_inn_connected", "n_inn_multi", "n_secondary", "n_ultra_cheap")


def _merge_closed_metric(a, b):
    m = {k: (a or {}).get(k, 0) + (b or {}).get(k, 0) for k in CLOSED_COUNTERS}
    m["nd_sum"] = round(m["nd_sum"], 2)
    m["by_service"] = _sum_dicts((a or {}).get("by_service"), (b or {}).get("by_service"))
    m["by_service_equip"] = _sum_dicts((a or {}).get("by_service_equip"), (b or {}).get("by_service_equip"))
    m["close_conv"] = pct(m["n_connected"], m["n_closed"])
    m["equip_share"] = pct(m["n_with_equip"], m["n_connected"])
    m["crossell_pct"] = pct(m["n_inn_multi"], m["n_inn_connected"])
    m["fraud_pct"] = pct(m["n_ultra_cheap"], m["n_secondary"])
    return m


def _merge_open_metric(a, b):
    m = {
        "n_open": (a or {}).get("n_open", 0) + (b or {}).get("n_open", 0),
        "n_stale30": (a or {}).get("n_stale30", 0) + (b or {}).get("n_stale30", 0),
        "by_week": _sum_dicts((a or {}).get("by_week"), (b or {}).get("by_week")),
        "stale_by_week": _sum_dicts((a or {}).get("stale_by_week"), (b or {}).get("stale_by_week")),
    }
    m["stale_rate"] = pct(m["n_stale30"], m["n_open"])
    return m


def _merge_scoped_periods(a, b, metric_fn):
    out = {}
    for per in sorted(set((a or {}).keys()) | set((b or {}).keys())):
        out[per] = {}
        for scope in ("total", "direction", "director", "teamlead", "employee", "employee_exec"):
            left = ((a or {}).get(per) or {}).get(scope, {})
            right = ((b or {}).get(per) or {}).get(scope, {})
            out[per][scope] = {
                name: metric_fn(left.get(name), right.get(name))
                for name in sorted(set(left.keys()) | set(right.keys()))
            }
    return out


def _merge_open(a, b):
    out = {"total": _merge_open_metric((a or {}).get("total"), (b or {}).get("total"))}
    for scope in ("direction", "director", "teamlead", "employee"):
        left = (a or {}).get(scope, {})
        right = (b or {}).get(scope, {})
        out[scope] = {
            name: _merge_open_metric(left.get(name), right.get(name))
            for name in sorted(set(left.keys()) | set(right.keys()))
        }
    return out


def _merge_daily(a, b):
    out = {}
    for day in sorted(set((a or {}).keys()) | set((b or {}).keys())):
        left = (a or {}).get(day, {})
        right = (b or {}).get(day, {})
        out[day] = {
            "n_total": left.get("n_total", 0) + right.get("n_total", 0),
            "n_connected": left.get("n_connected", 0) + right.get("n_connected", 0),
            "nd_sum": round(left.get("nd_sum", 0.0) + right.get("nd_sum", 0.0), 2),
        }
    return out


def _merge_cohort(a, b):
    left = ((a or {}).get("by_month") or {})
    right = ((b or {}).get("by_month") or {})
    out = {}
    for cm in sorted(set(left.keys()) | set(right.keys())):
        out[cm] = {}
        regs = set((left.get(cm) or {}).keys()) | set((right.get(cm) or {}).keys())
        for rm in sorted(regs):
            lv = (left.get(cm) or {}).get(rm, {})
            rv = (right.get(cm) or {}).get(rm, {})
            out[cm][rm] = {
                "n_connected": lv.get("n_connected", 0) + rv.get("n_connected", 0),
                "nd_sum": round(lv.get("nd_sum", 0.0) + rv.get("nd_sum", 0.0), 2),
                "n_with_equip": lv.get("n_with_equip", 0) + rv.get("n_with_equip", 0),
            }
    return {"by_month": out}


def _merge_employees(a, b):
    by_name = {}
    for item in (a or []) + (b or []):
        name = item.get("name")
        if not name:
            continue
        prev = by_name.get(name)
        if prev is None:
            by_name[name] = dict(item)
        else:
            prev["in_data"] = bool(prev.get("in_data")) or bool(item.get("in_data"))
    return sorted(by_name.values(), key=lambda x: x.get("name", ""))


def _date_value(s):
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def _build_time_meta(merged):
    weeks_meta = {}
    for w in merged["weeks"]:
        weeks_meta[w["key"]] = (
            datetime.strptime(w["start"], "%Y-%m-%d"),
            datetime.strptime(w["end"], "%Y-%m-%d"),
        )
    months_meta = {}
    for m in merged["months"]:
        months_meta[m["key"]] = month_bounds(m["key"])

    data_to = _date_value(merged["meta"].get("data_to"))
    if not data_to:
        return

    weeks_sorted = sorted(weeks_meta.keys())
    months_sorted = sorted(months_meta.keys())
    current_week, current_month, closed_months, open_month, open_weeks = split_timeline_periods(
        data_to, weeks_sorted, weeks_meta, months_sorted, months_meta
    )
    merged["weeks"] = build_weeks_info(weeks_sorted, weeks_meta)
    merged["months"] = build_months_info(months_sorted, months_meta)
    merged["timeline"] = build_timeline(closed_months, open_month, open_weeks, months_meta, weeks_meta, current_week)
    merged["meta"].update(
        {
            "current_week": current_week,
            "current_month": current_month,
            "closed_months": closed_months,
            "open_month": open_month,
            "open_weeks": open_weeks,
        }
    )


def _merge_data(muz, rd):
    merged = dict(muz)
    merged["meta"] = dict(muz.get("meta", {}))
    merged["meta"].update(
        {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "source_kind": "hybrid_muz_with_combined_rd_center",
            "source_file": "data/source/muz.xlsx + lm/\u0417\u0430\u044f\u0432\u043a\u0438+\u043e\u0431\u0440\u0430\u0449\u0435\u043d\u0438\u044f_\u043e\u0431\u044a\u0435\u0434\u0438\u043d\u0435\u043d\u043e.xlsx",
            "rd_center_source_file": os.path.relpath(RD_CENTER_PART_PATH, PROJECT_ROOT),
        }
    )
    for key in ("data_from", "close_from"):
        vals = [v for v in (muz.get("meta", {}).get(key), rd.get("meta", {}).get(key)) if v]
        merged["meta"][key] = min(vals) if vals else None
    for key in ("data_to", "close_to"):
        vals = [v for v in (muz.get("meta", {}).get(key), rd.get("meta", {}).get(key)) if v]
        merged["meta"][key] = max(vals) if vals else None
    merged["meta"]["n_rows_total"] = muz.get("meta", {}).get("n_rows_total", 0) + rd.get("meta", {}).get("n_rows_total", 0)
    merged["meta"]["n_rows_our"] = muz.get("meta", {}).get("n_rows_our", 0) + rd.get("meta", {}).get("n_rows_our", 0)

    merged["registered"] = {
        "by_week": _merge_scoped_periods(muz.get("registered", {}).get("by_week"), rd.get("registered", {}).get("by_week"), _merge_reg_metric),
        "by_month": _merge_scoped_periods(muz.get("registered", {}).get("by_month"), rd.get("registered", {}).get("by_month"), _merge_reg_metric),
    }
    merged["closed"] = {
        "by_week": _merge_scoped_periods(muz.get("closed", {}).get("by_week"), rd.get("closed", {}).get("by_week"), _merge_closed_metric),
        "by_month": _merge_scoped_periods(muz.get("closed", {}).get("by_month"), rd.get("closed", {}).get("by_month"), _merge_closed_metric),
    }
    merged["open"] = _merge_open(muz.get("open"), rd.get("open"))
    merged["daily"] = _merge_daily(muz.get("daily"), rd.get("daily"))
    merged["cohort"] = _merge_cohort(muz.get("cohort"), rd.get("cohort"))

    employees = _merge_employees(muz.get("employees"), rd.get("employees"))
    merged["employees"] = employees
    merged["teamleads"] = sorted(set(muz.get("teamleads", [])) | set(rd.get("teamleads", [])))
    merged["directors"] = sorted(set(muz.get("directors", [])) | set(rd.get("directors", [])))
    merged["directions"] = sorted({e.get("direction", "") for e in employees if e.get("direction")})
    merged["mrfs"] = sorted({e.get("mrf", "") for e in employees if e.get("mrf")})

    merged["weeks"] = sorted({w["key"]: w for w in muz.get("weeks", []) + rd.get("weeks", [])}.values(), key=lambda x: x["key"])
    merged["months"] = sorted({m["key"]: m for m in muz.get("months", []) + rd.get("months", [])}.values(), key=lambda x: x["key"])
    _build_time_meta(merged)

    repeat_n = muz.get("repeat_summary", {}).get("n_repeat", 0) + rd.get("repeat_summary", {}).get("n_repeat", 0)
    repeat_total = merged["meta"]["n_rows_our"]
    merged["repeat_summary"] = {
        "n_total_our": repeat_total,
        "n_repeat": repeat_n,
        "repeat_pct": pct(repeat_n, repeat_total),
        "sources": {
            "muz_without_rd_center": muz.get("repeat_summary", {}),
            "rd_center_combined": rd.get("repeat_summary", {}),
        },
    }

    raw_rows = []
    raw_headers = []
    for source_name, data in (("muz", muz), ("rd_center", rd)):
        raw = data.get("raw") or {}
        if not raw_headers and raw.get("headers"):
            raw_headers = raw.get("headers")
        for row in raw.get("rows", []):
            row = dict(row)
            row["source_dataset"] = source_name
            raw_rows.append(row)
    merged["raw"] = {"headers": raw_headers, "rows": raw_rows}
    return merged


def main() -> int:
    started = time.time()
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    RD_CENTER_PART_PATH.parent.mkdir(parents=True, exist_ok=True)

    _run_part(
        "MUZ without RD Center",
        PROJECT_ROOT / "scripts" / "prepare_dashboard_data_muz.py",
        PROJECT_ROOT,
        MUZ_PART_PATH,
        exclude=[RD_CENTER],
        include_teamleads=[ROMASHOV_TEAMLEAD],
    )
    _run_part(
        "RD Center from combined source",
        PROJECT_ROOT / "scripts" / "prepare_dashboard_data_combined.py",
        PROJECT_ROOT,
        RD_CENTER_PART_PATH,
        include=[RD_CENTER],
        exclude_teamleads=[ROMASHOV_TEAMLEAD],
    )

    merged = _merge_data(_load(MUZ_PART_PATH), _load(RD_CENTER_PART_PATH))
    hierarchy_xlsx = PROJECT_ROOT / "config" / "employee_hierarchy.xlsx"
    hierarchy_json = PROJECT_ROOT / "config" / "employee_hierarchy.json"
    hierarchy_path = hierarchy_xlsx if hierarchy_xlsx.exists() else hierarchy_json
    merged["upsell"] = build_upsell_dataset(str(PROJECT_ROOT / "data" / "source" / "muz.xlsx"), str(hierarchy_path))
    with OUT_PATH.open("w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, separators=(",", ":"))
    try:
        MUZ_PART_PATH.unlink()
    except OSError:
        pass
    print(f"[HYBRID] Saved {OUT_PATH} ({OUT_PATH.stat().st_size:,} bytes), elapsed={time.time() - started:.1f}s", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
