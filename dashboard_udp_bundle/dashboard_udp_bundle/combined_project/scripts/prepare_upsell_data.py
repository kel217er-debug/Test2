#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds data for the upsell tab from the main MUZ source."""

from __future__ import annotations

import json
import os
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

import openpyxl
from python_calamine import CalamineWorkbook

from dashboard_logic.excel_filters import COL, CONNECTED_STATUS
from prepare_dashboard_data_muz import clean_name, load_hierarchy, norm_str, parse_dt


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJECT_ROOT_PATH = Path(PROJECT_ROOT)
EXCLUDED_SERVICES = {"Другие услуги", "Заказ оборудования", "Прочее"}
RM_STEMS = {"rm", "\u0440\u043c"}
RM_CONNECTED_STATUS = "\u0423\u0441\u043b\u0443\u0433\u0430 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0430"
SERVICE_ALIASES = {
    "ШПД": "Интернет",
    "ВАТС": "Виртуальная АТС",
    "8800": "Номер 8800",
    "WiFi": "Wi-Fi для бизнеса",
    "ТВ для гостиниц": "Телевидение для гостиниц",
}
_DATE_PREFIX_RE = re.compile(r"^\s*(\d{1,2}\.\d{1,2}\.\d{4})(?:\s+(\d{1,2}:\d{2})(?::(\d{2}))?)?")


def _iter_data_rows(sheet):
    rows = iter(sheet.to_python())
    header_row = None
    for _ in range(50):
        candidate = next(rows, None)
        if candidate is None:
            break
        c0 = norm_str(candidate[0]).lower() if len(candidate) > 0 else ""
        c1 = norm_str(candidate[1]).lower() if len(candidate) > 1 else ""
        if c0 == "мрф" and c1 == "регион":
            header_row = candidate
            break
    if header_row is None:
        raise ValueError("Не удалось найти строку заголовков в muz.xlsx")
    return rows


def _month_key(value: datetime | None) -> str:
    return value.strftime("%Y-%m") if value else ""


def _parse_any_dt(value) -> datetime | None:
    parsed = parse_dt(value)
    if parsed is not None:
        return parsed
    if isinstance(value, date):
        return datetime.combine(value, time())
    text = norm_str(value)
    if not text:
        return None
    match = _DATE_PREFIX_RE.match(text)
    if not match:
        return None
    date_part, hm_part, sec_part = match.groups()
    if hm_part:
        if sec_part:
            text = f"{date_part} {hm_part}:{sec_part}"
            fmt = "%d.%m.%Y %H:%M:%S"
        else:
            text = f"{date_part} {hm_part}"
            fmt = "%d.%m.%Y %H:%M"
    else:
        text = date_part
        fmt = "%d.%m.%Y"
    try:
        return datetime.strptime(text, fmt)
    except ValueError:
        return None


def _previous_month(value: datetime) -> str:
    year = value.year
    month = value.month - 1
    if month == 0:
        year -= 1
        month = 12
    return f"{year:04d}-{month:02d}"


def _first_workday(year: int, month: int) -> date:
    current = date(year, month, 1)
    while current.weekday() >= 5:
        current += timedelta(days=1)
    return current


def _effective_period_month(connection_dt: datetime, registration_value) -> str:
    registration_dt = _parse_any_dt(registration_value)
    connection_month = _month_key(connection_dt)
    if (
        registration_dt
        and _month_key(registration_dt) != connection_month
        and connection_dt.date() == _first_workday(connection_dt.year, connection_dt.month)
    ):
        return _previous_month(connection_dt)
    return connection_month


def _normalize_service(value: str) -> str:
    service = norm_str(value)
    return SERVICE_ALIASES.get(service, service)


def _split_services(value: str) -> list[str]:
    services = []
    for part in norm_str(value).split(","):
        service = _normalize_service(part)
        if service and service not in EXCLUDED_SERVICES:
            services.append(service)
    return services


def _strip_email(value: str) -> str:
    return re.sub(r"\s+\S+@\S+\s*$", "", norm_str(value)).strip()


def _pick_manager(row) -> str:
    for idx in (COL["connector_name"], COL["exec_name"], COL["primary_name"]):
        if idx < len(row):
            name = clean_name(row[idx])
            if name:
                return name
    return ""


def _resolve_manager(raw_manager: str, emp_map: dict, manager_names: list[str]) -> str:
    raw_manager = _strip_email(raw_manager)
    manager = raw_manager if raw_manager in emp_map else ""
    if manager:
        return manager
    for name in manager_names:
        if raw_manager.startswith(name):
            return name
    return ""


def _find_rm_source() -> Path | None:
    search_roots = [
        PROJECT_ROOT_PATH,
        PROJECT_ROOT_PATH / "test",
        PROJECT_ROOT_PATH.parent / "test",
    ]
    for root in search_roots:
        if not root.exists():
            continue
        for candidate in root.glob("*.xlsx"):
            if candidate.name.startswith("~$"):
                continue
            if candidate.stem.casefold() in RM_STEMS:
                return candidate
    return None


def _first_value(row, *indexes: int) -> str:
    for index in indexes:
        if len(row) > index:
            value = norm_str(row[index])
            if value:
                return value
    return ""


def build_upsell_dataset(muz_path: str, hierarchy_path: str) -> dict:
    emp_map = load_hierarchy(hierarchy_path)
    manager_names = sorted(emp_map, key=len, reverse=True)
    wb = CalamineWorkbook.from_path(muz_path)
    sheet = wb.get_sheet_by_name("Sheet1")
    rows = _iter_data_rows(sheet)

    records = []
    periods = set()
    for source_row, row in enumerate(rows, start=8):
        if len(row) <= COL["final_status"]:
            continue

        final_status = norm_str(row[COL["final_status"]])
        if final_status != CONNECTED_STATUS:
            continue

        final_dt = parse_dt(row[COL["final_dt"]])
        if final_dt is None:
            continue

        inn = norm_str(row[COL["inn"]])
        if not inn:
            continue

        services = _split_services(row[COL["service"]])
        if not services:
            continue

        manager = _resolve_manager(_pick_manager(row), emp_map, manager_names)
        if not manager:
            continue

        meta = emp_map.get(manager, {})
        mrf = norm_str(meta.get("mrf")) or (norm_str(row[COL["mrf"]]) if len(row) > COL["mrf"] else "")
        direction = norm_str(meta.get("direction"))
        director = clean_name(meta.get("director"))
        teamlead = clean_name(meta.get("teamlead"))
        reg_dt = _parse_any_dt(row[COL["reg_dt"]]) if len(row) > COL["reg_dt"] else None
        date_key = final_dt.strftime("%Y-%m-%d")
        period_month = _effective_period_month(final_dt, row[COL["reg_dt"]] if len(row) > COL["reg_dt"] else None)
        periods.add(period_month)

        for service_index, service in enumerate(services):
            records.append(
                {
                    "mrf": mrf,
                    "rd": mrf,
                    "direction": direction,
                    "director": director,
                    "teamlead": teamlead,
                    "manager": manager,
                    "inn": inn,
                    "service": service,
                    "phone": norm_str(row[36]) if len(row) > 36 else "",
                    "request": norm_str(row[8]) if len(row) > 8 else "",
                    "date": date_key,
                    "periodMonth": period_month,
                    "registrationDate": reg_dt.strftime("%d.%m.%Y %H:%M") if reg_dt else "",
                    "sourceRow": source_row * 10 + service_index,
                    "source": "МУЗ",
                }
            )

    rm_path = _find_rm_source()
    if rm_path and rm_path.exists():
        rm_wb = openpyxl.load_workbook(rm_path, read_only=True, data_only=True)
        rm_ws = rm_wb.active
        for row_index, row in enumerate(rm_ws.iter_rows(min_row=2, values_only=True), start=2):
            status = norm_str(row[41] if len(row) > 41 else "")
            if status != RM_CONNECTED_STATUS:
                continue

            manager = _resolve_manager(_first_value(row, 52), emp_map, manager_names)
            if not manager:
                continue

            services = _split_services(_first_value(row, 35, 81))
            inn = _first_value(row, 93, 24)
            if not services or not inn:
                continue

            connection_dt = _parse_any_dt(row[49] if len(row) > 49 else None)
            if connection_dt is None:
                continue

            meta = emp_map.get(manager, {})
            mrf = norm_str(meta.get("mrf"))
            direction = norm_str(meta.get("direction"))
            director = clean_name(meta.get("director"))
            teamlead = clean_name(meta.get("teamlead"))
            registration_value = _first_value(row, 70, 8)
            registration_dt = _parse_any_dt(registration_value)
            period_month = _effective_period_month(connection_dt, registration_value)
            periods.add(period_month)

            for service_index, service in enumerate(services):
                records.append(
                    {
                        "mrf": mrf,
                        "rd": mrf,
                        "direction": direction,
                        "director": director,
                        "teamlead": teamlead,
                        "manager": manager,
                        "inn": inn,
                        "service": service,
                        "phone": _first_value(row, 104, 105, 32, 33),
                        "request": _first_value(row, 7, 69, 2),
                        "date": connection_dt.strftime("%Y-%m-%d"),
                        "periodMonth": period_month,
                        "registrationDate": registration_dt.strftime("%d.%m.%Y %H:%M") if registration_dt else norm_str(registration_value),
                        "sourceRow": row_index * 10 + service_index,
                        "source": "РМ",
                    }
                )
        rm_wb.close()

    periods_sorted = sorted(p for p in periods if p)
    return {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": os.path.relpath(muz_path, PROJECT_ROOT),
        "records": records,
        "periods": periods_sorted,
    }


def main() -> int:
    muz_path = os.path.join(PROJECT_ROOT, "data", "source", "muz.xlsx")
    hierarchy_xlsx = os.path.join(PROJECT_ROOT, "config", "employee_hierarchy.xlsx")
    hierarchy_json = os.path.join(PROJECT_ROOT, "config", "employee_hierarchy.json")
    hierarchy_path = hierarchy_xlsx if os.path.exists(hierarchy_xlsx) else hierarchy_json
    out_path = os.environ.get("DASH_UPSELL_OUT_PATH") or os.path.join(
        PROJECT_ROOT, "data", "processed", "upsell_data.json"
    )

    payload = build_upsell_dataset(muz_path, hierarchy_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    print(f"[UPSELL] Saved {out_path} ({len(payload['records'])} records)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
