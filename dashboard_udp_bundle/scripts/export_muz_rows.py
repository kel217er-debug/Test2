#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export source MUZ rows used in dashboard calculations to XLSX.

This script is intended to replace in-browser raw export (no huge embedded raw JSON).
The HTML dashboard can copy a ready-to-run command with current filters/period/tab.

Output:
  dist/muz_rows_<tab>_<period>_<date>.xlsx
"""

from __future__ import annotations

import argparse
import os
import re
from datetime import datetime
from typing import Iterable, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

import json

from dashboard_logic.excel_filters import COL, direction_for_channel, is_connected, is_our_channel


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_MUZ = os.path.join(PROJECT_ROOT, "data", "source", "muz.xlsx")
DEFAULT_OUT_DIR = os.path.join(PROJECT_ROOT, "dist")

_DT_RE = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})\s*$")
_MULTISPACE_RE = re.compile(r"\s+")


def norm_str(v: object) -> str:
    s = "" if v is None else str(v).strip()
    return s.replace("\u00a0", " ").strip()


def parse_dt(v: object) -> Optional[datetime]:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        m = _DT_RE.match(s)
        if m:
            d, mo, y, H, M, S = map(int, m.groups())
            try:
                return datetime(y, mo, d, H, M, S)
            except ValueError:
                return None
    return None


def iso_week_key(dt: datetime) -> str:
    y, w, _ = dt.isocalendar()
    return f"{y}-W{w:02d}"


def month_key(dt: datetime) -> str:
    return dt.strftime("%Y-%m")


def _norm_header(v: object) -> str:
    s = "" if v is None else str(v).strip().lower()
    s = s.replace("\u00a0", " ")
    s = _MULTISPACE_RE.sub(" ", s)
    return s


def _detect_header_row(rows_iter: Iterable[List[object]]) -> Tuple[List[object], Iterable[List[object]]]:
    buffered: List[List[object]] = []
    header_row: Optional[List[object]] = None
    for _ in range(50):
        row = next(rows_iter)
        buffered.append(row)
        c0 = norm_str(row[0]).lower() if len(row) > 0 else ""
        c1 = norm_str(row[1]).lower() if len(row) > 1 else ""
        if c0 == "мрф" and c1 == "регион":
            header_row = row
            break
    if header_row is None:
        header_row = next(rows_iter)
        buffered.append(header_row)

    def tail_iter():
        yield from rows_iter

    return header_row, tail_iter()

def clean_name(v: object) -> str:
    s = norm_str(v)
    if not s:
        return ""

    if "@" in s:
        parts = s.split()
        clean = []
        for p in parts:
            if "@" in p:
                break
            clean.append(p)
        s = " ".join(clean)

    tokens = re.findall(r"[A-Za-zА-Яа-яЁёЀ-ӿ-]+", s)
    cyr = [t for t in tokens if re.search(r"[А-Яа-яЁёЀ-ӿ]", t)]
    if len(cyr) >= 2:
        return " ".join(cyr[:3]).strip()
    return " ".join(tokens).strip()


def _norm_header_h(v: object) -> str:
    return norm_str(v).strip().lower().replace("\n", " ")


def load_hierarchy(hierarchy_path: str) -> dict:
    if hierarchy_path.lower().endswith(".json"):
        with open(hierarchy_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        employees = raw.get("employees", raw if isinstance(raw, list) else [])
        out = {}
        for info in employees:
            name = clean_name(info.get("name") or info.get("manager") or "")
            if not name:
                continue
            out[name] = {
                "teamlead": info.get("teamlead") or "",
                "director": info.get("director") or "",
                "direction": info.get("direction") or "",
                "mrf": info.get("mrf") or "",
            }
        return out

    from python_calamine import CalamineWorkbook

    wb = CalamineWorkbook.from_path(hierarchy_path)
    sheet_names = wb.sheet_names
    if not sheet_names:
        return {}

    preferred = None
    for n in sheet_names:
        ln = n.strip().lower()
        if ln in ("employees", "employee_hierarchy", "hierarchy", "справочник", "сотрудники", "иерархия"):
            preferred = n
            break
    sheet = wb.get_sheet_by_name(preferred or sheet_names[0])
    rows = list(sheet.to_python())
    if not rows:
        return {}

    headers = [_norm_header_h(v) for v in rows[0]]

    def col_index(*names):
        for name in names:
            if name in headers:
                return headers.index(name)
        return None

    idx_name = col_index("name", "employee", "fio", "фио", "сотрудник", "фамилия имя отчество")
    idx_teamlead = col_index("teamlead", "tl", "тимлид", "тим лид", "тим-лид", "тимлидер", "тим лидер")
    idx_director = col_index("director", "manager", "руководитель", "директор")
    idx_direction = col_index("direction", "направление")
    idx_mrf = col_index("mrf", "мрф")

    if idx_name is None:
        raise ValueError('Hierarchy Excel: required column "name" (or "ФИО") not found in header row')

    out = {}
    for row in rows[1:]:
        name = clean_name(row[idx_name] if idx_name < len(row) else "")
        if not name:
            continue
        out[name] = {
            "teamlead": clean_name(row[idx_teamlead]) if idx_teamlead is not None and idx_teamlead < len(row) else "",
            "director": clean_name(row[idx_director]) if idx_director is not None and idx_director < len(row) else "",
            "direction": norm_str(row[idx_direction]) if idx_direction is not None and idx_direction < len(row) else "",
            "mrf": norm_str(row[idx_mrf]) if idx_mrf is not None and idx_mrf < len(row) else "",
        }
    return out


def _match_filters(direction: str, mrf: str, director: str, teamlead: str,
                   row_direction: str, row_mrf: str, row_director: str, row_teamlead: str) -> bool:
    if direction and row_direction != direction:
        return False
    if mrf and row_mrf != mrf:
        return False
    if director and row_director != director:
        return False
    if teamlead and row_teamlead != teamlead:
        return False
    return True


def _autosize(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)


def _write_xlsx(out_path: str, headers: List[object], rows: List[List[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "MUZ"

    ws.append([str(h) if h is not None else "" for h in headers])
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r in rows:
        ws.append(r)

    _autosize(ws)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


def export_rows(muz_path: str, hierarchy_path: str, out_dir: str, tab: str, period_kind: str, period_key: str,
                direction: str, mrf: str, director: str, teamlead: str) -> str:
    from python_calamine import CalamineWorkbook

    emp_map = load_hierarchy(hierarchy_path)

    wb = CalamineWorkbook.from_path(muz_path)
    sheet = wb.get_sheet_by_name("Sheet1")
    rows_iter = iter(sheet.to_python())
    header_row, rows_iter = _detect_header_row(rows_iter)
    headers = header_row
    header_len = len(headers)

    out_rows: List[List[object]] = []

    for row in rows_iter:
        podr_prim = norm_str(row[COL["primary_podr"]] if COL["primary_podr"] < len(row) else None)
        if not is_our_channel(podr_prim):
            continue

        reg_dt = parse_dt(row[COL["reg_dt"]] if COL["reg_dt"] < len(row) else None)
        if reg_dt is None:
            continue

        final_status = norm_str(row[COL["final_status"]] if COL["final_status"] < len(row) else None)
        has_final = bool(final_status)
        final_dt = parse_dt(row[COL["final_dt"]] if has_final and COL["final_dt"] < len(row) else None)

        reg_w = iso_week_key(reg_dt)
        reg_m = month_key(reg_dt)
        cl_w = iso_week_key(final_dt) if final_dt else None
        cl_m = month_key(final_dt) if final_dt else None

        in_period_reg = (reg_m == period_key) if period_kind == "month" else (reg_w == period_key)
        in_period_cl = (cl_m == period_key) if period_kind == "month" else (cl_w == period_key)

        primary_name = clean_name(row[COL["primary_name"]] if COL["primary_name"] < len(row) else "")
        exec_name = clean_name(row[COL["exec_name"]] if COL["exec_name"] < len(row) else "")

        prim_info = emp_map.get(primary_name, {}) if primary_name else {}
        exec_info = emp_map.get(exec_name, {}) if exec_name else {}
        prim_dir = prim_info.get("direction") or direction_for_channel(podr_prim)
        prim_mrf = prim_info.get("mrf") or ""
        prim_dre = prim_info.get("director") or ""
        prim_tl = prim_info.get("teamlead") or ""

        exec_dir = exec_info.get("direction") or direction_for_channel(exec_info.get("podr") or "") or prim_dir
        exec_mrf = exec_info.get("mrf") or ""
        exec_dre = exec_info.get("director") or ""
        exec_tl = exec_info.get("teamlead") or ""

        if tab in ("teams", "employees"):
            if not (in_period_reg or in_period_cl):
                continue
            # Matches dashboard export semantics: for registered-period rows we allow either primary dims or exec dims
            # (because conversions use exec axis). For closed-period rows we require primary dims.
            reg_ok = False
            if in_period_reg:
                reg_ok = (
                    _match_filters(direction, mrf, director, teamlead, prim_dir, prim_mrf, prim_dre, prim_tl)
                    or _match_filters(direction, mrf, director, teamlead, exec_dir, exec_mrf, exec_dre, exec_tl)
                )
            cl_ok = in_period_cl and _match_filters(direction, mrf, director, teamlead, prim_dir, prim_mrf, prim_dre, prim_tl)
            if not (reg_ok or cl_ok):
                continue
        elif tab == "services":
            if not (has_final and is_connected(final_status) and in_period_cl):
                continue
            if not _match_filters(direction, mrf, director, teamlead, prim_dir, prim_mrf, prim_dre, prim_tl):
                continue
        elif tab == "open":
            if has_final:
                continue
            if not in_period_reg:
                continue
            if not _match_filters(direction, mrf, director, teamlead, prim_dir, prim_mrf, prim_dre, prim_tl):
                continue
        else:
            # unknown tab
            continue

        out_rows.append(list(row[:header_len]))

    stamp = datetime.now().strftime("%Y-%m-%d")
    out_name = f"muz_rows_{tab}_{period_key}_{stamp}.xlsx"
    out_path = os.path.join(out_dir, out_name)
    _write_xlsx(out_path, headers, out_rows)
    return out_path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--muz", default=DEFAULT_MUZ, help="Путь к muz.xlsx")
    ap.add_argument("--out-dir", default=DEFAULT_OUT_DIR, help="Папка для результата")
    ap.add_argument("--hierarchy", default=os.path.join(PROJECT_ROOT, "config", "employee_hierarchy.xlsx"),
                    help="Путь к employee_hierarchy.xlsx или .json")
    ap.add_argument("--tab", required=True, choices=["teams", "employees", "open", "services"], help="Вкладка")
    ap.add_argument("--period-kind", required=True, choices=["week", "month"], help="Тип периода")
    ap.add_argument("--period", required=True, help="Ключ периода (например 2026-W20 или 2026-05)")
    ap.add_argument("--dir", default="", help="Фильтр: направление")
    ap.add_argument("--mrf", default="", help="Фильтр: МРФ")
    ap.add_argument("--director", default="", help="Фильтр: руководитель")
    ap.add_argument("--teamlead", default="", help="Фильтр: тимлид")
    args = ap.parse_args()

    out = export_rows(
        muz_path=args.muz,
        hierarchy_path=args.hierarchy,
        out_dir=args.out_dir,
        tab=args.tab,
        period_kind=args.period_kind,
        period_key=args.period,
        direction=args.dir,
        mrf=args.mrf,
        director=args.director,
        teamlead=args.teamlead,
    )
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
