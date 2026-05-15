#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Экспорт строк, участвующих в расчёте конверсии.

Формирует 2 Excel-файла:
1) "обычная" конверсия (conv_regular): база = n_stayed_base
   -> строки, где заявка в базовых услугах и осталась в нашем канале.
2) "чистая" конверсия (conv_clean): база = n_base_all
   -> строки, где заявка в базовых услугах (независимо от того, осталась ли в канале).

Важно: логика фильтрации синхронизирована с `prepare_dashboard_data.py`:
- учитываем только строки нашего канала по `primary_podr`
- пропускаем строки без даты регистрации `reg_dt`
"""

import argparse
import os
import re
from datetime import datetime
from typing import Iterable, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from dashboard_logic.excel_filters import COL, is_base_service, is_connected, is_our_channel


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


def _norm_header(v: object) -> str:
    s = "" if v is None else str(v).strip().lower()
    s = s.replace("\u00a0", " ")
    s = _MULTISPACE_RE.sub(" ", s)
    return s


def _detect_header_row(rows_iter: Iterable[List[object]]) -> Tuple[List[object], Iterable[List[object]]]:
    """
    В некоторых выгрузках есть "преамбула". Логику определения строки заголовков
    повторяем из ETL: ищем первые колонки "МРФ" и "Регион".
    """
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
        # вернуть всё, что было после header_row: то есть buffered уже включает header_row,
        # поэтому пропускаем buffered целиком и продолжаем rows_iter.
        yield from rows_iter

    return header_row, tail_iter()


def _autosize(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            val = "" if cell.value is None else str(cell.value)
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)


def _write_xlsx(out_path: str, headers: List[object], rows: List[List[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Rows"

    ws.append([str(h) if h is not None else "" for h in headers] + ["in_base", "stayed", "connected"])
    for c in range(1, len(headers) + 4):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r in rows:
        ws.append(r)

    _autosize(ws)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


def export_conversion_rows(muz_path: str, out_dir: str) -> int:
    from python_calamine import CalamineWorkbook

    wb = CalamineWorkbook.from_path(muz_path)
    sheet = wb.get_sheet_by_name("Sheet1")
    rows_iter = iter(sheet.to_python())

    header_row, rows_iter = _detect_header_row(rows_iter)
    headers = header_row
    header_len = len(headers)

    regular_rows: List[List[object]] = []
    clean_rows: List[List[object]] = []

    n_read = 0
    for row in rows_iter:
        n_read += 1

        primary_podr = norm_str(row[COL["primary_podr"]] if COL["primary_podr"] < len(row) else None)
        if not is_our_channel(primary_podr):
            continue

        reg_dt = parse_dt(row[COL["reg_dt"]] if COL["reg_dt"] < len(row) else None)
        if reg_dt is None:
            continue

        service = norm_str(row[COL["service"]] if COL["service"] < len(row) else None).lower()
        current_podr = norm_str(row[COL["current_podr"]] if COL["current_podr"] < len(row) else None)
        final_status = norm_str(row[COL["final_status"]] if COL["final_status"] < len(row) else None)

        in_base = is_base_service(service)
        stayed = is_our_channel(current_podr)
        connected = is_connected(final_status)

        if not in_base:
            continue

        base = list(row[:header_len]) + ["да" if in_base else "нет", "да" if stayed else "нет", "да" if connected else "нет"]
        if stayed:
            regular_rows.append(base)
        clean_rows.append(base)

    out_regular = os.path.join(out_dir, "conversion_regular_rows.xlsx")
    out_clean = os.path.join(out_dir, "conversion_clean_rows.xlsx")
    _write_xlsx(out_regular, headers, regular_rows)
    _write_xlsx(out_clean, headers, clean_rows)

    print(f"Source: {os.path.abspath(muz_path)}")
    print(f"Read rows (after header): {n_read}")
    print(f"Regular rows (n_stayed_base): {len(regular_rows)} -> {os.path.abspath(out_regular)}")
    print(f"Clean rows (n_base_all): {len(clean_rows)} -> {os.path.abspath(out_clean)}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--muz", default=DEFAULT_MUZ, help="Путь к muz.xlsx")
    ap.add_argument("--out-dir", default=DEFAULT_OUT_DIR, help="Папка для результатов")
    args = ap.parse_args()
    return export_conversion_rows(args.muz, args.out_dir)


if __name__ == "__main__":
    raise SystemExit(main())

