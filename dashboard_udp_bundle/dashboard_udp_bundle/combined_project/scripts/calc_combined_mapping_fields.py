#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Считает/собирает показатели из `lm/Заявки+обращения_объединено.xlsx`
по маппингу из `muz_to_combined_mapping.txt`.

Важно: исходный Excel не изменяет — только читает и сохраняет отдельный файл
с вычисленными полями.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from dashboard_logic.combined_columns import (
    canonicalize_combined_dataframe,
    normalize_combined_header,
    resolve_combined_mapping_column,
)
from dashboard_logic.excel_filters_combined import COL


@dataclass(frozen=True)
class MapRow:
    key: str
    src1: str | None
    src2: str | None


FINAL_STATUSES = {"Услуга подключена", "Услуга не подключена"}
CONNECTED_STATUS = "Услуга подключена"
COMPUTED_MAPPING_KEYS = {"connection_result", "connector_name", "final_status", "final_dt", "hrs_since_reg"}
OPTIONAL_MAPPING_KEYS = {"sla_acc", "transfer_dt"}


def _norm(s: str) -> str:
    return " ".join(str(s).strip().split())

def _is_missing(v: object) -> bool:
    try:
        return pd.isna(v)
    except Exception:
        return v is None


def _read_mapping(mapping_path: Path) -> list[MapRow]:
    def _clean_src(v: str | None) -> str | None:
        if not v:
            return None
        v = v.strip()
        if not v:
            return None
        low = v.lower()
        # В маппинге иногда в колонках лежат формулы/пояснения, а не реальные названия столбцов.
        if low.startswith("(") or "формула" in low or low.startswith("если(") or "статус=" in low:
            return None
        return resolve_combined_mapping_column(v)

    rows: list[MapRow] = []
    for raw_line in mapping_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or "|" not in line:
            continue
        if line.startswith("Переменная |") or line.startswith("---------- |"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 4:
            continue
        key = parts[0]
        if not key or " " in key:
            continue
        src1 = _clean_src(parts[2])
        src2 = _clean_src(parts[3])
        rows.append(MapRow(key=key, src1=src1, src2=src2))
    if not rows:
        raise ValueError(f"Не удалось распарсить маппинг из {mapping_path}")
    return rows


def _pick_first_non_null(df: pd.DataFrame, candidates: list[str]) -> pd.Series:
    out = pd.Series([pd.NA] * len(df), index=df.index, dtype="object")
    for c in candidates:
        if not c or c not in df.columns:
            continue
        s = df[c]
        mask = out.isna() & s.notna() & (s.astype(str).str.strip() != "")
        out.loc[mask] = s.loc[mask]
    return out


def _required_source_columns(mapping_rows: list[MapRow]) -> list[str]:
    required_keys = set(COL) - COMPUTED_MAPPING_KEYS - OPTIONAL_MAPPING_KEYS
    needed = []
    for row in mapping_rows:
        if row.key not in required_keys:
            continue
        for src in (row.src1, row.src2):
            if src:
                needed.append(src)
    return sorted(set(needed))


def _validate_required_mapping_sources(df: pd.DataFrame, mapping_rows: list[MapRow]) -> None:
    present = {normalize_combined_header(c) for c in df.columns}
    errors: list[str] = []
    required_keys = set(COL) - COMPUTED_MAPPING_KEYS - OPTIONAL_MAPPING_KEYS
    for row in mapping_rows:
        if row.key not in required_keys:
            continue
        candidates = [src for src in (row.src1, row.src2) if src]
        if not candidates:
            continue
        if any(normalize_combined_header(src) in present for src in candidates):
            continue
        errors.append(f"{row.key}: ни одна из колонок не найдена ({', '.join(candidates)})")
    if errors:
        joined = "\n".join(f"  - {item}" for item in sorted(errors))
        raise ValueError(
            "Во входном Excel отсутствуют обязательные колонки для combined-расчёта:\n"
            f"{joined}"
        )


def _to_dt(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype="datetime64[ns]")
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    # Часто встречается формат "dd.mm.yyyy HH:MM МСК" (или похожие суффиксы таймзоны).
    cleaned = series.astype("object").copy()
    s = cleaned.astype(str)
    s = s.str.replace("\u00a0", " ", regex=False)
    # Drop trailing timezone/name token after a space (e.g. "МСК", "MSK", or mojibake like "ÌÑÊ").
    s = s.str.replace(r"\s+[^0-9:]+$", "", regex=True)
    s = s.str.strip()
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def _to_num(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    s = series.astype(str).str.replace("\xa0", "", regex=False).str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")


def build_fields(df: pd.DataFrame, mapping_rows: list[MapRow], now_dt: datetime) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)

    # First pass: simple "take from column A else from column B"
    for r in mapping_rows:
        candidates = [c for c in [r.src1, r.src2] if c]
        if candidates:
            out[r.key] = _pick_first_non_null(df, candidates)
        else:
            out[r.key] = pd.NA

    # Second pass: known computed fields (per mapping text)
    if "current_status" in out.columns:
        status = out["current_status"].fillna("").astype(str).str.strip()
    elif "Статус" in df.columns:
        status = df["Статус"].fillna("").astype(str).str.strip()
        out["current_status"] = status
    else:
        status = pd.Series([""] * len(df), index=df.index, dtype="object")

    if "connection_result" in out.columns:
        out["connection_result"] = status.map(lambda s: "да" if s == CONNECTED_STATUS else "")

    if "connector_name" in out.columns:
        # По маппингу: "ФИО исполнителя заявки (если Статус='Услуга подключена')" иначе пусто.
        fio_req = df["ФИО текущего исполнителя заявки"] if "ФИО текущего исполнителя заявки" in df.columns else None
        fio_inc = df["Обращение.ФИО исполнителя"] if "Обращение.ФИО исполнителя" in df.columns else None
        fio = pd.Series([pd.NA] * len(df), index=df.index, dtype="object")
        if fio_req is not None:
            fio = fio_req
        elif fio_inc is not None:
            fio = fio_inc
        out["connector_name"] = fio.where(status == CONNECTED_STATUS, other=pd.NA)

    if "final_status" in out.columns:
        out["final_status"] = status.where(status.isin(FINAL_STATUSES), other=pd.NA)

    if "final_dt" in out.columns:
        # По маппингу: "Дата, время обновления статуса заявки (если статус итоговый)"
        if "Дата, время обновления статуса заявки" in df.columns:
            upd = _to_dt(df["Дата, время обновления статуса заявки"])
            out["final_dt"] = upd.where(status.isin(FINAL_STATUSES), other=pd.NaT)
        else:
            out["final_dt"] = pd.NaT

    if "hrs_since_reg" in out.columns:
        # В объединённом файле этого поля нет — считаем как now - reg_dt (на момент запуска).
        reg_dt = _to_dt(out.get("reg_dt"))
        delta = (pd.Timestamp(now_dt) - reg_dt)
        out["hrs_since_reg"] = (delta.dt.total_seconds() / 3600.0).round(2)

    # Convert known datetime-like fields to datetime64 where possible (helps downstream ETL).
    for k in ("reg_dt", "accept_dt", "transfer_dt", "final_dt"):
        if k in out.columns:
            out[k] = _to_dt(out[k])

    if "equip_price" in out.columns:
        out["equip_price"] = _to_num(out["equip_price"])
    if "install_amount" in out.columns:
        out["install_amount"] = _to_num(out["install_amount"])
    if "monthly_amount" in out.columns:
        out["monthly_amount"] = _to_num(out["monthly_amount"])
    if "connected_services_cnt" in out.columns:
        out["connected_services_cnt"] = _to_num(out["connected_services_cnt"])

    # Normalize some strings (optional but helps downstream)
    for k in ("mrf", "region", "channel", "service", "inn", "segment", "exec_name", "exec_podr",
              "primary_name", "primary_podr", "final_reason", "current_exec", "current_podr"):
        if k in out.columns:
            out[k] = out[k].map(lambda v: pd.NA if _is_missing(v) else _norm(v))

    return out


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Собирает/вычисляет поля из 'Заявки+обращения_объединено.xlsx' по маппингу "
            "из 'muz_to_combined_mapping.txt' и сохраняет отдельный файл."
        )
    )
    parser.add_argument(
        "--input",
        default=str(Path("lm") / "Заявки+обращения_объединено.xlsx"),
        help="Путь к объединённому Excel (.xlsx).",
    )
    parser.add_argument(
        "--mapping",
        default="muz_to_combined_mapping.txt",
        help="Путь к текстовому маппингу (таблица с '|').",
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Лист Excel (имя или индекс). По умолчанию: 0 (первый лист).",
    )
    parser.add_argument(
        "--out",
        default=str(Path("lm") / "computed_muz_fields.xlsx"),
        help="Куда сохранить результат (.xlsx).",
    )
    parser.add_argument(
        "--keep-id",
        action="store_true",
        help="Сохранить идентификаторы из исходного файла: '№', 'Номер заявки', 'Номер обращения'.",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    mapping_path = Path(args.mapping)
    out_path = Path(args.out)

    if not input_path.exists():
        raise FileNotFoundError(f"Не найден входной файл: {input_path}")
    if not mapping_path.exists():
        raise FileNotFoundError(f"Не найден файл маппинга: {mapping_path}")

    mapping_rows = _read_mapping(mapping_path)
    df_raw = pd.read_excel(input_path, sheet_name=args.sheet, dtype=object)
    required_columns = _required_source_columns(mapping_rows)
    df, resolved_columns = canonicalize_combined_dataframe(df_raw, required_columns=required_columns)
    _validate_required_mapping_sources(df, mapping_rows)

    now_dt = datetime.now()
    fields = build_fields(df, mapping_rows, now_dt=now_dt)

    if args.keep_id:
        id_cols = [c for c in ["№", "Номер заявки", "Номер обращения"] if c in df.columns]
        if id_cols:
            fields = pd.concat([df[id_cols], fields], axis=1)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        fields.to_excel(out_path, index=False, na_rep="")
    except PermissionError:
        stamped = out_path.with_name(
            f"{out_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{out_path.suffix}"
        )
        fields.to_excel(stamped, index=False, na_rep="")
        out_path = stamped

    present_norm = {normalize_combined_header(c) for c in df.columns}
    missing_src = sorted(
        {
            r.src1
            for r in mapping_rows
            if r.src1 and normalize_combined_header(r.src1) not in present_norm
        }
        | {
            r.src2
            for r in mapping_rows
            if r.src2 and normalize_combined_header(r.src2) not in present_norm
        }
    )
    if missing_src:
        print("WARN: не найдены колонки (упомянуты в маппинге, но отсутствуют во входном Excel):")
        for c in missing_src:
            print(f"  - {c}")
    renamed_aliases = {k: v for k, v in resolved_columns.items() if v and v != [k]}
    if renamed_aliases:
        print("INFO: использованы алиасы колонок combined-выгрузки:")
        for canonical, actual in sorted(renamed_aliases.items()):
            print(f"  - {canonical}: {', '.join(actual)}")

    # SLA принятия (sla_acc) и дата передачи (transfer_dt) в маппинге описаны как формулы/пусто.
    # Если их нет в исходном Excel, они останутся пустыми.
    print(f"OK: сохранено {len(fields)} строк x {len(fields.columns)} колонок в {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

