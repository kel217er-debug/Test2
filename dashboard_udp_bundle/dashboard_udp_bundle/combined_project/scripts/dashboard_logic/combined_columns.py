"""Canonical column names and alias resolution for the combined Excel export."""

from __future__ import annotations

import re
from collections import defaultdict

import pandas as pd


COMBINED_COLUMN_ALIASES = {
    "Текущий канал обработки заявки": ("Канал обработки",),
    "ФИО текущего исполнителя заявки": ("ФИО исполнителя заявки",),
    "Email текущего исполнителя заявки": ("Email исполнителя заявки",),
}

COMBINED_MAPPING_ALIASES = {
    "Предлагаемая услуга (где итоговая услуга)": "Предлагаемая услуга",
    "Дата, время обновления статуса заявки (если статус итоговый)": "Дата, время обновления статуса заявки",
    'ФИО исполнителя заявки (если Статус="Услуга подключена")': "ФИО текущего исполнителя заявки",
    "Канал обработки": "Текущий канал обработки заявки",
    "ФИО исполнителя заявки": "ФИО текущего исполнителя заявки",
    "Email исполнителя заявки": "Email текущего исполнителя заявки",
}


def normalize_combined_header(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").replace("\n", " ").strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("ё", "е")
    return text


def resolve_combined_mapping_column(value: str | None) -> str | None:
    if not value:
        return None
    return COMBINED_MAPPING_ALIASES.get(value, value)


def canonicalize_combined_dataframe(
    df: pd.DataFrame,
    required_columns: list[str] | tuple[str, ...] | set[str] | None = None,
) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    """Ensure canonical columns exist and are filled from aliases when needed."""
    required = list(required_columns or [])
    alias_map: dict[str, list[str]] = defaultdict(list)
    for canonical, aliases in COMBINED_COLUMN_ALIASES.items():
        alias_map[canonical].append(canonical)
        alias_map[canonical].extend(list(aliases))
    for name in required:
        alias_map[name].append(name)

    normalized_actual: dict[str, list[str]] = defaultdict(list)
    for col in df.columns:
        normalized_actual[normalize_combined_header(col)].append(col)

    out = df.copy()
    resolved: dict[str, list[str]] = {}
    for canonical, candidates in alias_map.items():
        matches: list[str] = []
        seen = set()
        for candidate in candidates:
            for actual in normalized_actual.get(normalize_combined_header(candidate), []):
                if actual not in seen:
                    matches.append(actual)
                    seen.add(actual)
        if not matches:
            continue

        if canonical in out.columns:
            combined = out[canonical].copy()
        else:
            combined = pd.Series([pd.NA] * len(out), index=out.index, dtype="object")

        for actual in matches:
            source = out[actual]
            source_stripped = source.astype("object").map(lambda v: "" if pd.isna(v) else str(v).strip())
            needs_fill = combined.isna() | (combined.astype("object").map(lambda v: "" if pd.isna(v) else str(v).strip()) == "")
            combined = combined.where(~needs_fill, source.where(source_stripped != "", other=pd.NA))

        out[canonical] = combined
        resolved[canonical] = matches

    return out, resolved
