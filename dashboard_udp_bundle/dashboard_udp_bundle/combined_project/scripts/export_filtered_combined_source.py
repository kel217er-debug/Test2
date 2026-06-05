#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export combined source rows after applying exclusion registry."""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd

from prepare_dashboard_data_combined import _is_excluded_combined_row, _load_excluded_combined_rows


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    src_path = root / "lm" / "Заявки+обращения_объединено.xlsx"
    exclusions_path = root / "Del" / "исключаемые заявки и обращения.xlsx"
    out_path = root / "lm" / "Заявки+обращения_объединено_после_исключений.xlsx"

    if not src_path.exists():
        raise FileNotFoundError(src_path)
    if not exclusions_path.exists():
        raise FileNotFoundError(exclusions_path)

    df = pd.read_excel(src_path, sheet_name=0, dtype=object)
    excluded_keys = _load_excluded_combined_rows(str(exclusions_path))
    excluded_mask = df.apply(lambda row: _is_excluded_combined_row(row, excluded_keys), axis=1)
    filtered_df = df.loc[~excluded_mask].reset_index(drop=True)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    filtered_df.to_excel(out_path, index=False)

    print(
        f"[EXPORT FILTERED COMBINED] saved={out_path} "
        f"source_rows={len(df)} excluded_rows={int(excluded_mask.sum())} kept_rows={len(filtered_df)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
