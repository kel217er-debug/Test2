import argparse
import re
from pathlib import Path

import pandas as pd


def _split_linked_numbers(value: object) -> list[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    text = str(value).strip()
    if not text:
        return []
    parts = re.split(r"[,\n;]+", text)
    return [p.strip() for p in parts if p and p.strip()]


def build_merged(
    requests_path: Path,
    incidents_path: Path,
    requests_sheet: str | int = 0,
    incidents_sheet: str | int = 0,
    request_key: str = "Номер заявки",
    incident_links_key: str = "Номер связанной заявки",
) -> pd.DataFrame:
    req = pd.read_excel(requests_path, sheet_name=requests_sheet, dtype=object)
    inc = pd.read_excel(incidents_path, sheet_name=incidents_sheet, dtype=object)

    if request_key not in req.columns:
        raise KeyError(f"В файле заявок нет столбца '{request_key}'.")
    if incident_links_key not in inc.columns:
        raise KeyError(f"В файле обращений нет столбца '{incident_links_key}'.")

    req = req.copy()
    inc = inc.copy()

    # Чтобы не путать одинаковые названия столбцов из двух файлов — префиксуем колонки обращений
    # (кроме колонки со связанными заявками, которая нужна для алгоритма).
    incident_prefix = "Обращение."
    inc = inc.rename(
        columns={
            c: (c if c == incident_links_key else f"{incident_prefix}{c}")
            for c in inc.columns
        }
    )

    req["_key_request"] = req[request_key].astype(str).str.strip()
    req.loc[req[request_key].isna(), "_key_request"] = pd.NA

    inc["_linked_list"] = inc[incident_links_key].apply(_split_linked_numbers)
    inc_exploded = inc.explode("_linked_list", ignore_index=True)
    inc_exploded[f"{incident_prefix}Связанная_заявка_одно_значение"] = inc_exploded[
        "_linked_list"
    ]
    inc_exploded["_key_request"] = inc_exploded["_linked_list"].astype(str).str.strip()
    inc_exploded.loc[inc_exploded["_linked_list"].isna(), "_key_request"] = pd.NA

    merged = req.merge(
        inc_exploded.drop(columns=["_linked_list"]),
        how="left",
        on="_key_request",
        indicator=True,
    )

    # Ниже добавляем строки обращений, которые не нашли соответствующую строку из заявок
    matched_keys = set(req["_key_request"].dropna().astype(str))
    inc_unmatched = inc_exploded[
        inc_exploded["_key_request"].isna()
        | ~inc_exploded["_key_request"].astype(str).isin(matched_keys)
    ].copy()

    # Делаем "пустые" колонки заявок, чтобы схему можно было concat'ить
    req_cols = [c for c in req.columns if c != "_key_request"]
    for c in req_cols:
        if c not in inc_unmatched.columns:
            inc_unmatched[c] = pd.NA
    # и наоборот: если в merged есть колонки обращений, которых нет в inc_unmatched — добавим
    for c in merged.columns:
        if c not in inc_unmatched.columns:
            inc_unmatched[c] = pd.NA

    inc_unmatched["_merge"] = "right_only"
    inc_unmatched = inc_unmatched[merged.columns]

    out = pd.concat([merged, inc_unmatched], ignore_index=True)
    out["Источник строки"] = out["_merge"].map(
        {"both": "заявка+обращение", "left_only": "заявка", "right_only": "обращение"}
    )

    # Убираем служебные колонки
    out = out.drop(columns=["_key_request", "_merge"])

    # Порядок колонок: сначала заявки, затем обращения, затем "Источник строки"
    req_cols_out = [c for c in req.columns if c != "_key_request"]
    inc_cols_out = [c for c in out.columns if c not in req_cols_out + ["Источник строки"]]
    out = out[req_cols_out + inc_cols_out + ["Источник строки"]]

    # Нумерация строк (заголовок не нумеруется — это просто отдельная колонка данных)
    out.insert(0, "№", range(1, len(out) + 1))

    return out


def main() -> int:
    default_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description=(
            "Объединяет 'Список заявок.xlsx' и 'Список обращений.xlsx' по номеру заявки, "
            "учитывая множественные значения в 'Номер связанной заявки'. "
            "В конец добавляет обращения, не найденные в заявках."
        )
    )
    parser.add_argument(
        "--requests",
        default=str(default_dir / "Список заявок.xlsx"),
        help="Путь к файлу 'Список заявок.xlsx'.",
    )
    parser.add_argument(
        "--incidents",
        default=str(default_dir / "Список обращений.xlsx"),
        help="Путь к файлу 'Список обращений.xlsx'.",
    )
    parser.add_argument(
        "--out",
        default=str(default_dir / "Заявки+обращения_объединено.xlsx"),
        help="Куда сохранить результат (.xlsx).",
    )
    args = parser.parse_args()

    requests_path = Path(args.requests)
    incidents_path = Path(args.incidents)
    if not requests_path.exists():
        raise FileNotFoundError(f"Не найден файл заявок: {requests_path}")
    if not incidents_path.exists():
        raise FileNotFoundError(f"Не найден файл обращений: {incidents_path}")

    df = build_merged(requests_path, incidents_path)
    out_path = Path(args.out)
    df.to_excel(out_path, index=False)
    print(f"OK: сохранено {len(df)} строк в {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
