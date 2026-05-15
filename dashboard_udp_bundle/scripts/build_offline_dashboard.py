#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Builds the offline daily dashboard HTML.

Inputs:
- data/processed/daily_dashboard_data.json
- templates/daily_dashboard_template.html
- vendor/chartjs.umd.js

Output:
- dist/udp_daily_dashboard.html
"""

from __future__ import annotations

import datetime
import json
import os
import re


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(PROJECT_ROOT, 'data', 'processed', 'daily_dashboard_data.json')
TEMPLATE_PATH = os.path.join(PROJECT_ROOT, 'templates', 'daily_dashboard_template.html')
CHARTJS_PATH = os.path.join(PROJECT_ROOT, 'vendor', 'chartjs.umd.js')
OUT_PATH = os.path.join(PROJECT_ROOT, 'dist', 'udp_daily_dashboard.html')
RAW_JS_PATH = os.path.join(PROJECT_ROOT, 'dist', 'udp_daily_dashboard_raw.js')

os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)


def maybe_fix_mojibake_cp1251_utf8(s: str) -> str:
    """
    Fix common mojibake where UTF-8 bytes were decoded as CP1251 and then saved as UTF-8.

    Covers:
    - "Р…"/"С…" style sequences (Cyrillic mojibake)
    - "рџ…"/"вЂ…" style sequences (emoji/punctuation mojibake)
    """

    def try_fix(chunk: str) -> str:
        try:
            fixed = chunk.encode('cp1251').decode('utf-8')
        except Exception:
            return chunk

        # Accept if conversion yields Cyrillic and reduces typical markers.
        if re.search(r'[\u0400-\u04FF]', fixed):
            before = chunk.count('\u0420') + chunk.count('\u0421')  # 'Р' + 'С'
            after = fixed.count('\u0420') + fixed.count('\u0421')
            if after <= max(1, before // 3):
                return fixed

            if any(w in fixed for w in (
                'Ежеднев', 'дашборд', 'Ссылка', 'Входящий', 'Результат', 'Сотруд', 'Команд',
                'Руковод', 'Тимлид', 'Направ', 'Фильтр', 'Заяв', 'Подключ', 'Нагрузка',
            )):
                return fixed

        # Emoji/punctuation: accept if we got non-BMP (emoji) or common punctuation.
        if any(ord(ch) > 0xFFFF for ch in fixed):
            return fixed
        if any(sym in fixed for sym in ('—', '–', '•', '№', '₽')):
            return fixed

        return chunk

    # 1) Cyrillic mojibake runs (Р/С dense).
    s = re.sub(r'(?:[\u0420\u0421][^<>{}\n]{0,260}){2,}', lambda m: try_fix(m.group(0)), s)
    s = re.sub(r'[\u0420\u0421][\u0400-\u04FF0-9"\\-\\s,.%/()]{1,240}', lambda m: try_fix(m.group(0)), s)

    # 2) Emoji/punctuation mojibake tokens (e.g. "рџ”–", "вЂ”").
    s = re.sub(r'(?:\u0440\u045f.{1,8}|\u0432\u0402.{1,12}|\u0432\u201e.{0,3})', lambda m: try_fix(m.group(0)), s)

    return s


with open(DATA_PATH, 'r', encoding='utf-8') as f:
    data_raw = f.read()

# Split huge raw rows into a separate JS file to keep HTML size reasonable and loading faster.
raw_script = "window.DASH_RAW = null; window.DASH_RAW_SRC = 'udp_daily_dashboard_raw.js';"
try:
    data_obj = json.loads(data_raw)
    raw_obj = data_obj.pop('raw', None)
    if raw_obj is not None:
        with open(RAW_JS_PATH, 'w', encoding='utf-8') as f:
            f.write('window.DASH_RAW = ')
            json.dump(raw_obj, f, ensure_ascii=False, separators=(',', ':'))
            f.write(';')
        # Do not auto-load raw (can be ~100MB and slow the dashboard). UI loads on demand for exports.
        raw_script = "window.DASH_RAW = null; window.DASH_RAW_SRC = 'udp_daily_dashboard_raw.js';"
    data_raw = json.dumps(data_obj, ensure_ascii=False, separators=(',', ':'))
except Exception:
    # Fallback: keep the raw data inlined as-is if parsing/splitting fails.
    pass

with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
    tpl = f.read()

if os.path.exists(CHARTJS_PATH):
    with open(CHARTJS_PATH, 'r', encoding='utf-8') as f:
        chartjs_raw = f.read()
    # Safe embed: prevent closing an outer <script>.
    chartjs_safe = chartjs_raw.replace('</script>', '<\\/script>')
else:
    print(f"[HTML] WARN: {CHARTJS_PATH} not found. Charts will not work offline.")
    chartjs_safe = '/* chart.umd.js not found */'

# Safe embed JSON (no escaping of </script> needed, we do a generic </ -> <\/).
data_json = data_raw.replace('</', '<\\/')

# Fix mojibake in template + inlined JSON only (not in vendor bundle).
tpl = maybe_fix_mojibake_cp1251_utf8(tpl)
data_json = maybe_fix_mojibake_cp1251_utf8(data_json)

out = tpl.replace('/*__DATA__*/', data_json)
out = out.replace('/*__CHARTJS__*/', chartjs_safe)
out = out.replace('/*__RAW_SCRIPT__*/', raw_script)
out = out.replace('/*__BUILD_TIME__*/', datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))

with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(out)

print(f"[HTML] Built {OUT_PATH}, size={os.path.getsize(OUT_PATH):,} bytes")
