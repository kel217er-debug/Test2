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
import datetime
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(PROJECT_ROOT, 'data', 'processed', 'daily_dashboard_data.json')
TEMPLATE_PATH = os.path.join(PROJECT_ROOT, 'templates', 'daily_dashboard_template.html')
CHARTJS_PATH = os.path.join(PROJECT_ROOT, 'vendor', 'chartjs.umd.js')
OUT_PATH = os.path.join(PROJECT_ROOT, 'dist', 'udp_daily_dashboard.html')

os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)

with open(DATA_PATH, 'r', encoding='utf-8') as f:
    data_raw = f.read()

with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
    tpl = f.read()

if os.path.exists(CHARTJS_PATH):
    with open(CHARTJS_PATH, 'r', encoding='utf-8') as f:
        chartjs_raw = f.read()
    # Безопасная встройка — предотвращаем закрытие внешнего <script>
    chartjs_safe = chartjs_raw.replace('</script>', '<\\/script>')
else:
    print(f"[HTML] WARN: {CHARTJS_PATH} not found. Charts will not work offline.")
    chartjs_safe = '/* chart.umd.js not found */'

# Безопасная встройка JSON (без экранирования </script>)
data_json = data_raw.replace('</', '<\\/')
out = tpl.replace('/*__DATA__*/', data_json)
out = out.replace('/*__CHARTJS__*/', chartjs_safe)
out = out.replace('/*__BUILD_TIME__*/', datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))

with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(out)

print(f"[HTML] Built {OUT_PATH}, size={os.path.getsize(OUT_PATH):,} bytes")
