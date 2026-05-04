#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Runs the full daily dashboard refresh: data preparation, then HTML build."""
import os
import subprocess
import sys
import time
from datetime import datetime


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(PROJECT_ROOT, 'scripts')
SOURCE_DATA_DIR = os.path.join(PROJECT_ROOT, 'data', 'source')
PROCESSED_DATA_DIR = os.path.join(PROJECT_ROOT, 'data', 'processed')
DIST_DIR = os.path.join(PROJECT_ROOT, 'dist')

REQUIRED_FILES = [
    os.path.join('scripts', 'prepare_dashboard_data.py'),
    os.path.join('scripts', 'build_offline_dashboard.py'),
    os.path.join('config', 'employee_hierarchy.json'),
    os.path.join('templates', 'daily_dashboard_template.html'),
    os.path.join('vendor', 'chartjs.umd.js'),
]

STEPS = [
    ('DATA', os.path.join(SCRIPTS_DIR, 'prepare_dashboard_data.py'), os.path.join(PROCESSED_DATA_DIR, 'daily_dashboard_data.json')),
    ('HTML', os.path.join(SCRIPTS_DIR, 'build_offline_dashboard.py'), os.path.join(DIST_DIR, 'udp_daily_dashboard.html')),
]


def find_muz_file():
    if not os.path.isdir(SOURCE_DATA_DIR):
        return None
    for name in os.listdir(SOURCE_DATA_DIR):
        lower_name = name.lower()
        if lower_name.endswith('.xlsx') and ('muz' in lower_name or 'муз' in lower_name):
            return os.path.join(SOURCE_DATA_DIR, name)
    return None


def fmt_size(path):
    size = os.path.getsize(path)
    if size >= 1024 * 1024:
        return f'{size / (1024 * 1024):.2f} MB'
    if size >= 1024:
        return f'{size / 1024:.1f} KB'
    return f'{size} B'


def check_inputs():
    missing = [name for name in REQUIRED_FILES if not os.path.exists(os.path.join(PROJECT_ROOT, name))]
    muz_path = find_muz_file()
    if not muz_path:
        missing.append(os.path.join('data', 'source', '*muz*.xlsx'))

    if missing:
        print('[UPDATE] ERROR: missing required files:', flush=True)
        for name in missing:
            print(f'  - {name}', flush=True)
        return False

    print(f'[UPDATE] Source Excel: {muz_path}', flush=True)
    return True


def run_step(label, script_path, output_path):
    started = time.time()
    script_name = os.path.relpath(script_path, PROJECT_ROOT)
    output_name = os.path.relpath(output_path, PROJECT_ROOT)

    print(flush=True)
    print(f'[UPDATE] {label}: start {script_name}', flush=True)
    result = subprocess.run([sys.executable, '-u', script_path], cwd=PROJECT_ROOT)
    elapsed = time.time() - started

    if result.returncode != 0:
        print(f'[UPDATE] {label}: FAILED, code={result.returncode}, elapsed={elapsed:.1f}s', flush=True)
        return False

    if not os.path.exists(output_path):
        print(f'[UPDATE] {label}: FAILED, output not found: {output_path}', flush=True)
        return False

    mtime = datetime.fromtimestamp(os.path.getmtime(output_path)).strftime('%Y-%m-%d %H:%M:%S')
    print(f'[UPDATE] {label}: OK, elapsed={elapsed:.1f}s, output={output_name}, size={fmt_size(output_path)}, modified={mtime}', flush=True)
    return True


def main():
    total_started = time.time()
    print('[UPDATE] Full dashboard update started', flush=True)
    print(f'[UPDATE] Project root: {PROJECT_ROOT}', flush=True)
    print(f'[UPDATE] Python: {sys.executable}', flush=True)

    if not check_inputs():
        return 1

    for label, script_name, output_name in STEPS:
        if not run_step(label, script_name, output_name):
            print(flush=True)
            print('[UPDATE] Stopped because previous step failed.', flush=True)
            return 1

    elapsed = time.time() - total_started
    print(flush=True)
    print(f'[UPDATE] Done. dist/udp_daily_dashboard.html is updated. Total elapsed={elapsed:.1f}s', flush=True)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
