#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Runs the full daily dashboard refresh for the hybrid MUZ + RD Center dataset."""

import os
import subprocess
import sys
import time
from datetime import datetime


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(PROJECT_ROOT, "scripts")
LM_DIR = os.path.join(PROJECT_ROOT, "lm")

PROCESSED_DATA_PATH = os.path.join(PROJECT_ROOT, "data", "processed", "daily_dashboard_data.json")
OUT_HTML_PATH = os.path.join(PROJECT_ROOT, "dist", "udp_daily_dashboard.html")
MERGED_XLSX_PATH = os.path.join(LM_DIR, "Заявки+обращения_объединено.xlsx")

REQUIRED_FILES = [
    os.path.join("lm", "merge_zayavki_obrasheniya.py"),
    os.path.join("lm", "Список заявок.xlsx"),
    os.path.join("lm", "Список обращений.xlsx"),
    os.path.join("scripts", "prepare_dashboard_data_combined.py"),
    os.path.join("scripts", "prepare_dashboard_data_muz.py"),
    os.path.join("scripts", "prepare_dashboard_data_hybrid.py"),
    os.path.join("scripts", "prepare_upsell_data.py"),
    os.path.join("scripts", "build_offline_dashboard_combined.py"),
    os.path.join("scripts", "calc_combined_mapping_fields.py"),
    os.path.join("muz_to_combined_mapping.txt"),
    os.path.join("lm", "Заявки+обращения_объединено.xlsx"),
    os.path.join("data", "source", "muz.xlsx"),
    os.path.join("config", "employee_hierarchy.xlsx"),
    os.path.join("config", "combined_source_exclusions.json"),
    os.path.join("templates", "daily_dashboard_template.html"),
    os.path.join("vendor", "chartjs.umd.js"),
]


STEPS = [
    ("MERGE", os.path.join(LM_DIR, "merge_zayavki_obrasheniya.py"), MERGED_XLSX_PATH),
    ("DATA", os.path.join(SCRIPTS_DIR, "prepare_dashboard_data_hybrid.py"), PROCESSED_DATA_PATH),
    ("HTML", os.path.join(SCRIPTS_DIR, "build_offline_dashboard_combined.py"), OUT_HTML_PATH),
]


def fmt_size(path):
    size = os.path.getsize(path)
    if size >= 1024 * 1024:
        return f"{size / (1024 * 1024):.2f} MB"
    if size >= 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size} B"


def check_inputs():
    missing = []
    for name in REQUIRED_FILES:
        full = os.path.join(PROJECT_ROOT, name)
        if os.path.exists(full):
            continue
        if name == os.path.join("config", "employee_hierarchy.xlsx"):
            if os.path.exists(os.path.join(PROJECT_ROOT, "config", "employee_hierarchy.json")):
                continue
        missing.append(name)
    if missing:
        print("[UPDATE COMBINED] ERROR: missing required files:", flush=True)
        for name in missing:
            print(f"  - {name}", flush=True)
        return False
    return True


def run_step(label, script_path, output_path):
    started = time.time()
    script_name = os.path.relpath(script_path, PROJECT_ROOT)
    output_name = os.path.relpath(output_path, PROJECT_ROOT)

    print(flush=True)
    print(f"[UPDATE COMBINED] {label}: start {script_name}", flush=True)
    result = subprocess.run([sys.executable, "-u", script_path], cwd=PROJECT_ROOT)
    elapsed = time.time() - started

    if result.returncode != 0:
        print(f"[UPDATE COMBINED] {label}: FAILED, code={result.returncode}, elapsed={elapsed:.1f}s", flush=True)
        return False
    if not os.path.exists(output_path):
        print(f"[UPDATE COMBINED] {label}: FAILED, output not found: {output_path}", flush=True)
        return False

    mtime = datetime.fromtimestamp(os.path.getmtime(output_path)).strftime("%Y-%m-%d %H:%M:%S")
    print(
        f"[UPDATE COMBINED] {label}: OK, elapsed={elapsed:.1f}s, output={output_name}, size={fmt_size(output_path)}, modified={mtime}",
        flush=True,
    )
    return True


def main():
    total_started = time.time()
    print("[UPDATE COMBINED] Full dashboard update started", flush=True)
    print(f"[UPDATE COMBINED] Project root: {PROJECT_ROOT}", flush=True)
    print(f"[UPDATE COMBINED] Python: {sys.executable}", flush=True)

    if not check_inputs():
        return 1

    for label, script_path, output_path in STEPS:
        if not run_step(label, script_path, output_path):
            print(flush=True)
            print("[UPDATE COMBINED] Stopped because previous step failed.", flush=True)
            return 1

    elapsed = time.time() - total_started
    print(flush=True)
    print(f"[UPDATE COMBINED] Done. dist/udp_daily_dashboard.html is updated. Total elapsed={elapsed:.1f}s", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
