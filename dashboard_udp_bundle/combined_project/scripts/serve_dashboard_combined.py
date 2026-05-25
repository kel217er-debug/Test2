from __future__ import annotations

import argparse
import os
import shutil
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from socketserver import ThreadingMixIn


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8766


def _qs_get(qs: dict, key: str, default: str = "") -> str:
    v = qs.get(key)
    if not v:
        return default
    if isinstance(v, list):
        return str(v[0])
    return str(v)


def _pick_hierarchy_path() -> str:
    xlsx = os.path.join(PROJECT_ROOT, "config", "employee_hierarchy.xlsx")
    if os.path.exists(xlsx):
        return xlsx
    json_path = os.path.join(PROJECT_ROOT, "config", "employee_hierarchy.json")
    if os.path.exists(json_path):
        return json_path
    return xlsx


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


class DashboardHandler(SimpleHTTPRequestHandler):
    """Serves static files and provides /api/export_rows (combined)."""

    input_path: str | None = None
    hierarchy_path: str | None = None

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PROJECT_ROOT, **kwargs)

    def _send_text(self, status: int, text: str, content_type: str = "text/plain; charset=utf-8") -> None:
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_xlsx_file(self, path: str, download_name: str) -> None:
        try:
            size = os.path.getsize(path)
            f = open(path, "rb")
        except FileNotFoundError:
            self._send_text(404, f"Not found: {path}")
            return
        except OSError as e:
            self._send_text(500, f"Failed to open file: {e!r}")
            return

        try:
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            quoted = urllib.parse.quote(download_name)
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quoted}")
            self.send_header("Content-Length", str(size))
            self.end_headers()
            shutil.copyfileobj(f, self.wfile)
        finally:
            try:
                f.close()
            except Exception:
                pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == "/api/health":
            self._send_text(200, "ok")
            return

        # Пока только статика. Экспорт строк можно добавить позже, если понадобится.
        return super().do_GET()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default=DEFAULT_HOST)
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument("--input", default="", help="Optional path to combined xlsx")
    ap.add_argument("--hierarchy", default="", help="Optional path to employee_hierarchy.xlsx/.json")
    args = ap.parse_args()

    os.chdir(PROJECT_ROOT)

    DashboardHandler.input_path = args.input.strip() or None
    DashboardHandler.hierarchy_path = args.hierarchy.strip() or None

    httpd = ThreadingHTTPServer((args.host, args.port), DashboardHandler)
    print(f"[SERVE COMBINED] Root: {PROJECT_ROOT}", flush=True)
    print(f"[SERVE COMBINED] http://{args.host}:{args.port}/dist/udp_daily_dashboard.html", flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[SERVE COMBINED] Stopped", flush=True)
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
