from __future__ import annotations

import argparse
import os
import shutil
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from socketserver import ThreadingMixIn


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_HOST = '127.0.0.1'
DEFAULT_PORT = 8765


def _qs_get(qs: dict, key: str, default: str = '') -> str:
    v = qs.get(key)
    if not v:
        return default
    if isinstance(v, list):
        return str(v[0])
    return str(v)


def _pick_hierarchy_path() -> str:
    xlsx = os.path.join(PROJECT_ROOT, 'config', 'employee_hierarchy.xlsx')
    if os.path.exists(xlsx):
        return xlsx

    json_path = os.path.join(PROJECT_ROOT, 'config', 'employee_hierarchy.json')
    if os.path.exists(json_path):
        return json_path

    cfg_dir = os.path.join(PROJECT_ROOT, 'config')
    if os.path.isdir(cfg_dir):
        for name in os.listdir(cfg_dir):
            low = name.lower()
            if low.startswith('employee_hierarchy') and low.endswith('.json'):
                return os.path.join(cfg_dir, name)

    return xlsx


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


class DashboardHandler(SimpleHTTPRequestHandler):
    """Serves static files and provides /api/export_muz_rows endpoint."""

    muz_path: str | None = None
    hierarchy_path: str | None = None

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PROJECT_ROOT, **kwargs)

    def _send_text(self, status: int, text: str, content_type: str = 'text/plain; charset=utf-8') -> None:
        data = text.encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_xlsx_file(self, path: str, download_name: str) -> None:
        try:
            size = os.path.getsize(path)
            f = open(path, 'rb')
        except FileNotFoundError:
            self._send_text(404, f'Not found: {path}')
            return
        except OSError as e:
            self._send_text(500, f'Failed to open file: {e!r}')
            return

        try:
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            quoted = urllib.parse.quote(download_name)
            self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{quoted}")
            self.send_header('Content-Length', str(size))
            self.end_headers()
            shutil.copyfileobj(f, self.wfile)
        finally:
            try:
                f.close()
            except Exception:
                pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/api/health':
            self._send_text(200, 'ok')
            return

        if parsed.path == '/api/export_muz_rows':
            qs = urllib.parse.parse_qs(parsed.query or '')
            tab = _qs_get(qs, 'tab')
            period_kind = _qs_get(qs, 'period_kind')
            period = _qs_get(qs, 'period')

            direction = _qs_get(qs, 'dir', '')
            mrf = _qs_get(qs, 'mrf', '')
            director = _qs_get(qs, 'director', '')
            teamlead = _qs_get(qs, 'teamlead', '')

            try:
                import export_muz_rows  # type: ignore
            except Exception as e:
                self._send_text(500, f'Failed to import export_muz_rows: {e!r}')
                return

            if tab not in ('teams', 'employees', 'open', 'services'):
                self._send_text(400, 'Bad request: tab must be one of teams/employees/open/services')
                return
            if period_kind not in ('week', 'month'):
                self._send_text(400, 'Bad request: period_kind must be week or month')
                return
            if not period:
                self._send_text(400, 'Bad request: period is required')
                return

            hierarchy_path = self.hierarchy_path or _pick_hierarchy_path()
            muz_path = self.muz_path or export_muz_rows.DEFAULT_MUZ

            try:
                out_path = export_muz_rows.export_rows(
                    muz_path=muz_path,
                    hierarchy_path=hierarchy_path,
                    out_dir=export_muz_rows.DEFAULT_OUT_DIR,
                    tab=tab,
                    period_kind=period_kind,
                    period_key=period,
                    direction=direction,
                    mrf=mrf,
                    director=director,
                    teamlead=teamlead,
                )
            except Exception as e:
                self._send_text(500, f'Export failed: {e!r}')
                return

            self._send_xlsx_file(out_path, os.path.basename(out_path))
            return

        return super().do_GET()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--host', default=DEFAULT_HOST)
    ap.add_argument('--port', type=int, default=DEFAULT_PORT)
    ap.add_argument('--muz', default='', help='Optional path to muz.xlsx')
    ap.add_argument('--hierarchy', default='', help='Optional path to employee_hierarchy.xlsx/.json')
    args = ap.parse_args()

    os.chdir(PROJECT_ROOT)

    DashboardHandler.muz_path = args.muz.strip() or None
    DashboardHandler.hierarchy_path = args.hierarchy.strip() or None

    httpd = ThreadingHTTPServer((args.host, args.port), DashboardHandler)
    print(f'[SERVE] Root: {PROJECT_ROOT}', flush=True)
    print(f'[SERVE] http://{args.host}:{args.port}/dist/udp_daily_dashboard.html', flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print('\n[SERVE] Stopped', flush=True)
        return 0


if __name__ == '__main__':
    raise SystemExit(main())
