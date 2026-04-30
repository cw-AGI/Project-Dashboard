#!/usr/bin/env python3
"""
Local dashboard server with save endpoint.

Usage:
  python3 scripts/dashboard_server.py --port 8765

Then open:
  http://localhost:8765/WN_Dashboard_v3.6.html
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PROJECTS_DIR = ROOT / "projects"
HTML_FILE = ROOT / "WN_Dashboard_v3.6.html"


def _json_response(handler: SimpleHTTPRequestHandler, status: int, payload: dict) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def _safe_project_id(pid: str) -> str:
    pid = str(pid or "").strip()
    if not pid:
        raise ValueError("projectId is required")
    if "/" in pid or "\\" in pid or ".." in pid:
        raise ValueError("invalid projectId path")
    return pid


def _merge_index_entries(existing: list, incoming: list) -> list:
    out = []
    by_id = {}
    for row in existing or []:
        pid = str((row or {}).get("id") or "").strip()
        if not pid:
            continue
        by_id[pid] = row
        out.append(pid)
    for row in incoming or []:
        pid = str((row or {}).get("id") or "").strip()
        if not pid:
            continue
        if pid in by_id:
            by_id[pid] = {**by_id[pid], **row}
        else:
            by_id[pid] = row
            out.append(pid)
    return [by_id[pid] for pid in out if pid in by_id]


def _read_index_entries() -> list:
    idx_path = PROJECTS_DIR / "index.json"
    if not idx_path.exists():
        return []
    try:
        return json.loads(idx_path.read_text(encoding="utf-8"))
    except Exception:
        return []


def _write_index_entries(entries: list) -> None:
    idx_path = PROJECTS_DIR / "index.json"
    idx_path.write_text(json.dumps(entries, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _sync_embedded_bundle(index_entries: list) -> int:
    if not HTML_FILE.exists():
        return 0
    active = [p for p in (index_entries or []) if p.get("active", True)]
    bundle = []
    for p in active:
        pid = str(p.get("id", "")).strip()
        if not pid:
            continue
        proj_dir = PROJECTS_DIR / pid
        cfg_path = proj_dir / "config.json"
        data_path = proj_dir / "data.json"
        if not cfg_path.exists() or not data_path.exists():
            continue
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        data = json.loads(data_path.read_text(encoding="utf-8"))
        bundle.append({"meta": p, "config": cfg, "data": data})
    text = HTML_FILE.read_text(encoding="utf-8")
    text = re.sub(
        r"var BUNDLE = .*?;\s*var IDX\s*=\s*.*?;",
        "var BUNDLE = "
        + json.dumps(bundle, ensure_ascii=False, separators=(",", ":"))
        + ";\n  var IDX    = "
        + json.dumps(active, ensure_ascii=False, separators=(",", ":"))
        + ";",
        text,
        count=1,
        flags=re.S,
    )
    HTML_FILE.write_text(text, encoding="utf-8")
    return len(bundle)


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_POST(self) -> None:  # noqa: N802
        if self.path == "/api/save-project":
            return self._handle_save_project()
        if self.path == "/api/create-project":
            return self._handle_create_project()
        if self.path == "/api/delete-project":
            return self._handle_delete_project()
        _json_response(self, 404, {"error": "Not found"})

    def _read_payload(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length > 0 else b"{}"
        payload = json.loads(raw.decode("utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("payload must be object")
        return payload

    def _handle_save_project(self) -> None:
        try:
            payload = self._read_payload()

            pid = _safe_project_id(payload.get("projectId"))
            replace_pid = payload.get("replaceProjectId")
            if replace_pid is not None:
                replace_pid = _safe_project_id(replace_pid)
            config = payload.get("config")
            data = payload.get("data")
            index_entries = payload.get("indexEntries")
            if not isinstance(config, dict) or not isinstance(data, dict):
                raise ValueError("config and data must be objects")

            project_dir = PROJECTS_DIR / pid
            project_dir.mkdir(parents=True, exist_ok=True)
            (project_dir / "config.json").write_text(
                json.dumps(config, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            (project_dir / "data.json").write_text(
                json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )

            if isinstance(index_entries, list):
                existing_entries = _read_index_entries()
                merged = _merge_index_entries(existing_entries, index_entries)
                if replace_pid and replace_pid != pid:
                    merged = [row for row in merged if str((row or {}).get("id") or "").strip() != replace_pid]
                _write_index_entries(merged)
                index_entries = merged
                synced_projects = _sync_embedded_bundle(index_entries)
            else:
                synced_projects = _sync_embedded_bundle([])

            _json_response(
                self,
                200,
                {
                    "ok": True,
                    "projectId": pid,
                    "savedAt": datetime.now().isoformat(timespec="seconds"),
                    "indexEntries": index_entries if isinstance(index_entries, list) else None,
                    "embeddedSyncedProjects": synced_projects,
                },
            )
        except Exception as exc:  # noqa: BLE001
            _json_response(self, 400, {"ok": False, "error": str(exc)})

    def _handle_create_project(self) -> None:
        try:
            payload = self._read_payload()
            pid = _safe_project_id(payload.get("projectId"))
            pname = str(payload.get("projectName") or pid).strip() or pid
            color = str(payload.get("color") or "#1d4ed8").strip() or "#1d4ed8"

            project_dir = PROJECTS_DIR / pid
            project_dir.mkdir(parents=True, exist_ok=True)
            cfg_path = project_dir / "config.json"
            data_path = project_dir / "data.json"

            if not cfg_path.exists():
                cfg = {
                    "project_name": pname,
                    "project": pid,
                    "client": "VNPT",
                    "start_date": "",
                    "end_date": "",
                    "required_complete_date": "",
                    "region": "",
                    "pm": "",
                    "weekly_target": 20,
                    "po_no": "",
                }
                cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            if not data_path.exists():
                data = {
                    "generated": datetime.now().isoformat(timespec="seconds"),
                    "project": pid,
                    "sites": [],
                    "issues": {"raised": 0, "closed": 0, "open": 0, "weekly": []},
                    "resources": {"weekly": []},
                    "rolling_plan": {"weekly": []},
                    "risks": [],
                }
                data_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

            entries = _read_index_entries()
            incoming = [{
                "id": pid,
                "name": pname,
                "path": f"projects/{pid}/",
                "color": color,
                "active": True,
                "updated": datetime.now().strftime("%Y-%m-%d"),
            }]
            entries = _merge_index_entries(entries, incoming)
            _write_index_entries(entries)
            synced_projects = _sync_embedded_bundle(entries)
            _json_response(self, 200, {"ok": True, "projectId": pid, "indexEntries": entries, "embeddedSyncedProjects": synced_projects})
        except Exception as exc:  # noqa: BLE001
            _json_response(self, 400, {"ok": False, "error": str(exc)})

    def _handle_delete_project(self) -> None:
        try:
            payload = self._read_payload()
            pid = _safe_project_id(payload.get("projectId"))
            project_dir = PROJECTS_DIR / pid
            if project_dir.exists():
                shutil.rmtree(project_dir)
            entries = [e for e in _read_index_entries() if str((e or {}).get("id") or "").strip() != pid]
            _write_index_entries(entries)
            synced_projects = _sync_embedded_bundle(entries)
            _json_response(self, 200, {"ok": True, "projectId": pid, "indexEntries": entries, "embeddedSyncedProjects": synced_projects})
        except Exception as exc:  # noqa: BLE001
            _json_response(self, 400, {"ok": False, "error": str(exc)})


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    server = ThreadingHTTPServer(("0.0.0.0", args.port), DashboardHandler)
    print(f"Serving dashboard at http://localhost:{args.port}")
    print("Save endpoint: POST /api/save-project")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()

