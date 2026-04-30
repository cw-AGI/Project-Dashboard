#!/usr/bin/env python3
"""
Sync embedded IDX/BUNDLE in WN_Dashboard_v3.6.html from projects/*.json.

Why:
- file:// mode in the dashboard uses embedded snapshot.
- http(s) mode uses projects/index.json + projects/<id>/{config,data}.json.
- This script keeps file:// fallback aligned with real project JSON files.
"""

from __future__ import annotations

import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
def _discover_dashboard_files() -> list[Path]:
    return sorted(
        p for p in ROOT.glob("WN_Dashboard_v*.html")
        if p.is_file()
    )
PROJECTS = ROOT / "projects"
INDEX = PROJECTS / "index.json"


def main() -> None:
    existing_html = _discover_dashboard_files()
    if not existing_html:
        raise SystemExit("Missing dashboard files matching: WN_Dashboard_v*.html")
    if not INDEX.exists():
        raise SystemExit(f"Missing projects index: {INDEX}")

    idx = json.loads(INDEX.read_text(encoding="utf-8"))
    active = [p for p in idx if p.get("active", True)]
    bundle = []
    for p in active:
        pid = str(p.get("id", "")).strip()
        if not pid:
            continue
        proj_dir = PROJECTS / pid
        cfg_path = proj_dir / "config.json"
        data_path = proj_dir / "data.json"
        if not cfg_path.exists() or not data_path.exists():
            print(f"Skip {pid}: missing config/data json")
            continue
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        data = json.loads(data_path.read_text(encoding="utf-8"))
        bundle.append({"meta": p, "config": cfg, "data": data})

    pattern = re.compile(
        r"var BUNDLE = .*?;\s*var IDX\s*=\s*.*?;",
        re.S,
    )
    repl = (
        "var BUNDLE = "
        + json.dumps(bundle, ensure_ascii=False, separators=(",", ":"))
        + ";\n"
        + "  var IDX    = "
        + json.dumps(active, ensure_ascii=False, separators=(",", ":"))
        + ";"
    )
    synced_files = []
    skipped_files = []
    for html in existing_html:
        text = html.read_text(encoding="utf-8")
        if not pattern.search(text):
            skipped_files.append(html.name)
            continue
        text = pattern.sub(repl, text, count=1)
        html.write_text(text, encoding="utf-8")
        synced_files.append(html.name)
    if not synced_files:
        raise SystemExit("No compatible dashboard files with embedded BUNDLE/IDX block were found.")
    msg = f"Synced embedded snapshot: {len(bundle)} project(s) -> {', '.join(synced_files)}"
    if skipped_files:
        msg += f" | skipped (no embedded block): {', '.join(skipped_files)}"
    print(msg)


if __name__ == "__main__":
    main()

