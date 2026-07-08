"""Launcher for CLI and Tkinter GUI.

* python run_tool.py                        → opens Tkinter GUI (app.py)
* python run_tool.py --sql "SELECT …"       → run one query in CLI mode
"""
from __future__ import annotations
import argparse, importlib.util, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_PATH = ROOT / "app.py"

spec = importlib.util.spec_from_file_location("app", APP_PATH)
app = importlib.util.module_from_spec(spec)  # type: ignore[arg-type]
spec.loader.exec_module(app)                # type: ignore[union-attr]

def _cli(sql: str) -> None:
    from core_functions import run_query
    cols, rows = run_query(sql)
    print(" | ".join(cols))
    print("-" * 60)
    for r in rows[:10]:
        print(" | ".join(str(v) for v in r))
    if len(rows) > 10:
        print(f"… ({len(rows)} rows total)")

def main() -> None:
    p = argparse.ArgumentParser(description="SQLWithPython launcher")
    p.add_argument("--sql", help="Run single SQL in CLI mode")
    args = p.parse_args()
    if args.sql:
        _cli(args.sql)
    else:
        app.main()  # type: ignore[attr-defined]

if __name__ == "__main__":
    main()