"""Shared business logic for SQLWithPython.

Features
--------
* load_database()   – Load demo.sql to in-memory SQLite, or connect to
                      MySQL if .env exists.
* run_query(sql)    – Execute a SELECT statement and return (columns, rows).
* export_excel()    – Dump query result to .xlsx via openpyxl.
"""
from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Any, Iterable, Tuple, List
import re

try:
    import pymysql  # optional
except ModuleNotFoundError:
    pymysql = None  # type: ignore

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    load_dotenv = lambda *_a, **_k: None  # type: ignore

ROOT = Path(__file__).resolve().parent
DEMO_SQL = ROOT / "demo.sql"

def _init_sqlite() -> sqlite3.Connection:
    conn = sqlite3.connect(":memory:")
    conn.executescript(DEMO_SQL.read_text(encoding="utf-8-sig"))
    return conn

def _init_mysql():
    if pymysql is None:
        raise RuntimeError("pymysql not installed")
    load_dotenv()
    return pymysql.connect(
        host=os.getenv("MYSQL_HOST", "localhost"),
        user=os.getenv("MYSQL_USER", "root"),
        password=os.getenv("MYSQL_PASSWORD", ""),
        database=os.getenv("MYSQL_DB", "demo"),
        charset="utf8mb4",
        autocommit=True,
    )

def load_database():
    if Path(".env").exists():
        return _init_mysql()
    return _init_sqlite()

def run_query(sql: str) -> Tuple[List[str], List[Tuple[Any, ...]]]:
    if not re.match(r"^\s*select", sql, re.I):
        raise ValueError("Only SELECT statements are permitted.")
    conn = load_database()
    cur = conn.cursor()
    cur.execute(sql)
    columns = [c[0] for c in cur.description]
    rows = cur.fetchall()
    return columns, rows

def export_excel(path: Path, columns: Iterable[str], rows: Iterable[Iterable[Any]]):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(list(columns))
    for r in rows:
        ws.append(["NULL" if v is None else v for v in r])
    wb.save(path)