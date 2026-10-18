import json
import os
import sqlite3

from flask import current_app, g
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_SETTINGS = {
    "office_name": "নারায়ণগঞ্জ সিটি কর্পোরেশন",
    "office_address": "নগর ভবন, নারায়ণগঞ্জ",
    "festival_pct": 100,
    "boishakhi_pct": 20,
}

# পুরনো ডাটাবেসে নতুন কলাম যোগ (CREATE TABLE IF NOT EXISTS পুরনো টেবিল বদলায় না)
MIGRATIONS = {
    "employees": {
        "inc_2026": "INTEGER NOT NULL DEFAULT 1", "inc_2027": "INTEGER NOT NULL DEFAULT 1",
        "new_basic_override": "REAL", "pf": "INTEGER NOT NULL DEFAULT 0",
        "special_rate": "TEXT NOT NULL DEFAULT 'auto'", "tiffin": "INTEGER NOT NULL DEFAULT 1",
        "transport": "INTEGER NOT NULL DEFAULT 1", "wash": "INTEGER NOT NULL DEFAULT 0",
    },
    "payslips": {
        "mobile": "REAL NOT NULL DEFAULT 0", "wash": "REAL NOT NULL DEFAULT 0", "special": "REAL NOT NULL DEFAULT 0",
        "pf_office": "REAL NOT NULL DEFAULT 0", "pf_fund": "REAL NOT NULL DEFAULT 0", "pf_own": "REAL NOT NULL DEFAULT 0",
        "phase": "TEXT",
    },
    "payroll_runs": {"include_arrears": "INTEGER NOT NULL DEFAULT 0"},
    "users": {"must_change": "INTEGER NOT NULL DEFAULT 0"},
}


def db_path():
    return current_app.config["DATABASE"]


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(db_path(), timeout=15)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(_exc=None):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


def init_db():
    conn = get_db()
    conn.execute("PRAGMA journal_mode=WAL")  # একাধিক ব্যবহারকারী একসাথে কাজ করলে ভালো
    with open(os.path.join(BASE_DIR, "schema.sql"), encoding="utf-8") as f:
        conn.executescript(f.read())
    for table, cols in MIGRATIONS.items():
        have = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
        for col, ddl in cols.items():
            if col not in have:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")

    for key, value in DEFAULT_SETTINGS.items():
        conn.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            (key, json.dumps(value, ensure_ascii=False)),
        )
    if not conn.execute("SELECT 1 FROM users WHERE role = 'admin'").fetchone():
        # প্রথম অ্যাডমিন: ADMIN_PASSWORD env থাকলে সেটি, নইলে admin123; প্রথম লগইনে পরিবর্তন বাধ্যতামূলক
        conn.execute(
            "INSERT INTO users (username, password_hash, role, must_change) VALUES (?, ?, 'admin', 1)",
            ("admin", generate_password_hash(os.environ.get("ADMIN_PASSWORD", "admin123"))),
        )
    conn.commit()


def get_settings():
    rows = get_db().execute("SELECT key, value FROM settings").fetchall()
    settings = dict(DEFAULT_SETTINGS)
    settings.update({r["key"]: json.loads(r["value"]) for r in rows})
    return settings


def save_setting(key, value):
    get_db().execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, json.dumps(value, ensure_ascii=False)),
    )

