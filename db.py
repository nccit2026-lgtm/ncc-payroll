import json
import os
import sqlite3

from flask import current_app, g
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# জাতীয় বেতন স্কেল ২০২৬ (৯ম পে স্কেল) — প্রারম্ভিক ও সর্বোচ্চ মূল বেতন।
# গেজেটের সাথে মিলিয়ে "বেতন স্কেল ও ভাতা" পাতা থেকে যেকোনো সময় সংশোধন করা যাবে।
DEFAULT_PAY_GRADES = {
    1: (156000, 156000),
    2: (132000, 152980),
    3: (113000, 148800),
    4: (100000, 142400),
    5: (86000, 139700),
    6: (71000, 134020),
    7: (58000, 126800),
    8: (46000, 110800),
    9: (44000, 106120),
    10: (32000, 77280),
    11: (25000, 60500),
    12: (24300, 58700),
    13: (24000, 58000),
    14: (23500, 56800),
    15: (22800, 55200),
    16: (21900, 52900),
    17: (21400, 51900),
    18: (21000, 50900),
    19: (20500, 49600),
    20: (20000, 48400),
}

DEFAULT_SETTINGS = {
    "office_name": "নারায়ণগঞ্জ সিটি কর্পোরেশন",
    "office_address": "নগর ভবন, নারায়ণগঞ্জ",
    # বাড়ি ভাড়া ভাতা: গ্রেডের পরিসর অনুযায়ী মূল বেতনের শতাংশ (সর্বনিম্ন টাকা সহ)
    "house_rent_bands": [
        {"from": 1, "to": 4, "pct": 40, "min": 0},
        {"from": 5, "to": 9, "pct": 45, "min": 0},
        {"from": 10, "to": 15, "pct": 50, "min": 0},
        {"from": 16, "to": 20, "pct": 60, "min": 0},
    ],
    "medical_below_50": 3000,
    "medical_50_plus": 4000,
    "education_per_child": 500,
    "education_max_children": 2,
    "tiffin_amount": 200,
    "tiffin_from_grade": 11,
    "transport_amount": 300,
    "transport_from_grade": 11,
    "revenue_stamp": 10,
    "festival_pct": 100,
    "boishakhi_pct": 20,
    "increment_pct": 5,
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

    for grade, (start, mx) in DEFAULT_PAY_GRADES.items():
        conn.execute(
            "INSERT OR IGNORE INTO pay_grades (grade, start_basic, max_basic) VALUES (?, ?, ?)",
            (grade, start, mx),
        )
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


def get_pay_grades():
    rows = get_db().execute("SELECT * FROM pay_grades ORDER BY grade").fetchall()
    return {r["grade"]: r for r in rows}
