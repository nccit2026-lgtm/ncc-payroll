"""বেতন হিসাবের নিয়ম — মূল বেতন থেকে ভাতা ও কর্তন নির্ণয়।"""
import calendar
import json
from datetime import date

from db import get_db, get_settings


def _round(x):
    return int(round(x))


def age_on(dob, on):
    if not dob:
        return None
    y, m, d = (int(p) for p in dob.split("-"))
    return on.year - y - ((on.month, on.day) < (m, d))


def is_over_50(dob, on):
    """৫০ বছর এক দিন পূর্ণ হলে উচ্চতর চিকিৎসা ভাতা প্রযোজ্য।"""
    if not dob:
        return False
    y, m, d = (int(p) for p in dob.split("-"))
    try:
        fiftieth = date(y + 50, m, d)
    except ValueError:  # ২৯ ফেব্রুয়ারি
        fiftieth = date(y + 50, 3, 1)
    return on > fiftieth


def house_rent_for(basic, grade, settings):
    for band in settings["house_rent_bands"]:
        if grade and int(band["from"]) <= grade <= int(band["to"]):
            return max(_round(basic * float(band["pct"]) / 100), int(band.get("min") or 0))
    return 0


def compute_components(emp, settings, year, month, include_festival=False, include_boishakhi=False):
    """একজন কর্মচারীর এক মাসের সম্পূর্ণ বেতন বিবরণী হিসাব করে dict ফেরত দেয়।"""
    month_end = date(year, month, calendar.monthrange(year, month)[1])
    basic = int(emp["basic"] or 0)
    grade = emp["grade"]

    def pick(override, auto):
        return int(override) if override is not None else auto

    house_rent = 0 if emp["govt_quarter"] else house_rent_for(basic, grade, settings)
    medical = settings["medical_50_plus"] if is_over_50(emp["dob"], month_end) else settings["medical_below_50"]
    education = min(int(emp["children"] or 0), int(settings["education_max_children"])) * int(
        settings["education_per_child"]
    )
    tiffin = settings["tiffin_amount"] if grade and grade >= int(settings["tiffin_from_grade"]) else 0
    transport = settings["transport_amount"] if grade and grade >= int(settings["transport_from_grade"]) else 0

    items = get_db().execute(
        "SELECT kind, name, amount FROM employee_items WHERE employee_id = ? ORDER BY id", (emp["id"],)
    ).fetchall()

    slip = {
        "basic": basic,
        "house_rent": pick(emp["ovr_house_rent"], house_rent),
        "medical": pick(emp["ovr_medical"], int(medical)),
        "education": pick(emp["ovr_education"], education),
        "tiffin": pick(emp["ovr_tiffin"], int(tiffin)),
        "transport": pick(emp["ovr_transport"], int(transport)),
        "festival": _round(basic * float(settings["festival_pct"]) / 100) if include_festival else 0,
        "boishakhi": _round(basic * float(settings["boishakhi_pct"]) / 100) if include_boishakhi else 0,
        "other_allowance": [{"name": i["name"], "amount": i["amount"]} for i in items if i["kind"] == "allowance"],
        "gpf": _round(basic * float(emp["gpf_pct"] or 0) / 100),
        "revenue_stamp": int(settings["revenue_stamp"]),
        "other_deduction": [{"name": i["name"], "amount": i["amount"]} for i in items if i["kind"] == "deduction"],
    }
    return finalize_totals(slip)


ALLOWANCE_FIELDS = ["basic", "house_rent", "medical", "education", "tiffin", "transport", "festival", "boishakhi"]
DEDUCTION_FIELDS = ["gpf", "revenue_stamp"]

FIELD_LABELS = {
    "basic": "মূল বেতন",
    "house_rent": "বাড়ি ভাড়া ভাতা",
    "medical": "চিকিৎসা ভাতা",
    "education": "শিক্ষা সহায়ক ভাতা",
    "tiffin": "টিফিন ভাতা",
    "transport": "যাতায়াত ভাতা",
    "festival": "উৎসব ভাতা",
    "boishakhi": "বাংলা নববর্ষ ভাতা",
    "gpf": "জিপিএফ চাঁদা",
    "revenue_stamp": "রাজস্ব স্ট্যাম্প",
}


def _items(value):
    return json.loads(value) if isinstance(value, str) else (value or [])


def finalize_totals(slip):
    other_a = sum(int(i["amount"]) for i in _items(slip["other_allowance"]))
    other_d = sum(int(i["amount"]) for i in _items(slip["other_deduction"]))
    slip["gross"] = sum(int(slip[f]) for f in ALLOWANCE_FIELDS) + other_a
    slip["total_deduction"] = sum(int(slip[f]) for f in DEDUCTION_FIELDS) + other_d
    slip["net"] = slip["gross"] - slip["total_deduction"]
    return slip


def generate_run(year, month, include_festival, include_boishakhi, note=""):
    """নির্দিষ্ট মাসের জন্য সকল সক্রিয় কর্মচারীর খসড়া বেতন শিট তৈরি করে।"""
    conn = get_db()
    settings = get_settings()
    cur = conn.execute(
        "INSERT INTO payroll_runs (year, month, include_festival, include_boishakhi, note) VALUES (?, ?, ?, ?, ?)",
        (year, month, int(include_festival), int(include_boishakhi), note),
    )
    run_id = cur.lastrowid
    employees = conn.execute("SELECT * FROM employees WHERE status = 'active' ORDER BY grade, code").fetchall()
    for emp in employees:
        insert_slip(run_id, emp, compute_components(emp, settings, year, month, include_festival, include_boishakhi))
    conn.commit()
    return run_id


def insert_slip(run_id, emp, s):
    get_db().execute(
        """INSERT INTO payslips (run_id, employee_id, code, name_bn, designation, department, grade, account_no,
               basic, house_rent, medical, education, tiffin, transport, festival, boishakhi, other_allowance,
               gpf, revenue_stamp, other_deduction, gross, total_deduction, net)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            run_id, emp["id"], emp["code"], emp["name_bn"], emp["designation"], emp["department"], emp["grade"],
            emp["account_no"], s["basic"], s["house_rent"], s["medical"], s["education"], s["tiffin"],
            s["transport"], s["festival"], s["boishakhi"], json.dumps(s["other_allowance"], ensure_ascii=False),
            s["gpf"], s["revenue_stamp"], json.dumps(s["other_deduction"], ensure_ascii=False), s["gross"],
            s["total_deduction"], s["net"],
        ),
    )


def regenerate_slip(run, slip_row):
    """খসড়া শিটে কর্মচারীর বর্তমান তথ্য অনুযায়ী একটি বিবরণী পুনরায় হিসাব করে।"""
    conn = get_db()
    emp = conn.execute("SELECT * FROM employees WHERE id = ?", (slip_row["employee_id"],)).fetchone()
    s = compute_components(emp, get_settings(), run["year"], run["month"], run["include_festival"], run["include_boishakhi"])
    conn.execute("DELETE FROM payslips WHERE id = ?", (slip_row["id"],))
    insert_slip(run["id"], emp, s)
