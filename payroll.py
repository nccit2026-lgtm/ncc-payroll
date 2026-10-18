"""মাসিক বেতন শিট — payscale.py এর নিয়মে প্রতি কর্মচারীর মাসের বেতন বিবরণী তৈরি।"""
import json

from db import get_db, get_settings
from payscale import age_on_date, arrears, month_components, phase_of, r2

ALLOWANCE_FIELDS = ["basic", "house_rent", "medical", "education", "tiffin", "transport", "mobile", "wash", "special",
                    "pf_office", "festival", "boishakhi"]
DEDUCTION_FIELDS = ["pf_fund", "pf_own"]  # রাজস্ব স্ট্যাম্প কাটা হয় না

FIELD_LABELS = {
    "basic": "মূল বেতন",
    "house_rent": "বাড়িভাড়া ভাতা",
    "medical": "চিকিৎসা ভাতা",
    "education": "শিক্ষা সহায়ক ভাতা",
    "tiffin": "টিফিন ভাতা",
    "transport": "যাতায়াত ভাতা",
    "mobile": "মোবাইল ভাতা",
    "wash": "ধোলাই ভাতা",
    "special": "বিশেষ সুবিধা (মহার্ঘ ভাতা)",
    "pf_office": "অফিস প্রদত্ত ভবিষ্য তহবিল (PF)",
    "festival": "উৎসব ভাতা",
    "boishakhi": "বাংলা নববর্ষ ভাতা",
    "pf_fund": "অফিস প্রদত্ত PF (ফান্ডে জমা)",
    "pf_own": "নিজ PF অংশ",
    "revenue_stamp": "রাজস্ব স্ট্যাম্প",
}

ARREARS_LABEL = "বকেয়া: জুলাই–সেপ্টেম্বর ২০২৬"


def age_on(dob, on):
    return age_on_date(dob, on)


def compute_components(emp, settings, year, month, include_festival=False, include_boishakhi=False,
                       include_arrears=False):
    """একজন কর্মচারীর এক মাসের সম্পূর্ণ বেতন বিবরণী; হিসাব করা না গেলে (তথ্য অসম্পূর্ণ) None।"""
    c, _plan = month_components(emp, year, month)
    if c is None:
        return None
    items = get_db().execute(
        "SELECT kind, name, amount FROM employee_items WHERE employee_id = ? ORDER BY id", (emp["id"],)
    ).fetchall()
    other_a = [{"name": i["name"], "amount": i["amount"]} for i in items if i["kind"] == "allowance"]
    other_d = [{"name": i["name"], "amount": i["amount"]} for i in items if i["kind"] == "deduction"]
    if include_arrears:
        a = arrears(emp)
        if a and a["total"]:
            if a["total"] > 0:
                other_a.append({"name": ARREARS_LABEL, "amount": a["total"]})
            else:
                other_d.append({"name": ARREARS_LABEL + " (সমন্বয়)", "amount": -a["total"]})
            if a["pf_total"]:
                other_a.append({"name": "বকেয়ার উপর অফিস প্রদত্ত PF", "amount": a["pf_total"]})
                other_d.append({"name": "বকেয়ার উপর PF (অফিস + নিজ অংশ)", "amount": r2(2 * a["pf_total"])})
    b = c["basic"]
    slip = {
        **{k: c[k] for k in ["basic", "house_rent", "medical", "education", "tiffin", "transport", "mobile", "wash",
                             "special", "pf_office", "pf_fund", "pf_own"]},
        "festival": r2(b * float(settings["festival_pct"]) / 100) if include_festival else 0,
        "boishakhi": r2(b * float(settings["boishakhi_pct"]) / 100) if include_boishakhi else 0,
        "other_allowance": other_a,
        "other_deduction": other_d,
        "phase": phase_of(year, month)[0],
    }
    return finalize_totals(slip)


def _items(value):
    return json.loads(value) if isinstance(value, str) else (value or [])


def finalize_totals(slip):
    other_a = sum(float(i["amount"]) for i in _items(slip["other_allowance"]))
    other_d = sum(float(i["amount"]) for i in _items(slip["other_deduction"]))
    slip["gross"] = r2(sum(float(slip[f]) for f in ALLOWANCE_FIELDS) + other_a)
    slip["total_deduction"] = r2(sum(float(slip[f]) for f in DEDUCTION_FIELDS) + other_d)
    slip["net"] = r2(slip["gross"] - slip["total_deduction"])
    return slip


def generate_run(year, month, include_festival, include_boishakhi, note="", include_arrears=False):
    """নির্দিষ্ট মাসের জন্য সকল সক্রিয় কর্মচারীর খসড়া বেতন শিট তৈরি করে; যাদের হিসাব হয়নি তাদের তালিকা ফেরত দেয়।"""
    conn = get_db()
    settings = get_settings()
    cur = conn.execute(
        "INSERT INTO payroll_runs (year, month, include_festival, include_boishakhi, include_arrears, note) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (year, month, int(include_festival), int(include_boishakhi), int(include_arrears), note),
    )
    run_id = cur.lastrowid
    skipped = []
    for emp in conn.execute("SELECT * FROM employees WHERE status = 'active' ORDER BY department, grade, code"):
        s = compute_components(emp, settings, year, month, include_festival, include_boishakhi, include_arrears)
        if s is None:
            skipped.append(emp)
        else:
            insert_slip(run_id, emp, s)
    conn.commit()
    return run_id, skipped


SLIP_COLUMNS = ALLOWANCE_FIELDS + DEDUCTION_FIELDS + ["gross", "total_deduction", "net", "phase"]


def insert_slip(run_id, emp, s):
    cols = ["run_id", "employee_id", "code", "name_bn", "designation", "department", "grade", "account_no",
            "other_allowance", "other_deduction"] + SLIP_COLUMNS
    vals = [run_id, emp["id"], emp["code"], emp["name_bn"], emp["designation"], emp["department"], emp["grade"],
            emp["account_no"], json.dumps(s["other_allowance"], ensure_ascii=False),
            json.dumps(s["other_deduction"], ensure_ascii=False)] + [s[c] for c in SLIP_COLUMNS]
    get_db().execute(f"INSERT INTO payslips ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})", vals)


def regenerate_slip(run, slip_row):
    """খসড়া শিটে কর্মচারীর বর্তমান তথ্য অনুযায়ী একটি বিবরণী পুনরায় হিসাব করে।"""
    conn = get_db()
    emp = conn.execute("SELECT * FROM employees WHERE id = ?", (slip_row["employee_id"],)).fetchone()
    s = compute_components(emp, get_settings(), run["year"], run["month"], run["include_festival"],
                           run["include_boishakhi"], run["include_arrears"])
    if s is None:
        return False
    conn.execute("DELETE FROM payslips WHERE id = ?", (slip_row["id"],))
    insert_slip(run["id"], emp, s)
    return True


def group_by_department(slips):
    """শাখা অনুযায়ী ভাগ — প্রথমবার যে ক্রমে শাখা আসে সেই ক্রমে (আইডি ক্রম অনুযায়ী)।"""
    groups = {}
    for s in slips:
        groups.setdefault(s["department"] or "শাখা উল্লেখ নেই", []).append(s)
    return list(groups.items())
