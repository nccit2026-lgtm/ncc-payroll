"""মাসিক বেতন শিট — payscale.py এর নিয়মে প্রতি কর্মচারীর মাসের বেতন বিবরণী তৈরি।"""
import json

from bn import bn_num
from db import get_db, get_settings
from payscale import age_on_date, allowance_basic, arrears, month_components, phase_of, r2

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

# সারণির ছোট কলাম-নাম
SHORT_LABELS = {"basic": "মূল বেতন", "house_rent": "বাড়িভাড়া", "medical": "চিকিৎসা", "education": "শিক্ষা",
                "tiffin": "টিফিন", "transport": "যাতায়াত", "mobile": "মোবাইল", "wash": "ধোলাই", "special": "মহার্ঘ",
                "pf_office": "অফিস PF", "festival": "উৎসব", "boishakhi": "নববর্ষ"}

ARREARS_LABEL = "বকেয়া: জুলাই–সেপ্টেম্বর ২০২৬"

RELIGIONS = ["ইসলাম", "হিন্দু", "বৌদ্ধ", "খ্রিস্টান", "অন্যান্য"]
# বেতন ছাড়া আলাদা শিটের ধরন
RUN_KINDS = {
    "salary": "মাসিক বেতন",
    "festival": "উৎসব ভাতা",
    "boishakhi": "বাংলা নববর্ষ ভাতা",
    "arrears": "বকেয়া (জুলাই–সেপ্টেম্বর ২০২৬)",
    "other": "অন্যান্য ভাতা",
}
# উৎসব ভাতা: ঈদ (ইসলাম) মূল বেতনের ১ গুণ; হিন্দু/বৌদ্ধ/খ্রিস্টান ধর্মের উৎসবে ২ গুণ
FESTIVAL_MULTIPLIER = {"ইসলাম": 1, "হিন্দু": 2, "বৌদ্ধ": 2, "খ্রিস্টান": 2}


def age_on(dob, on):
    return age_on_date(dob, on)


def compute_components(emp, settings, year, month, include_festival=False, include_boishakhi=False,
                       include_arrears=False, run=None):
    """একজন কর্মচারীর এক মাসের সম্পূর্ণ বেতন বিবরণী; হিসাব করা না গেলে (তথ্য অসম্পূর্ণ) None।"""
    c, plan = month_components(emp, year, month)
    if c is None:
        return None
    items = get_db().execute(
        "SELECT kind, name, amount FROM employee_items WHERE employee_id = ? ORDER BY id", (emp["id"],)
    ).fetchall()
    other_a = [{"name": i["name"], "amount": i["amount"]} for i in items if i["kind"] == "allowance"]
    other_d = [{"name": i["name"], "amount": i["amount"]} for i in items if i["kind"] == "deduction"]
    other_d += loan_deductions(emp, run or {"id": 0, "year": year, "month": month})
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
    b = allowance_basic(emp, plan, year, month)  # উৎসব/নববর্ষ ভাতার ভিত্তি
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


def _empty_slip():
    return {**{f: 0 for f in ALLOWANCE_FIELDS + DEDUCTION_FIELDS}, "other_allowance": [], "other_deduction": []}


def compute_extra(emp, run, settings):
    """বেতন ছাড়া শিট (উৎসব/নববর্ষ/বকেয়া/অন্যান্য ভাতা)। ফেরত: (slip, None) অথবা (None, বাদ পড়ার কারণ)।"""
    kind = run["kind"]
    if kind == "festival" and run["religion"]:
        if not emp["religion"]:
            return None, "ধর্ম দেওয়া নেই"
        if emp["religion"] != run["religion"]:
            return None, ""  # অন্য ধর্ম — এই উৎসব ভাতা প্রযোজ্য নয়
    c, plan = month_components(emp, run["year"], run["month"])
    if c is None:
        return None, "গ্রেড/মূল বেতন অসম্পূর্ণ"
    base = allowance_basic(emp, plan, run["year"], run["month"])
    slip = _empty_slip()
    slip["ref_basic"] = base
    slip["phase"] = run["title"] or RUN_KINDS[kind]
    if kind == "festival":
        slip["festival"] = r2(base * float(run["multiplier"] or 1))
    elif kind == "boishakhi":
        slip["boishakhi"] = r2(base * float(settings["boishakhi_pct"]) / 100)
    elif kind == "arrears":
        a = arrears(emp)
        if not a or not a["total"]:
            return None, ""
        if a["total"] > 0:
            slip["other_allowance"].append({"name": ARREARS_LABEL, "amount": a["total"]})
        else:
            slip["other_deduction"].append({"name": ARREARS_LABEL + " (সমন্বয়)", "amount": -a["total"]})
        if a["pf_total"]:
            slip["other_allowance"].append({"name": "বকেয়ার উপর অফিস প্রদত্ত PF", "amount": a["pf_total"]})
            slip["other_deduction"].append({"name": "বকেয়ার উপর PF (অফিস + নিজ অংশ)", "amount": r2(2 * a["pf_total"])})
    elif kind == "other":
        amt = r2(base * float(run["pct"]) / 100) if run["pct"] else r2(float(run["amount"] or 0))
        slip["other_allowance"].append({"name": run["title"] or "ভাতা", "amount": amt})
    return finalize_totals(slip), None


def slip_for(emp, run, settings):
    """যেকোনো ধরনের শিটের জন্য একজনের বিবরণী: (slip, None) অথবা (None, কারণ)।"""
    if run["kind"] == "salary":
        s = compute_components(emp, settings, run["year"], run["month"], run["include_festival"],
                               run["include_boishakhi"], run["include_arrears"], run=run)
        return (s, None) if s else (None, "গ্রেড/মূল বেতন অসম্পূর্ণ")
    return compute_extra(emp, run, settings)


def arrears_already_given(exclude_run_id=None):
    """জুলাই–সেপ্টেম্বর ২০২৬ এর বকেয়া আগে কোনো শিটে দেওয়া হয়েছে কিনা (দুবার দেওয়া ঠেকাতে)।"""
    return get_db().execute(
        "SELECT * FROM payroll_runs WHERE (kind = 'arrears' OR include_arrears = 1) AND manual = 0 AND id != ?",
        (exclude_run_id or 0,)).fetchone()


def _items(value):
    return json.loads(value) if isinstance(value, str) else (value or [])


def finalize_totals(slip):
    other_a = sum(float(i["amount"]) for i in _items(slip["other_allowance"]))
    other_d = sum(float(i["amount"]) for i in _items(slip["other_deduction"]))
    slip["gross"] = r2(sum(float(slip[f]) for f in ALLOWANCE_FIELDS) + other_a)
    slip["total_deduction"] = r2(sum(float(slip[f]) for f in DEDUCTION_FIELDS) + other_d)
    slip["net"] = r2(slip["gross"] - slip["total_deduction"])
    return slip


def generate_run(year, month, include_festival, include_boishakhi, note="", include_arrears=False, kind="salary",
                 title=None, religion=None, multiplier=None, pct=None, amount=None):
    """সকল সক্রিয় কর্মচারীর খসড়া শিট তৈরি করে; ফেরত: (run_id, [(কর্মচারী, কারণ)] যাদের হিসাব হয়নি)।"""
    conn = get_db()
    settings = get_settings()
    cur = conn.execute(
        "INSERT INTO payroll_runs (year, month, kind, title, religion, multiplier, pct, amount, include_festival, "
        "include_boishakhi, include_arrears, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (year, month, kind, title, religion, multiplier, pct, amount, int(include_festival), int(include_boishakhi),
         int(include_arrears), note),
    )
    run = conn.execute("SELECT * FROM payroll_runs WHERE id = ?", (cur.lastrowid,)).fetchone()
    skipped = []
    for emp in conn.execute("SELECT * FROM employees WHERE status = 'active' ORDER BY code"):
        s, why = slip_for(emp, run, settings)
        if s is not None:
            insert_slip(run["id"], emp, s)
        elif why:
            skipped.append((emp, why))
    conn.commit()
    return run["id"], skipped


SLIP_COLUMNS = ALLOWANCE_FIELDS + DEDUCTION_FIELDS + ["gross", "total_deduction", "net", "phase", "ref_basic"]


def insert_slip(run_id, emp, s):
    cols = ["run_id", "employee_id", "code", "name_bn", "designation", "department", "grade", "account_no",
            "other_allowance", "other_deduction"] + SLIP_COLUMNS
    vals = [run_id, emp["id"], emp["code"], emp["name_bn"], emp["designation"], emp["department"], emp["grade"],
            emp["account_no"], json.dumps(s["other_allowance"], ensure_ascii=False),
            json.dumps(s["other_deduction"], ensure_ascii=False)] + [s.get(c) for c in SLIP_COLUMNS]
    get_db().execute(f"INSERT INTO payslips ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})", vals)
    record_loan_payments(run_id, emp["id"], s["other_deduction"])


def regenerate_slip(run, slip_row):
    """খসড়া শিটে কর্মচারীর বর্তমান তথ্য অনুযায়ী একটি বিবরণী পুনরায় হিসাব করে।"""
    conn = get_db()
    emp = conn.execute("SELECT * FROM employees WHERE id = ?", (slip_row["employee_id"],)).fetchone()
    s, _why = slip_for(emp, run, get_settings())
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


# ---------------------------------------------------------------- ঋণের কিস্তি

LOAN_KINDS = ["গৃহ নির্মাণ ঋণ", "ভবিষ্য তহবিল অগ্রিম", "অন্যান্য ঋণ", "মোটরসাইকেল ঋণ", "কম্পিউটার ঋণ"]


def loan_progress(loan, exclude_run_id=None):
    """পরিশোধিত (চূড়ান্ত শিটে কাটা + আগের) ও খসড়ায় থাকা কিস্তির সংখ্যা।"""
    rows = get_db().execute(
        "SELECT r.status, p.amount FROM loan_payments p JOIN payroll_runs r ON r.id = p.run_id "
        "WHERE p.loan_id = ? AND p.run_id != ?", (loan["id"], exclude_run_id or 0)).fetchall()
    final = sum(1 for r in rows if r["status"] == "final")
    draft = sum(1 for r in rows if r["status"] != "final")
    paid = loan["paid_before"] + final
    total = loan["total_installments"]
    return {"paid": paid, "draft": draft, "total": total,
            "remaining": max(total - paid, 0) if total else None,
            "remaining_amount": r2(max(total - paid, 0) * loan["installment"]) if total else None,
            "done": bool(total) and paid >= total}


def loan_deductions(emp, run):
    """এই শিটে কাটার মতো ঋণের কিস্তি: [{name, amount, loan_id}]। মোট কিস্তি পূর্ণ হলে আর কাটে না।"""
    out = []
    ym = f"{run['year']:04d}-{run['month']:02d}"
    for loan in get_db().execute("SELECT * FROM loans WHERE employee_id = ? AND status = 'active' ORDER BY id",
                                 (emp["id"],)):
        if loan["start_ym"] and loan["start_ym"] > ym:
            continue
        pr = loan_progress(loan, exclude_run_id=run["id"])
        n = pr["paid"] + pr["draft"]  # এর আগের সব কিস্তি (চূড়ান্ত + অন্য খসড়া)
        if pr["total"] and n >= pr["total"]:
            continue
        label = f"{loan['kind']} কিস্তি ({bn_num(n + 1)}" + (f"/{bn_num(pr['total'])})" if pr["total"] else ")")
        out.append({"name": label, "amount": loan["installment"], "loan_id": loan["id"]})
    return out


def record_loan_payments(run_id, employee_id, other_deduction):
    """শিটে বসানো কিস্তিগুলো loan_payments এ লেখা (আগেরগুলো মুছে)।"""
    conn = get_db()
    conn.execute("DELETE FROM loan_payments WHERE run_id = ? AND loan_id IN "
                 "(SELECT id FROM loans WHERE employee_id = ?)", (run_id, employee_id))
    for item in _items(other_deduction):
        if item.get("loan_id"):
            conn.execute("INSERT OR REPLACE INTO loan_payments (loan_id, run_id, amount) VALUES (?, ?, ?)",
                         (item["loan_id"], run_id, item["amount"]))
