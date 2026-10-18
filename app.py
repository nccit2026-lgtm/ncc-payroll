import json
import math
import os
import secrets
import time
from datetime import date, datetime, timedelta
from functools import wraps

from flask import (Flask, abort, flash, g, redirect, render_template, request, send_file, session, url_for)
from openpyxl import load_workbook
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import check_password_hash, generate_password_hash

import excel
from bn import MONTHS_BN, bn_date, bn_num, bn_pct, bn_taka, month_label
from db import close_db, get_db, get_pay_grades, get_settings, init_db, save_setting
from payroll import (ALLOWANCE_FIELDS, DEDUCTION_FIELDS, FIELD_LABELS, age_on, compute_components, finalize_totals,
                     generate_run, insert_slip, regenerate_slip)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ভুল পাসওয়ার্ডের সীমা: একই IP + আইডি থেকে ১৫ মিনিটে ৫ বার
FAILED_LOGINS = {}
MAX_FAILED_LOGINS = 5
LOCKOUT_SECONDS = 15 * 60


def _secret_key(instance_dir):
    if os.environ.get("SECRET_KEY"):
        return os.environ["SECRET_KEY"]
    path = os.path.join(instance_dir, "secret_key")
    if not os.path.exists(path):
        with open(path, "w") as f:
            f.write(secrets.token_hex(32))
    with open(path) as f:
        return f.read().strip()


def create_app(database=None):
    instance_dir = os.path.join(BASE_DIR, "instance")
    os.makedirs(instance_dir, exist_ok=True)
    app = Flask(__name__)
    app.config.update(
        DATABASE=database or os.environ.get("DATABASE", os.path.join(instance_dir, "payroll.db")),
        SECRET_KEY=_secret_key(instance_dir),
        MAX_CONTENT_LENGTH=10 * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        # HTTPS এ চালালে SECURE_COOKIES=1 দিন
        SESSION_COOKIE_SECURE=os.environ.get("SECURE_COOKIES") == "1",
        PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
    )
    if os.environ.get("BEHIND_PROXY") == "1":  # nginx / হোস্টিং প্রক্সির পেছনে
        app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    app.teardown_appcontext(close_db)
    app.jinja_env.filters.update(bn=bn_num, taka=bn_taka, bndate=bn_date, pct=bn_pct)
    app.jinja_env.globals.update(MONTHS_BN=MONTHS_BN, month_label=month_label, FIELD_LABELS=FIELD_LABELS,
                                 loads=json.loads)
    with app.app_context():
        init_db()
    register_routes(app)
    return app


# ---------------------------------------------------------------- সহায়ক ফাংশন

BN_TO_EN = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")


def to_int(value, default=0):
    if value is None:
        return default
    s = str(value).translate(BN_TO_EN).replace(",", "").strip()
    if s == "":
        return default
    try:
        return int(round(float(s)))
    except ValueError:
        return default


def to_float(value, default=0.0):
    s = str(value or "").translate(BN_TO_EN).replace(",", "").strip()
    try:
        return float(s) if s else default
    except ValueError:
        return default


def optional_int(value):
    """খালি ঘর = None (স্বয়ংক্রিয় হিসাব)।"""
    s = str(value or "").translate(BN_TO_EN).replace(",", "").strip()
    return to_int(s) if s else None


def clean_date(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    s = str(value or "").translate(BN_TO_EN).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def current_user():
    uid = session.get("uid")
    if uid is None:
        return None
    if "user" not in g:
        g.user = get_db().execute("SELECT * FROM users WHERE id = ?", (uid,)).fetchone()
    return g.user


def login_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        if current_user() is None:
            return redirect(url_for("login", next=request.path))
        if current_user()["must_change"] and request.endpoint not in ("change_password", "logout"):
            flash("নিরাপত্তার জন্য প্রথমবার লগইনের পর পাসওয়ার্ড পরিবর্তন করুন।", "warn")
            return redirect(url_for("change_password"))
        return view(*a, **kw)
    return wrapped


def admin_required(view):
    @wraps(view)
    @login_required
    def wrapped(*a, **kw):
        if current_user()["role"] != "admin":
            abort(403)
        return view(*a, **kw)
    return wrapped


def xlsx_response(buf, filename):
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


EMPLOYEE_TEXT_FIELDS = ["code", "name_bn", "name_en", "father_name", "mother_name", "gender", "nid", "mobile",
                        "email", "address", "designation", "department", "employment_type", "status", "gpf_no", "tin",
                        "bank_name", "bank_branch", "account_no"]
EMPLOYEE_DATE_FIELDS = ["dob", "join_date", "prl_date"]
OVERRIDE_FIELDS = ["ovr_house_rent", "ovr_medical", "ovr_education", "ovr_tiffin", "ovr_transport"]


def employee_from_form(form):
    data = {f: (form.get(f) or "").strip() or None for f in EMPLOYEE_TEXT_FIELDS}
    data.update({f: clean_date(form.get(f)) for f in EMPLOYEE_DATE_FIELDS})
    data["status"] = data["status"] or "active"
    data["grade"] = to_int(form.get("grade"), None)
    data["basic"] = to_int(form.get("basic"))
    data["children"] = to_int(form.get("children"))
    data["govt_quarter"] = 1 if form.get("govt_quarter") else 0
    data["gpf_pct"] = to_float(form.get("gpf_pct"), 0)
    data.update({f: optional_int(form.get(f)) for f in OVERRIDE_FIELDS})
    return data


def save_employee(data, emp_id=None):
    conn = get_db()
    cols = list(data.keys())
    if emp_id:
        conn.execute(f"UPDATE employees SET {', '.join(c + ' = ?' for c in cols)} WHERE id = ?",
                     [data[c] for c in cols] + [emp_id])
        conn.execute("UPDATE users SET username = ? WHERE employee_id = ?", (data["code"], emp_id))
    else:
        cur = conn.execute(f"INSERT INTO employees ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                           [data[c] for c in cols])
        emp_id = cur.lastrowid
        # প্রাথমিক পাসওয়ার্ড = কর্মচারী আইডি; প্রথম লগইনে পরিবর্তন বাধ্যতামূলক
        conn.execute("INSERT INTO users (username, password_hash, role, employee_id, must_change) "
                     "VALUES (?, ?, 'employee', ?, 1)", (data["code"], generate_password_hash(data["code"]), emp_id))
    return emp_id


def basic_range_warning(grade, basic):
    pg = get_pay_grades().get(grade)
    if pg and not (pg["start_basic"] <= basic <= pg["max_basic"]):
        flash(f"সতর্কতা: গ্রেড-{bn_num(grade)} এর মূল বেতনের পরিসর {bn_taka(pg['start_basic'])}–"
              f"{bn_taka(pg['max_basic'])} টাকা, কিন্তু দেওয়া হয়েছে {bn_taka(basic)} টাকা।", "warn")


def run_totals():
    return get_db().execute(
        """SELECT r.*, COUNT(p.id) AS headcount, COALESCE(SUM(p.gross),0) AS gross,
                  COALESCE(SUM(p.total_deduction),0) AS deduction, COALESCE(SUM(p.net),0) AS net
           FROM payroll_runs r LEFT JOIN payslips p ON p.run_id = r.id
           GROUP BY r.id ORDER BY r.year, r.month""").fetchall()


def get_run_or_404(run_id):
    run = get_db().execute("SELECT * FROM payroll_runs WHERE id = ?", (run_id,)).fetchone()
    if run is None:
        abort(404)
    return run


def require_draft(run):
    if run["status"] != "draft":
        flash("চূড়ান্ত বেতন শিট পরিবর্তন করা যায় না। প্রয়োজনে আগে 'খসড়ায় ফেরত' দিন।", "error")
        return False
    return True


def employee_slips(emp_id, final_only):
    sql = ("SELECT p.*, r.year, r.month, r.status FROM payslips p JOIN payroll_runs r ON r.id = p.run_id "
           "WHERE p.employee_id = ?" + (" AND r.status = 'final'" if final_only else "") +
           " ORDER BY r.year DESC, r.month DESC")
    return get_db().execute(sql, (emp_id,)).fetchall()


# ---------------------------------------------------------------- রাউট

def register_routes(app):

    @app.before_request
    def csrf_protect():
        if "csrf" not in session:
            session["csrf"] = secrets.token_hex(16)
        if request.method == "POST" and request.form.get("csrf") != session["csrf"]:
            abort(400, "অবৈধ অনুরোধ (CSRF)। পাতাটি রিফ্রেশ করে আবার চেষ্টা করুন।")

    @app.context_processor
    def inject():
        return {"user": current_user(), "csrf": session.get("csrf", ""), "settings": get_settings()}

    # ---------- লগইন

    @app.route("/")
    def index():
        user = current_user()
        if user is None:
            return redirect(url_for("login"))
        return redirect(url_for("admin_dashboard" if user["role"] == "admin" else "my_dashboard"))

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            key = (request.remote_addr, username.lower())
            now = time.time()
            fails = [t for t in FAILED_LOGINS.get(key, []) if now - t < LOCKOUT_SECONDS]
            if len(fails) >= MAX_FAILED_LOGINS:
                flash("বারবার ভুল পাসওয়ার্ডের কারণে ১৫ মিনিটের জন্য লগইন বন্ধ। পরে চেষ্টা করুন।", "error")
                return render_template("login.html"), 429
            user = get_db().execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
            if user and check_password_hash(user["password_hash"], request.form.get("password", "")):
                FAILED_LOGINS.pop(key, None)
                if user["employee_id"]:
                    emp = get_db().execute("SELECT status FROM employees WHERE id = ?", (user["employee_id"],)).fetchone()
                    if emp and emp["status"] != "active":
                        flash("এই অ্যাকাউন্টটি নিষ্ক্রিয়। প্রশাসনের সাথে যোগাযোগ করুন।", "error")
                        return render_template("login.html")
                session.clear()
                session.permanent = True
                session["uid"] = user["id"]
                nxt = request.args.get("next", "")
                return redirect(nxt if nxt.startswith("/") and not nxt.startswith("//") else url_for("index"))
            FAILED_LOGINS[key] = fails + [now]
            flash("আইডি অথবা পাসওয়ার্ড সঠিক নয়।", "error")
        return render_template("login.html")

    @app.route("/logout", methods=["POST"])
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.route("/password", methods=["GET", "POST"])
    @login_required
    def change_password():
        if request.method == "POST":
            user = current_user()
            new = request.form.get("new", "")
            if not check_password_hash(user["password_hash"], request.form.get("old", "")):
                flash("বর্তমান পাসওয়ার্ড সঠিক নয়।", "error")
            elif len(new) < 6:
                flash("নতুন পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের হতে হবে।", "error")
            elif new != request.form.get("confirm"):
                flash("নতুন পাসওয়ার্ড দুইবার একই হয়নি।", "error")
            elif new == user["username"]:
                flash("আইডি আর পাসওয়ার্ড একই রাখা যাবে না।", "error")
            else:
                get_db().execute("UPDATE users SET password_hash = ?, must_change = 0 WHERE id = ?",
                                 (generate_password_hash(new), user["id"]))
                get_db().commit()
                flash("পাসওয়ার্ড পরিবর্তন হয়েছে।", "ok")
                return redirect(url_for("index"))
        return render_template("password.html")

    # ---------- অ্যাডমিন ড্যাশবোর্ড

    @app.route("/admin")
    @admin_required
    def admin_dashboard():
        conn = get_db()
        runs = run_totals()
        stats = conn.execute(
            "SELECT COUNT(*) AS total, SUM(status = 'active') AS active, "
            "COALESCE(SUM(CASE WHEN status = 'active' THEN basic END), 0) AS basic_sum FROM employees").fetchone()
        depts = conn.execute(
            "SELECT COALESCE(department, '—') AS department, COUNT(*) AS n FROM employees WHERE status = 'active' "
            "GROUP BY department ORDER BY n DESC").fetchall()
        today = date.today()
        this_run = conn.execute("SELECT * FROM payroll_runs WHERE year = ? AND month = ?",
                                (today.year, today.month)).fetchone()
        return render_template("admin/dashboard.html", stats=stats, runs=runs[-12:], last=runs[-1] if runs else None,
                               prev=runs[-2] if len(runs) > 1 else None, depts=depts, this_run=this_run, today=today)

    # ---------- কর্মচারী

    @app.route("/admin/employees")
    @admin_required
    def employees():
        q = request.args.get("q", "").strip()
        dept = request.args.get("dept", "")
        status = request.args.get("status", "active")
        sql, args = "SELECT * FROM employees WHERE 1=1", []
        if q:
            sql += " AND (code LIKE ? OR name_bn LIKE ? OR name_en LIKE ? OR designation LIKE ? OR mobile LIKE ?)"
            args += [f"%{q}%"] * 5
        if dept:
            sql += " AND department = ?"
            args.append(dept)
        if status:
            sql += " AND status = ?"
            args.append(status)
        rows = get_db().execute(sql + " ORDER BY grade, code", args).fetchall()
        depts = [r[0] for r in get_db().execute(
            "SELECT DISTINCT department FROM employees WHERE department IS NOT NULL ORDER BY department")]
        return render_template("admin/employees.html", rows=rows, q=q, dept=dept, status=status, depts=depts)

    def employee_form_context(emp=None):
        conn = get_db()
        depts = [r[0] for r in conn.execute(
            "SELECT DISTINCT department FROM employees WHERE department IS NOT NULL ORDER BY department")]
        desigs = [r[0] for r in conn.execute(
            "SELECT DISTINCT designation FROM employees WHERE designation IS NOT NULL ORDER BY designation")]
        return {"emp": emp, "grades": get_pay_grades(), "depts": depts, "desigs": desigs}

    @app.route("/admin/employees/new", methods=["GET", "POST"])
    @admin_required
    def employee_new():
        if request.method == "POST":
            data = employee_from_form(request.form)
            if not data["code"] or not data["name_bn"]:
                flash("আইডি ও নাম অবশ্যই দিতে হবে।", "error")
            elif get_db().execute("SELECT 1 FROM users WHERE username = ?", (data["code"],)).fetchone():
                flash("এই আইডি আগে থেকেই আছে।", "error")
            else:
                emp_id = save_employee(data)
                get_db().commit()
                basic_range_warning(data["grade"], data["basic"])
                flash(f"কর্মচারী যুক্ত হয়েছে। লগইন আইডি ও প্রাথমিক পাসওয়ার্ড: {data['code']}", "ok")
                return redirect(url_for("employee_view", emp_id=emp_id))
            return render_template("admin/employee_form.html", **employee_form_context(data))
        return render_template("admin/employee_form.html", **employee_form_context())

    @app.route("/admin/employees/<int:emp_id>/edit", methods=["GET", "POST"])
    @admin_required
    def employee_edit(emp_id):
        emp = get_db().execute("SELECT * FROM employees WHERE id = ?", (emp_id,)).fetchone() or abort(404)
        if request.method == "POST":
            data = employee_from_form(request.form)
            clash = get_db().execute("SELECT 1 FROM users WHERE username = ? AND COALESCE(employee_id, 0) != ?",
                                     (data["code"], emp_id)).fetchone()
            if not data["code"] or not data["name_bn"]:
                flash("আইডি ও নাম অবশ্যই দিতে হবে।", "error")
            elif clash:
                flash("এই আইডি অন্য কারো নামে আছে।", "error")
            else:
                save_employee(data, emp_id)
                get_db().commit()
                basic_range_warning(data["grade"], data["basic"])
                flash("তথ্য হালনাগাদ হয়েছে।", "ok")
                return redirect(url_for("employee_view", emp_id=emp_id))
            data["id"] = emp_id
            return render_template("admin/employee_form.html", **employee_form_context(data))
        return render_template("admin/employee_form.html", **employee_form_context(emp))

    @app.route("/admin/employees/<int:emp_id>")
    @admin_required
    def employee_view(emp_id):
        conn = get_db()
        emp = conn.execute("SELECT * FROM employees WHERE id = ?", (emp_id,)).fetchone() or abort(404)
        items = conn.execute("SELECT * FROM employee_items WHERE employee_id = ? ORDER BY kind, id", (emp_id,)).fetchall()
        account = conn.execute("SELECT * FROM users WHERE employee_id = ?", (emp_id,)).fetchone()
        today = date.today()
        preview = compute_components(emp, get_settings(), today.year, today.month)
        return render_template("admin/employee_view.html", emp=emp, items=items, account=account, preview=preview,
                               slips=employee_slips(emp_id, final_only=False), age=age_on(emp["dob"], today))

    @app.route("/admin/employees/<int:emp_id>/items", methods=["POST"])
    @admin_required
    def employee_item_add(emp_id):
        name = request.form.get("name", "").strip()
        kind = request.form.get("kind")
        if name and kind in ("allowance", "deduction"):
            get_db().execute("INSERT INTO employee_items (employee_id, kind, name, amount) VALUES (?, ?, ?, ?)",
                             (emp_id, kind, name, to_int(request.form.get("amount"))))
            get_db().commit()
            flash("যুক্ত হয়েছে।", "ok")
        return redirect(url_for("employee_view", emp_id=emp_id) + "#salary")

    @app.route("/admin/items/<int:item_id>/delete", methods=["POST"])
    @admin_required
    def employee_item_delete(item_id):
        item = get_db().execute("SELECT * FROM employee_items WHERE id = ?", (item_id,)).fetchone() or abort(404)
        get_db().execute("DELETE FROM employee_items WHERE id = ?", (item_id,))
        get_db().commit()
        return redirect(url_for("employee_view", emp_id=item["employee_id"]) + "#salary")

    @app.route("/admin/employees/<int:emp_id>/reset-password", methods=["POST"])
    @admin_required
    def employee_reset_password(emp_id):
        emp = get_db().execute("SELECT code FROM employees WHERE id = ?", (emp_id,)).fetchone() or abort(404)
        get_db().execute("UPDATE users SET password_hash = ?, must_change = 1 WHERE employee_id = ?",
                         (generate_password_hash(emp["code"]), emp_id))
        get_db().commit()
        flash(f"পাসওয়ার্ড রিসেট হয়েছে। নতুন প্রাথমিক পাসওয়ার্ড: {emp['code']}", "ok")
        return redirect(url_for("employee_view", emp_id=emp_id))

    @app.route("/admin/employees/<int:emp_id>/history.xlsx")
    @admin_required
    def employee_history_xlsx(emp_id):
        emp = get_db().execute("SELECT * FROM employees WHERE id = ?", (emp_id,)).fetchone() or abort(404)
        slips = list(reversed(employee_slips(emp_id, final_only=False)))
        return xlsx_response(excel.employee_history_workbook(emp, slips, get_settings()), f"salary-{emp['code']}.xlsx")

    @app.route("/admin/employees/export.xlsx")
    @admin_required
    def employees_export():
        rows = get_db().execute("SELECT * FROM employees ORDER BY grade, code").fetchall()
        return xlsx_response(excel.employees_workbook(rows, get_settings()), "employees.xlsx")

    @app.route("/admin/employees/template.xlsx")
    @admin_required
    def employees_template():
        return xlsx_response(excel.employees_workbook([], get_settings(), template=True), "employee-import-template.xlsx")

    @app.route("/admin/employees/import", methods=["GET", "POST"])
    @admin_required
    def employees_import():
        if request.method == "POST":
            f = request.files.get("file")
            if not f or not f.filename.lower().endswith(".xlsx"):
                flash("একটি .xlsx ফাইল নির্বাচন করুন।", "error")
                return redirect(request.url)
            try:
                ws = load_workbook(f, data_only=True).active
            except Exception:
                flash("ফাইলটি পড়া যায়নি। টেমপ্লেট অনুযায়ী .xlsx ফাইল দিন।", "error")
                return redirect(request.url)
            header_to_field = {h.rstrip("*").strip(): fld for fld, h in excel.EMPLOYEE_COLUMNS}
            header = [header_to_field.get(str(c.value or "").rstrip("*").strip()) for c in ws[1]]
            if "code" not in header or "name_bn" not in header:
                flash("প্রথম সারিতে টেমপ্লেটের শিরোনাম (আইডি, নাম) পাওয়া যায়নি।", "error")
                return redirect(request.url)
            conn = get_db()
            added = updated = 0
            errors = []
            for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
                raw = {fld: v for fld, v in zip(header, row) if fld}
                if not any(v not in (None, "") for v in raw.values()):
                    continue
                code = str(raw.get("code") or "").strip()
                if not code or not raw.get("name_bn"):
                    errors.append(f"সারি {bn_num(r_idx)}: আইডি/নাম নেই")
                    continue
                existing = conn.execute("SELECT * FROM employees WHERE code = ?", (code,)).fetchone()
                form = dict(existing) if existing else {}
                form.update({k: ("" if v is None else v) for k, v in raw.items()})
                for k in OVERRIDE_FIELDS:
                    form[k] = "" if form.get(k) is None else form[k]
                form["govt_quarter"] = "1" if to_int(form.get("govt_quarter")) else ""
                if "gpf_pct" not in raw and not existing:
                    form["gpf_pct"] = 5
                data = employee_from_form({k: (v if isinstance(v, (date, datetime)) else str(v))
                                           for k, v in form.items()})
                data.update({f: clean_date(form.get(f)) for f in EMPLOYEE_DATE_FIELDS})
                if existing:
                    save_employee(data, existing["id"])
                    updated += 1
                elif conn.execute("SELECT 1 FROM users WHERE username = ?", (code,)).fetchone():
                    errors.append(f"সারি {bn_num(r_idx)}: আইডি '{code}' ব্যবহারকারী নাম হিসেবে আগেই আছে")
                else:
                    save_employee(data)
                    added += 1
            conn.commit()
            flash(f"ইমপোর্ট সম্পন্ন: নতুন {bn_num(added)} জন, হালনাগাদ {bn_num(updated)} জন।", "ok")
            for e in errors[:20]:
                flash(e, "warn")
            return redirect(url_for("employees"))
        return render_template("admin/import.html")

    # ---------- বেতন স্কেল ও ভাতা

    @app.route("/admin/payscale", methods=["GET", "POST"])
    @admin_required
    def payscale():
        conn = get_db()
        if request.method == "POST":
            for grade in range(1, 21):
                conn.execute("UPDATE pay_grades SET start_basic = ?, max_basic = ? WHERE grade = ?",
                             (to_int(request.form.get(f"start_{grade}")), to_int(request.form.get(f"max_{grade}")), grade))
            bands = []
            for i in range(4):
                bands.append({"from": to_int(request.form.get(f"hr_from_{i}")), "to": to_int(request.form.get(f"hr_to_{i}")),
                              "pct": to_float(request.form.get(f"hr_pct_{i}")),
                              "min": to_int(request.form.get(f"hr_min_{i}"))})
            save_setting("house_rent_bands", bands)
            for key in ["medical_below_50", "medical_50_plus", "education_per_child", "education_max_children",
                        "tiffin_amount", "tiffin_from_grade", "transport_amount", "transport_from_grade",
                        "revenue_stamp"]:
                save_setting(key, to_int(request.form.get(key)))
            for key in ["festival_pct", "boishakhi_pct", "increment_pct"]:
                save_setting(key, to_float(request.form.get(key)))
            for key in ["office_name", "office_address"]:
                save_setting(key, request.form.get(key, "").strip())
            conn.commit()
            flash("বেতন স্কেল ও ভাতার নিয়ম সংরক্ষিত হয়েছে। নতুন বেতন শিট এই নিয়মে তৈরি হবে।", "ok")
            return redirect(url_for("payscale"))
        return render_template("admin/payscale.html", grades=get_pay_grades())

    @app.route("/admin/increment", methods=["POST"])
    @admin_required
    def annual_increment():
        """সক্রিয় সকল কর্মচারীর মূল বেতনে বার্ষিক ইনক্রিমেন্ট (গ্রেডের সর্বোচ্চ সীমা পর্যন্ত, ১০ টাকায় পূর্ণ)।"""
        conn = get_db()
        pct = float(get_settings()["increment_pct"])
        grades = get_pay_grades()
        changed = capped = 0
        for emp in conn.execute("SELECT id, grade, basic FROM employees WHERE status = 'active'").fetchall():
            pg = grades.get(emp["grade"])
            new = int(math.ceil(emp["basic"] * (1 + pct / 100) / 10) * 10)
            if pg:
                if new >= pg["max_basic"]:
                    capped += 1
                new = min(new, pg["max_basic"])
            if new != emp["basic"]:
                conn.execute("UPDATE employees SET basic = ? WHERE id = ?", (new, emp["id"]))
                changed += 1
        conn.commit()
        flash(f"{bn_num(pct)}% বার্ষিক ইনক্রিমেন্ট দেওয়া হয়েছে: {bn_num(changed)} জনের মূল বেতন বেড়েছে, "
              f"{bn_num(capped)} জন গ্রেডের সর্বোচ্চ ধাপে।", "ok")
        return redirect(url_for("payscale"))

    # ---------- মাসিক বেতন শিট

    @app.route("/admin/payroll")
    @admin_required
    def payroll_list():
        runs = list(reversed(run_totals()))
        today = date.today()
        return render_template("admin/payroll_list.html", runs=runs, today=today)

    @app.route("/admin/payroll/new", methods=["POST"])
    @admin_required
    def payroll_new():
        year, month = to_int(request.form.get("year")), to_int(request.form.get("month"))
        if not (2000 <= year <= 2100 and 1 <= month <= 12):
            flash("সঠিক মাস ও বছর দিন।", "error")
            return redirect(url_for("payroll_list"))
        existing = get_db().execute("SELECT id FROM payroll_runs WHERE year = ? AND month = ?", (year, month)).fetchone()
        if existing:
            flash(f"{month_label(year, month)} এর বেতন শিট আগেই তৈরি হয়েছে।", "warn")
            return redirect(url_for("payroll_view", run_id=existing["id"]))
        run_id = generate_run(year, month, bool(request.form.get("festival")), bool(request.form.get("boishakhi")),
                              request.form.get("note", "").strip())
        flash(f"{month_label(year, month)} এর খসড়া বেতন শিট তৈরি হয়েছে। যাচাই করে 'চূড়ান্ত করুন' চাপুন।", "ok")
        return redirect(url_for("payroll_view", run_id=run_id))

    @app.route("/admin/payroll/<int:run_id>")
    @admin_required
    def payroll_view(run_id):
        run = get_run_or_404(run_id)
        slips = get_db().execute("SELECT * FROM payslips WHERE run_id = ? ORDER BY grade, code", (run_id,)).fetchall()
        totals = {f: sum(s[f] for s in slips) for f in ALLOWANCE_FIELDS + DEDUCTION_FIELDS + ["gross", "total_deduction", "net"]}
        missing = get_db().execute(
            "SELECT * FROM employees WHERE status = 'active' AND id NOT IN (SELECT employee_id FROM payslips WHERE run_id = ?)",
            (run_id,)).fetchall()
        prev = get_db().execute(
            "SELECT COALESCE(SUM(p.net),0) AS net FROM payroll_runs r JOIN payslips p ON p.run_id = r.id "
            "WHERE r.year * 12 + r.month = ? ", (run["year"] * 12 + run["month"] - 1,)).fetchone()
        return render_template("admin/payroll_view.html", run=run, slips=slips, totals=totals, missing=missing,
                               prev_net=prev["net"] if prev else 0)

    @app.route("/admin/payroll/<int:run_id>/excel")
    @admin_required
    def payroll_excel(run_id):
        run = get_run_or_404(run_id)
        slips = get_db().execute("SELECT * FROM payslips WHERE run_id = ? ORDER BY grade, code", (run_id,)).fetchall()
        return xlsx_response(excel.payroll_workbook(run, slips, get_settings()),
                             f"salary-sheet-{run['year']}-{run['month']:02d}.xlsx")

    @app.route("/admin/payroll/<int:run_id>/finalize", methods=["POST"])
    @admin_required
    def payroll_finalize(run_id):
        get_run_or_404(run_id)
        get_db().execute("UPDATE payroll_runs SET status = 'final', finalized_at = CURRENT_TIMESTAMP WHERE id = ?", (run_id,))
        get_db().commit()
        flash("বেতন শিট চূড়ান্ত হয়েছে। কর্মচারীরা এখন নিজ ড্যাশবোর্ডে বেতন বিবরণী দেখতে পাবেন।", "ok")
        return redirect(url_for("payroll_view", run_id=run_id))

    @app.route("/admin/payroll/<int:run_id>/reopen", methods=["POST"])
    @admin_required
    def payroll_reopen(run_id):
        get_run_or_404(run_id)
        get_db().execute("UPDATE payroll_runs SET status = 'draft', finalized_at = NULL WHERE id = ?", (run_id,))
        get_db().commit()
        flash("বেতন শিট আবার খসড়া অবস্থায় ফেরত গেছে।", "warn")
        return redirect(url_for("payroll_view", run_id=run_id))

    @app.route("/admin/payroll/<int:run_id>/delete", methods=["POST"])
    @admin_required
    def payroll_delete(run_id):
        run = get_run_or_404(run_id)
        if require_draft(run):
            get_db().execute("DELETE FROM payroll_runs WHERE id = ?", (run_id,))
            get_db().commit()
            flash(f"{month_label(run['year'], run['month'])} এর খসড়া বেতন শিট মুছে ফেলা হয়েছে।", "ok")
            return redirect(url_for("payroll_list"))
        return redirect(url_for("payroll_view", run_id=run_id))

    @app.route("/admin/payroll/<int:run_id>/recalc", methods=["POST"])
    @admin_required
    def payroll_recalc(run_id):
        """খসড়া শিটের সকলের বেতন বর্তমান তথ্য অনুযায়ী আবার হিসাব + নতুন যুক্ত কর্মচারী অন্তর্ভুক্ত।"""
        run = get_run_or_404(run_id)
        if require_draft(run):
            conn = get_db()
            for slip in conn.execute("SELECT * FROM payslips WHERE run_id = ?", (run_id,)).fetchall():
                regenerate_slip(run, slip)
            settings = get_settings()
            for emp in conn.execute("SELECT * FROM employees WHERE status = 'active' AND id NOT IN "
                                    "(SELECT employee_id FROM payslips WHERE run_id = ?)", (run_id,)).fetchall():
                insert_slip(run_id, emp, compute_components(emp, settings, run["year"], run["month"],
                                                            run["include_festival"], run["include_boishakhi"]))
            conn.commit()
            flash("সকলের বেতন বর্তমান তথ্য অনুযায়ী পুনরায় হিসাব করা হয়েছে।", "ok")
        return redirect(url_for("payroll_view", run_id=run_id))

    @app.route("/admin/payroll/<int:run_id>/slip/<int:slip_id>/remove", methods=["POST"])
    @admin_required
    def payslip_remove(run_id, slip_id):
        run = get_run_or_404(run_id)
        if require_draft(run):
            get_db().execute("DELETE FROM payslips WHERE id = ? AND run_id = ?", (slip_id, run_id))
            get_db().commit()
            flash("এই মাসের শিট থেকে বাদ দেওয়া হয়েছে।", "ok")
        return redirect(url_for("payroll_view", run_id=run_id))

    @app.route("/admin/payroll/<int:run_id>/slip/<int:slip_id>/edit", methods=["GET", "POST"])
    @admin_required
    def payslip_edit(run_id, slip_id):
        run = get_run_or_404(run_id)
        slip = get_db().execute("SELECT * FROM payslips WHERE id = ? AND run_id = ?", (slip_id, run_id)).fetchone() or abort(404)
        if request.method == "POST":
            if not require_draft(run):
                return redirect(url_for("payroll_view", run_id=run_id))
            if request.form.get("action") == "recalc":
                regenerate_slip(run, slip)
                get_db().commit()
                flash("কর্মচারীর বর্তমান তথ্য অনুযায়ী পুনরায় হিসাব করা হয়েছে।", "ok")
                return redirect(url_for("payroll_view", run_id=run_id))
            s = {f: to_int(request.form.get(f)) for f in ALLOWANCE_FIELDS + DEDUCTION_FIELDS}
            for kind in ("other_allowance", "other_deduction"):
                names, amounts = request.form.getlist(f"{kind}_name"), request.form.getlist(f"{kind}_amount")
                s[kind] = [{"name": n.strip(), "amount": to_int(a)} for n, a in zip(names, amounts) if n.strip()]
            finalize_totals(s)
            get_db().execute(
                f"UPDATE payslips SET {', '.join(f + ' = ?' for f in ALLOWANCE_FIELDS + DEDUCTION_FIELDS)}, "
                "other_allowance = ?, other_deduction = ?, gross = ?, total_deduction = ?, net = ? WHERE id = ?",
                [s[f] for f in ALLOWANCE_FIELDS + DEDUCTION_FIELDS] +
                [json.dumps(s["other_allowance"], ensure_ascii=False), json.dumps(s["other_deduction"], ensure_ascii=False),
                 s["gross"], s["total_deduction"], s["net"], slip_id])
            get_db().commit()
            flash(f"{slip['name_bn']} এর এ মাসের বেতন সংশোধন হয়েছে।", "ok")
            return redirect(url_for("payroll_view", run_id=run_id))
        return render_template("admin/payslip_edit.html", run=run, slip=slip, allowance_fields=ALLOWANCE_FIELDS,
                               deduction_fields=DEDUCTION_FIELDS)

    # ---------- রিপোর্ট

    @app.route("/admin/reports")
    @admin_required
    def reports():
        runs = run_totals()
        rows, prev = [], None
        for r in runs:
            rows.append({"run": r, "diff": (r["net"] - prev["net"]) if prev else None,
                         "hdiff": (r["headcount"] - prev["headcount"]) if prev else None})
            prev = r
        years = {}
        for r in runs:
            y = years.setdefault(r["year"], {"gross": 0, "deduction": 0, "net": 0, "months": 0})
            y["gross"] += r["gross"]
            y["deduction"] += r["deduction"]
            y["net"] += r["net"]
            y["months"] += 1
        return render_template("admin/reports.html", rows=list(reversed(rows)), years=years,
                               chart=runs[-12:])

    @app.route("/admin/reports/excel")
    @admin_required
    def reports_excel():
        return xlsx_response(excel.history_workbook(run_totals(), get_settings()), "monthly-salary-summary.xlsx")

    # ---------- কর্মচারীর নিজস্ব ড্যাশবোর্ড

    def my_employee():
        user = current_user()
        if user["role"] != "employee" or not user["employee_id"]:
            abort(403)
        return get_db().execute("SELECT * FROM employees WHERE id = ?", (user["employee_id"],)).fetchone()

    @app.route("/me")
    @login_required
    def my_dashboard():
        emp = my_employee()
        slips = employee_slips(emp["id"], final_only=True)
        year = date.today().year
        ytd = sum(s["net"] for s in slips if s["year"] == year)
        return render_template("me/dashboard.html", emp=emp, slips=slips, latest=slips[0] if slips else None,
                               chart=list(reversed(slips[:12])), ytd=ytd, year=year)

    @app.route("/me/profile")
    @login_required
    def my_profile():
        emp = my_employee()
        return render_template("me/profile.html", emp=emp, age=age_on(emp["dob"], date.today()))

    @app.route("/me/history.xlsx")
    @login_required
    def my_history_xlsx():
        emp = my_employee()
        slips = list(reversed(employee_slips(emp["id"], final_only=True)))
        return xlsx_response(excel.employee_history_workbook(emp, slips, get_settings()), f"salary-{emp['code']}.xlsx")

    @app.route("/payslip/<int:slip_id>")
    @login_required
    def payslip(slip_id):
        slip = get_db().execute(
            "SELECT p.*, r.year, r.month, r.status FROM payslips p JOIN payroll_runs r ON r.id = p.run_id WHERE p.id = ?",
            (slip_id,)).fetchone() or abort(404)
        user = current_user()
        if user["role"] != "admin" and (slip["employee_id"] != user["employee_id"] or slip["status"] != "final"):
            abort(403)
        emp = get_db().execute("SELECT * FROM employees WHERE id = ?", (slip["employee_id"],)).fetchone()
        return render_template("payslip.html", slip=slip, emp=emp)

    @app.errorhandler(403)
    def forbidden(_e):
        return render_template("error.html", code=403, message="এই পাতাটি দেখার অনুমতি আপনার নেই।"), 403

    @app.errorhandler(404)
    def not_found(_e):
        return render_template("error.html", code=404, message="পাতাটি পাওয়া যায়নি।"), 404


app = create_app()

if __name__ == "__main__":
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", 5000)),
            debug=os.environ.get("FLASK_DEBUG") == "1")
