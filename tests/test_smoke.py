"""সম্পূর্ণ কাজের প্রবাহের স্মোক টেস্ট:  python -m pytest tests  অথবা  python tests/test_smoke.py"""
import io
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from openpyxl import load_workbook  # noqa: E402

from app import create_app  # noqa: E402
from excel import EMPLOYEE_COLUMNS  # noqa: E402


def csrf(client):
    with client.session_transaction() as s:
        return s.get("csrf")


def post(client, url, data=None, **kw):
    client.get("/login")  # নিশ্চিত করে session এ csrf আছে
    d = dict(data or {})
    d["csrf"] = csrf(client)
    return client.post(url, data=d, **kw)


def test_full_flow():
    db = os.path.join(tempfile.mkdtemp(), "t.db")
    app = create_app(db)
    c = app.test_client()

    assert c.get("/admin").status_code == 302
    r = post(c, "/login", {"username": "admin", "password": "admin123"}, follow_redirects=True)
    assert "ড্যাশবোর্ড" in r.get_data(as_text=True)

    # নতুন কর্মচারী
    r = post(c, "/admin/employees/new", {"code": "E1", "name_bn": "পরীক্ষা কর্মচারী", "grade": "14", "basic": "23500",
                                         "children": "3", "gpf_pct": "5", "dob": "1970-01-01", "status": "active"},
             follow_redirects=True)
    assert r.status_code == 200
    # বাড়ি ভাড়া 50% = 11750, চিকিৎসা ৫০+ = 4000, শিক্ষা 2x500, টিফিন 200, যাতায়াত 300
    html = r.get_data(as_text=True)
    assert "১১,৭৫০" in html and "৪,০০০" in html

    emp_id = int(re.search(r"/admin/employees/(\d+)/edit", html).group(1))
    post(c, f"/admin/employees/{emp_id}/items", {"kind": "deduction", "name": "ঋণ কিস্তি", "amount": "১০০০"})

    # Excel ইমপোর্ট
    tpl = c.get("/admin/employees/template.xlsx")
    assert tpl.status_code == 200
    wb = load_workbook(io.BytesIO(tpl.data))
    ws = wb.active
    ws.cell(row=3, column=1, value="E2")
    ws.cell(row=3, column=2, value="দ্বিতীয় কর্মচারী")
    ws.cell(row=3, column=[f for f, _ in EMPLOYEE_COLUMNS].index("grade") + 1, value=20)
    ws.cell(row=3, column=[f for f, _ in EMPLOYEE_COLUMNS].index("basic") + 1, value=20000)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    r = post(c, "/admin/employees/import", {"file": (buf, "e.xlsx")}, content_type="multipart/form-data",
             follow_redirects=True)
    assert "নতুন ২ জন" in r.get_data(as_text=True), r.get_data(as_text=True)[:3000]

    # বেতন শিট
    r = post(c, "/admin/payroll/new", {"year": "2026", "month": "10", "festival": "1"}, follow_redirects=True)
    html = r.get_data(as_text=True)
    assert "বেতন শিট — অক্টোবর ২০২৬" in html
    run_id = int(re.search(r"/admin/payroll/(\d+)/excel", html).group(1))
    slip_id = int(re.search(rf"/admin/payroll/{run_id}/slip/(\d+)/edit", html).group(1))

    with app.app_context():
        from db import get_db
        s = get_db().execute("SELECT * FROM payslips WHERE employee_id = ?", (emp_id,)).fetchone()
        assert (s["basic"], s["house_rent"], s["medical"], s["education"], s["tiffin"], s["transport"], s["festival"]) == \
            (23500, 11750, 4000, 1000, 200, 300, 23500)
        assert s["gross"] == 64250 and s["gpf"] == 1175 and s["total_deduction"] == 1175 + 10 + 1000
        assert s["net"] == 64250 - 2185

    r = c.get(f"/admin/payroll/{run_id}/excel")
    assert r.status_code == 200
    wb = load_workbook(io.BytesIO(r.data))
    assert wb.sheetnames == ["বেতন শিট", "ব্যাংক অ্যাডভাইস", "শাখাভিত্তিক সারাংশ"]

    # একজনের বেতন হাতে সংশোধন
    edit = c.get(f"/admin/payroll/{run_id}/slip/{slip_id}/edit")
    assert edit.status_code == 200
    form = {f: "0" for f in ["basic", "house_rent", "medical", "education", "tiffin", "transport", "festival",
                             "boishakhi", "gpf", "revenue_stamp"]}
    form.update({"basic": "20000", "other_allowance_name": "বকেয়া", "other_allowance_amount": "500"})
    post(c, f"/admin/payroll/{run_id}/slip/{slip_id}/edit", form)
    with app.app_context():
        s = get_db().execute("SELECT * FROM payslips WHERE id = ?", (slip_id,)).fetchone()
        assert s["gross"] == 20500 and s["net"] == 20500

    # কর্মচারী খসড়া শিট দেখতে পারবে না
    emp = app.test_client()
    post(emp, "/login", {"username": "E1", "password": "E1"})
    assert emp.get("/me").headers["Location"].endswith("/password")
    post(emp, "/password", {"old": "E1", "new": "secret99", "confirm": "secret99"})
    assert "এখনো কোনো বেতনের রেকর্ড নেই" in emp.get("/me").get_data(as_text=True)
    assert emp.get("/admin").status_code == 403

    post(c, f"/admin/payroll/{run_id}/finalize")
    html = emp.get("/me").get_data(as_text=True)
    assert "অক্টোবর ২০২৬" in html
    with app.app_context():
        own = get_db().execute("SELECT id FROM payslips WHERE employee_id = ?", (emp_id,)).fetchone()["id"]
        other = get_db().execute("SELECT id FROM payslips WHERE employee_id != ?", (emp_id,)).fetchone()["id"]
    assert emp.get(f"/payslip/{own}").status_code == 200
    assert emp.get(f"/payslip/{other}").status_code == 403
    assert emp.get("/me/history.xlsx").status_code == 200

    for url in ["/admin", "/admin/employees", f"/admin/employees/{emp_id}", f"/admin/employees/{emp_id}/edit",
                "/admin/payroll", "/admin/reports", "/admin/reports/excel", "/admin/payscale", "/admin/employees/export.xlsx",
                f"/admin/employees/{emp_id}/history.xlsx", "/admin/employees/import"]:
        assert c.get(url).status_code == 200, url

    # চূড়ান্ত শিট মোছা যাবে না
    post(c, f"/admin/payroll/{run_id}/delete")
    with app.app_context():
        assert get_db().execute("SELECT 1 FROM payroll_runs WHERE id = ?", (run_id,)).fetchone()

    # ইনক্রিমেন্ট
    post(c, "/admin/increment")
    with app.app_context():
        assert get_db().execute("SELECT basic FROM employees WHERE id = ?", (emp_id,)).fetchone()["basic"] == 24680

    # CSRF ছাড়া POST বাতিল
    assert c.post("/admin/increment").status_code == 400


if __name__ == "__main__":
    test_full_flow()
    print("সব টেস্ট সফল")
