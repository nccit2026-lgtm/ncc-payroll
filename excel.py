"""Excel (.xlsx) ফাইল তৈরি — মাসিক বেতন শিট, ব্যাংক অ্যাডভাইস, মাসভিত্তিক সারাংশ, কর্মচারী ইমপোর্ট টেমপ্লেট।"""
import io
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from bn import MONTHS_BN, bn_num

FONT = "Nirmala UI"
HEAD_FILL = PatternFill("solid", fgColor="1F5F4A")
TOTAL_FILL = PatternFill("solid", fgColor="E6F0EC")
THIN = Side(style="thin", color="B7C4BE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0"


def _title(ws, office, title, ncols):
    for row, (text, size) in enumerate([(office["office_name"], 15), (office["office_address"], 11), (title, 13)], 1):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name=FONT, size=size, bold=row != 2)
        c.alignment = Alignment(horizontal="center")


def _header(ws, row, headers):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 36


def _style_body(ws, first, last, ncols, money_from):
    for r in range(first, last + 1):
        for col in range(1, ncols + 1):
            c = ws.cell(row=r, column=col)
            c.font = Font(name=FONT, bold=c.font.bold)
            c.border = BORDER
            if col >= money_from:
                c.number_format = MONEY


def _totals_row(ws, row, first, last, label_col, money_cols):
    ws.cell(row=row, column=label_col, value="সর্বমোট").font = Font(name=FONT, bold=True)
    for col in range(1, ws.max_column + 1):
        ws.cell(row=row, column=col).fill = TOTAL_FILL
    for col in money_cols:
        L = get_column_letter(col)
        c = ws.cell(row=row, column=col, value=f"=SUM({L}{first}:{L}{last})")
        c.font = Font(name=FONT, bold=True)


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _save(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def payroll_workbook(run, slips, settings):
    wb = Workbook()
    ws = wb.active
    ws.title = "বেতন শিট"
    period = f"{MONTHS_BN[run['month']]} {bn_num(run['year'])}"

    allowance_cols = [
        ("basic", "মূল বেতন"), ("house_rent", "বাড়ি ভাড়া"), ("medical", "চিকিৎসা"),
        ("education", "শিক্ষা সহায়ক"), ("tiffin", "টিফিন"), ("transport", "যাতায়াত"),
    ]
    if run["include_festival"]:
        allowance_cols.append(("festival", "উৎসব ভাতা"))
    if run["include_boishakhi"]:
        allowance_cols.append(("boishakhi", "নববর্ষ ভাতা"))

    info = ["ক্রম", "আইডি", "নাম", "পদবি", "শাখা", "গ্রেড", "হিসাব নং"]
    headers = (info + [h for _, h in allowance_cols] + ["অন্যান্য ভাতা", "মোট বেতন", "জিপিএফ", "রাজস্ব স্ট্যাম্প",
               "অন্যান্য কর্তন", "মোট কর্তন", "নিট প্রদেয়"])
    n = len(headers)
    _title(ws, settings, f"মাসিক বেতন শিট — {period}", n)
    hr = 5
    _header(ws, hr, headers)

    r = hr
    for i, s in enumerate(slips, 1):
        r += 1
        oa = sum(int(x["amount"]) for x in json.loads(s["other_allowance"]))
        od = sum(int(x["amount"]) for x in json.loads(s["other_deduction"]))
        values = ([i, s["code"], s["name_bn"], s["designation"], s["department"], s["grade"], s["account_no"]]
                  + [s[k] for k, _ in allowance_cols]
                  + [oa, s["gross"], s["gpf"], s["revenue_stamp"], od, s["total_deduction"], s["net"]])
        for col, v in enumerate(values, 1):
            ws.cell(row=r, column=col, value=v)
    first, last = hr + 1, max(r, hr + 1)
    money_from = len(info) + 1
    _style_body(ws, first, last, n, money_from)
    _totals_row(ws, last + 1, first, last, 3, range(money_from, n + 1))
    ws.freeze_panes = ws.cell(row=hr + 1, column=4)
    _widths(ws, [6, 10, 24, 20, 16, 7, 18] + [12] * (n - len(info)))
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ব্যাংক অ্যাডভাইস
    bank = wb.create_sheet("ব্যাংক অ্যাডভাইস")
    bh = ["ক্রম", "আইডি", "নাম", "পদবি", "হিসাব নং", "নিট প্রদেয় (টাকা)"]
    _title(bank, settings, f"ব্যাংক অ্যাডভাইস — {period}", len(bh))
    _header(bank, hr, bh)
    for i, s in enumerate(slips, 1):
        for col, v in enumerate([i, s["code"], s["name_bn"], s["designation"], s["account_no"], s["net"]], 1):
            bank.cell(row=hr + i, column=col, value=v)
    blast = hr + max(len(slips), 1)
    _style_body(bank, hr + 1, blast, len(bh), 6)
    _totals_row(bank, blast + 1, hr + 1, blast, 3, [6])
    _widths(bank, [6, 10, 26, 22, 22, 18])

    # শাখাভিত্তিক সারাংশ
    summary = wb.create_sheet("শাখাভিত্তিক সারাংশ")
    sh = ["শাখা", "কর্মচারী সংখ্যা", "মোট বেতন", "মোট কর্তন", "নিট প্রদেয়"]
    _title(summary, settings, f"শাখাভিত্তিক সারাংশ — {period}", len(sh))
    _header(summary, hr, sh)
    by_dept = {}
    for s in slips:
        d = by_dept.setdefault(s["department"] or "—", [0, 0, 0, 0])
        d[0] += 1
        d[1] += s["gross"]
        d[2] += s["total_deduction"]
        d[3] += s["net"]
    for i, (dept, vals) in enumerate(sorted(by_dept.items()), 1):
        for col, v in enumerate([dept] + vals, 1):
            summary.cell(row=hr + i, column=col, value=v)
    slast = hr + max(len(by_dept), 1)
    _style_body(summary, hr + 1, slast, len(sh), 3)
    _totals_row(summary, slast + 1, hr + 1, slast, 1, range(2, 6))
    _widths(summary, [26, 16, 18, 16, 18])
    return _save(wb)


def history_workbook(runs, settings):
    """মাসভিত্তিক বেতন প্রদানের সারাংশ ও পূর্ববর্তী মাসের সাথে পার্থক্য।"""
    wb = Workbook()
    ws = wb.active
    ws.title = "মাসভিত্তিক সারাংশ"
    headers = ["মাস", "অবস্থা", "কর্মচারী সংখ্যা", "মোট বেতন", "মোট কর্তন", "নিট প্রদান", "আগের মাসের তুলনায় পার্থক্য"]
    _title(ws, settings, "মাসভিত্তিক বেতন প্রদানের হিসাব", len(headers))
    _header(ws, 5, headers)
    prev = None
    for i, run in enumerate(runs, 1):
        diff = run["net"] - prev if prev is not None else None
        vals = [f"{MONTHS_BN[run['month']]} {run['year']}", "চূড়ান্ত" if run["status"] == "final" else "খসড়া",
                run["headcount"], run["gross"], run["deduction"], run["net"], diff]
        for col, v in enumerate(vals, 1):
            ws.cell(row=5 + i, column=col, value=v)
        prev = run["net"]
    last = 5 + max(len(runs), 1)
    _style_body(ws, 6, last, len(headers), 4)
    _widths(ws, [18, 10, 16, 18, 16, 18, 24])
    return _save(wb)


def employee_history_workbook(emp, slips, settings):
    wb = Workbook()
    ws = wb.active
    ws.title = "বেতন ইতিহাস"
    headers = ["মাস", "মূল বেতন", "বাড়ি ভাড়া", "চিকিৎসা", "শিক্ষা", "টিফিন", "যাতায়াত", "উৎসব/নববর্ষ",
               "অন্যান্য ভাতা", "মোট বেতন", "মোট কর্তন", "নিট প্রাপ্তি"]
    _title(ws, settings, f"বেতন ইতিহাস — {emp['name_bn']} ({emp['code']})", len(headers))
    _header(ws, 5, headers)
    for i, s in enumerate(slips, 1):
        oa = sum(int(x["amount"]) for x in json.loads(s["other_allowance"]))
        vals = [f"{MONTHS_BN[s['month']]} {s['year']}", s["basic"], s["house_rent"], s["medical"], s["education"],
                s["tiffin"], s["transport"], s["festival"] + s["boishakhi"], oa, s["gross"], s["total_deduction"],
                s["net"]]
        for col, v in enumerate(vals, 1):
            ws.cell(row=5 + i, column=col, value=v)
    last = 5 + max(len(slips), 1)
    _style_body(ws, 6, last, len(headers), 2)
    _totals_row(ws, last + 1, 6, last, 1, range(2, len(headers) + 1))
    _widths(ws, [18] + [13] * (len(headers) - 1))
    return _save(wb)


# কর্মচারী ইমপোর্ট/এক্সপোর্ট কলাম: (ডাটাবেস ফিল্ড, শিরোনাম)
EMPLOYEE_COLUMNS = [
    ("code", "আইডি*"), ("name_bn", "নাম (বাংলা)*"), ("name_en", "নাম (ইংরেজি)"), ("father_name", "পিতার নাম"),
    ("mother_name", "মাতার নাম"), ("gender", "লিঙ্গ"), ("dob", "জন্ম তারিখ (YYYY-MM-DD)"), ("nid", "এনআইডি"),
    ("mobile", "মোবাইল"), ("email", "ইমেইল"), ("address", "ঠিকানা"), ("designation", "পদবি"),
    ("department", "শাখা/বিভাগ"), ("employment_type", "চাকরির ধরন"), ("join_date", "যোগদানের তারিখ (YYYY-MM-DD)"),
    ("prl_date", "পিআরএল তারিখ (YYYY-MM-DD)"), ("grade", "গ্রেড"), ("basic", "মূল বেতন"),
    ("children", "সন্তান সংখ্যা"), ("govt_quarter", "সরকারি বাসা (১=হ্যাঁ)"), ("gpf_pct", "জিপিএফ %"),
    ("gpf_no", "জিপিএফ নং"), ("tin", "টিআইএন"), ("bank_name", "ব্যাংক"), ("bank_branch", "শাখা (ব্যাংক)"),
    ("account_no", "হিসাব নং"),
]


def employees_workbook(employees, settings, template=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "কর্মচারী"
    _header(ws, 1, [h for _, h in EMPLOYEE_COLUMNS])
    if template:
        sample = {"code": "NCC-0001", "name_bn": "মোঃ রহিম উদ্দিন", "name_en": "Md. Rahim Uddin", "gender": "পুরুষ",
                  "dob": "1980-05-12", "mobile": "01700000000", "designation": "উচ্চমান সহকারী",
                  "department": "প্রশাসন", "employment_type": "স্থায়ী", "join_date": "2008-07-01", "grade": 14,
                  "basic": 23500, "children": 2, "govt_quarter": 0, "gpf_pct": 5, "bank_name": "সোনালী ব্যাংক",
                  "bank_branch": "নারায়ণগঞ্জ কর্পোরেট", "account_no": "0000000000000"}
        employees = [sample]
    for r, emp in enumerate(employees, 2):
        for col, (field, _) in enumerate(EMPLOYEE_COLUMNS, 1):
            v = emp[field] if field in emp.keys() else None
            ws.cell(row=r, column=col, value=v).font = Font(name=FONT)
    for i in range(1, len(EMPLOYEE_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.freeze_panes = "C2"
    return _save(wb)
