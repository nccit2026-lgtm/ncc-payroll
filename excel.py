"""Excel (.xlsx) ফাইল তৈরি — মাসিক বেতন শিট, ব্যাংক অ্যাডভাইস, মাসভিত্তিক সারাংশ, কর্মচারী ইমপোর্ট টেমপ্লেট।"""
import io
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from bn import MONTHS_BN, bn_num
from payroll import group_by_department

FONT = "Nirmala UI"
HEAD_FILL = PatternFill("solid", fgColor="1F5F4A")
TOTAL_FILL = PatternFill("solid", fgColor="E6F0EC")
DEPT_FILL = PatternFill("solid", fgColor="F3F7F5")
GRAND_FILL = PatternFill("solid", fgColor="CFE3DA")
THIN = Side(style="thin", color="B7C4BE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0"
MONEY_PAISA = "#,##0.00"


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
                v = c.value
                c.number_format = MONEY_PAISA if isinstance(v, float) and not v.is_integer() else MONEY


def _totals_row(ws, row, first, last, label_col, money_cols, label="সর্বমোট", ncols=None):
    ws.cell(row=row, column=label_col, value=label).font = Font(name=FONT, bold=True)
    for col in range(1, (ncols or ws.max_column) + 1):
        ws.cell(row=row, column=col).fill = TOTAL_FILL
        ws.cell(row=row, column=col).border = BORDER
    for col in money_cols:
        L = get_column_letter(col)
        c = ws.cell(row=row, column=col, value=f"=SUM({L}{first}:{L}{last})")
        c.font = Font(name=FONT, bold=True)
        paisa = any(isinstance(ws.cell(row=r, column=col).value, float)
                    and not ws.cell(row=r, column=col).value.is_integer() for r in range(first, last + 1))
        c.number_format = MONEY_PAISA if paisa else MONEY


def _signature(ws, row, ncols, settings):
    """নিচে সইয়ের ঘর: বামে প্রস্তুতকারী, ডানে হিসাবরক্ষণ কর্মকর্তা (নাম ও পদবি সেটিং থেকে)।"""
    right = max(ncols - 3, 4)
    lines_left = ["……………………………", "প্রস্তুতকারী"]
    lines_right = ["……………………………", settings.get("accounts_officer_name") or "",
                   settings.get("accounts_officer_designation") or "হিসাবরক্ষণ কর্মকর্তা", settings["office_name"]]
    for i, t in enumerate(lines_left):
        ws.cell(row=row + i, column=2, value=t).font = Font(name=FONT, bold=i == 1)
    for i, t in enumerate(x for x in lines_right if x):
        ws.cell(row=row + i, column=right, value=t).font = Font(name=FONT, bold=i == 1)


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

    # যে ভাতা এ মাসে কারো নেই সেই কলাম বাদ
    all_cols = [("basic", "মূল বেতন"), ("house_rent", "বাড়িভাড়া"), ("medical", "চিকিৎসা"),
                ("education", "শিক্ষা সহায়ক"), ("tiffin", "টিফিন"), ("transport", "যাতায়াত"), ("mobile", "মোবাইল"),
                ("wash", "ধোলাই"), ("special", "বিশেষ সুবিধা (মহার্ঘ)"), ("pf_office", "অফিস প্রদত্ত PF"),
                ("festival", "উৎসব ভাতা"), ("boishakhi", "নববর্ষ ভাতা")]
    allowance_cols = [(k, h) for k, h in all_cols if k in ("basic", "house_rent", "medical") or any(s[k] for s in slips)]
    has_pf = any(s["pf_fund"] or s["pf_own"] for s in slips)
    has_stamp = any(s["revenue_stamp"] for s in slips)
    ded_cols = ([("pf_fund", "PF (অফিস অংশ জমা)"), ("pf_own", "PF (নিজ অংশ)")] if has_pf else []) + \
        ([("revenue_stamp", "রাজস্ব স্ট্যাম্প")] if has_stamp else [])

    info = ["ক্রম", "আইডি", "নাম", "পদবি", "গ্রেড", "হিসাব নং"]
    headers = (info + [h for _, h in allowance_cols] + ["অন্যান্য ভাতা / বকেয়া", "মোট বেতন"] + [h for _, h in ded_cols]
               + ["অন্যান্য কর্তন", "মোট কর্তন", "নিট প্রদেয়"])
    n = len(headers)
    money_from = len(info) + 1
    _title(ws, settings, f"মাসিক বেতন শিট — {period}", n)
    hr = 5
    _header(ws, hr, headers)

    # শাখাভিত্তিক: শাখার নাম → কর্মচারীরা → উপমোট → একটি ফাঁকা সারি
    r = hr
    subtotal_rows = []
    for dept, group in group_by_department(slips):
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
        c = ws.cell(row=r, column=1, value=f"{dept}  ({bn_num(len(group))} জন)")
        c.font = Font(name=FONT, bold=True, size=12, color="1F5F4A")
        c.fill = DEPT_FILL
        c.border = BORDER
        first = r + 1
        for i, s in enumerate(group, 1):
            r += 1
            oa = round(sum(float(x["amount"]) for x in json.loads(s["other_allowance"])), 2)
            od = round(sum(float(x["amount"]) for x in json.loads(s["other_deduction"])), 2)
            values = ([i, s["code"], s["name_bn"], s["designation"], s["grade"], s["account_no"]]
                      + [s[k] for k, _ in allowance_cols] + [oa, s["gross"]] + [s[k] for k, _ in ded_cols]
                      + [od, s["total_deduction"], s["net"]])
            for col, v in enumerate(values, 1):
                ws.cell(row=r, column=col, value=v)
        _style_body(ws, first, r, n, money_from)
        r += 1
        _totals_row(ws, r, first, r - 1, 3, range(money_from, n + 1), label="উপমোট", ncols=n)
        subtotal_rows.append(r)
        r += 1  # শাখার মাঝে বিরতি
    r += 1
    ws.cell(row=r, column=3, value=f"সর্বমোট ({bn_num(len(slips))} জন)").font = Font(name=FONT, bold=True, size=12)
    for col in range(1, n + 1):
        cell = ws.cell(row=r, column=col)
        cell.fill = GRAND_FILL
        cell.border = BORDER
        if col >= money_from and subtotal_rows:
            L = get_column_letter(col)
            cell.value = "=" + "+".join(f"{L}{x}" for x in subtotal_rows)
            cell.font = Font(name=FONT, bold=True, size=12)
            cell.number_format = MONEY_PAISA
    _signature(ws, r + 4, n, settings)
    ws.freeze_panes = f"D{hr + 1}"
    _widths(ws, [6, 11, 26, 22, 7, 18] + [12] * (n - len(info)))
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
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
    _signature(bank, blast + 5, len(bh), settings)
    _widths(bank, [6, 11, 26, 22, 22, 18])

    # শাখাভিত্তিক সারাংশ
    summary = wb.create_sheet("শাখাভিত্তিক সারাংশ")
    sh = ["শাখা", "কর্মচারী সংখ্যা", "মোট বেতন", "মোট কর্তন", "নিট প্রদেয়"]
    _title(summary, settings, f"শাখাভিত্তিক সারাংশ — {period}", len(sh))
    _header(summary, hr, sh)
    by_dept = {}
    for dept, group in group_by_department(slips):
        by_dept[dept] = [len(group), round(sum(x["gross"] for x in group), 2),
                         round(sum(x["total_deduction"] for x in group), 2), round(sum(x["net"] for x in group), 2)]
    for i, (dept, vals) in enumerate(by_dept.items(), 1):
        for col, v in enumerate([dept] + vals, 1):
            summary.cell(row=hr + i, column=col, value=v)
    slast = hr + max(len(by_dept), 1)
    _style_body(summary, hr + 1, slast, len(sh), 3)
    _totals_row(summary, slast + 1, hr + 1, slast, 1, range(2, 6))
    _signature(summary, slast + 5, len(sh), settings)
    _widths(summary, [40, 16, 18, 16, 18])
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
    headers = ["মাস", "মূল বেতন", "বাড়িভাড়া", "চিকিৎসা", "শিক্ষা", "টিফিন", "যাতায়াত", "মোবাইল/ধোলাই/মহার্ঘ/PF",
               "উৎসব/নববর্ষ", "অন্যান্য ভাতা/বকেয়া", "মোট বেতন", "মোট কর্তন", "নিট প্রাপ্তি"]
    _title(ws, settings, f"বেতন ইতিহাস — {emp['name_bn']} ({emp['code']})", len(headers))
    _header(ws, 5, headers)
    for i, s in enumerate(slips, 1):
        oa = round(sum(float(x["amount"]) for x in json.loads(s["other_allowance"])), 2)
        others = s["mobile"] + s["wash"] + s["special"] + s["pf_office"]
        vals = [f"{MONTHS_BN[s['month']]} {s['year']}", s["basic"], s["house_rent"], s["medical"], s["education"],
                s["tiffin"], s["transport"], others, s["festival"] + s["boishakhi"], oa, s["gross"], s["total_deduction"],
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
    ("code", "আইডি*"), ("name_bn", "নাম (বাংলা)*"), ("name_en", "নাম (ইংরেজি)"), ("designation", "পদবি"),
    ("department", "শাখা/বিভাগ"), ("grade", "গ্রেড"), ("basic", "মূল বেতন (৩০-০৬-২০২৬)"),
    ("inc_2026", "ইনক্রিমেন্ট ০১-০৭-২০২৬ (হ্যাঁ/না)"), ("inc_2027", "ইনক্রিমেন্ট ০১-০৭-২০২৭ (হ্যাঁ/না)"),
    ("new_basic_override", "নতুন মূল বেতন ০১-০৭-২০২৬ (ঐচ্ছিক)"), ("pf", "PF প্রযোজ্য (হ্যাঁ/না)"),
    ("children", "সন্তান সংখ্যা (০/১/২)"), ("tiffin", "টিফিন ভাতা (হ্যাঁ/না)"), ("transport", "যাতায়াত ভাতা (হ্যাঁ/না)"),
    ("wash", "ধোলাই ভাতা (হ্যাঁ/না)"), ("govt_quarter", "সরকারি বাসা (হ্যাঁ/না)"),
    ("special_rate", "মহার্ঘ হার (auto/0.10/0.15/0)"),
    ("join_date", "যোগদানের তারিখ (YYYY-MM-DD)"), ("dob", "জন্ম তারিখ (YYYY-MM-DD)"),
    ("prl_date", "পিআরএল তারিখ (YYYY-MM-DD)"), ("employment_type", "চাকরির ধরন"), ("status", "অবস্থা (active/inactive)"),
    ("father_name", "পিতার নাম"), ("mother_name", "মাতার নাম"), ("gender", "লিঙ্গ"), ("nid", "এনআইডি"),
    ("mobile", "মোবাইল"), ("email", "ইমেইল"), ("address", "ঠিকানা"), ("tin", "টিআইএন"),
    ("bank_name", "ব্যাংক"), ("bank_branch", "শাখা (ব্যাংক)"), ("account_no", "হিসাব নং"),
]
FLAG_COLUMNS = {"inc_2026", "inc_2027", "pf", "tiffin", "transport", "wash", "govt_quarter"}


def employees_workbook(employees, settings, template=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "কর্মচারী"
    _header(ws, 1, [h for _, h in EMPLOYEE_COLUMNS])
    if template:
        sample = {"code": "NCC-0001", "name_bn": "মোঃ রহিম উদ্দিন", "name_en": "Md. Rahim Uddin", "gender": "পুরুষ",
                  "dob": "1980-05-12", "mobile": "01700000000", "designation": "অফিস সহকারী",
                  "department": "সাধারণ শাখা", "employment_type": "স্থায়ী", "join_date": "2008-07-01", "grade": 14,
                  "basic": 22380, "inc_2026": 1, "inc_2027": 1, "pf": 0, "children": 2, "tiffin": 1, "transport": 1,
                  "wash": 0, "govt_quarter": 0, "special_rate": "auto", "status": "active",
                  "bank_name": "সোনালী ব্যাংক", "bank_branch": "নারায়ণগঞ্জ কর্পোরেট", "account_no": "0000000000000"}
        employees = [sample]
    for r, emp in enumerate(employees, 2):
        for col, (field, _) in enumerate(EMPLOYEE_COLUMNS, 1):
            v = emp[field] if field in emp.keys() else None
            if field in FLAG_COLUMNS and v is not None:
                v = "হ্যাঁ" if v else "না"
            ws.cell(row=r, column=col, value=v).font = Font(name=FONT)
    for i in range(1, len(EMPLOYEE_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.freeze_panes = "C2"
    return _save(wb)
