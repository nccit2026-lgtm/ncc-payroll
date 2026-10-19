"""জাতীয় বেতনস্কেল ২০১৫ → ২০২৬: পে ফিক্সেশন, কিস্তি, ভাতা ও বকেয়া (নারায়ণগঞ্জ সিটি কর্পোরেশন এলাকা)।

"নাসিক বেতন হিসাব ২০২৬" ক্যালকুলেটরের নিয়ম অনুযায়ী:
- ৩০-০৬-২০২৬ পর্যন্ত: ২০১৫ স্কেল, ২০১৫ আদেশের ভাতা ও বিশেষ সুবিধা (মহার্ঘ ভাতা)
- ০১-০৭-২০২৬: ১ম কিস্তি (১–৯ গ্রেড ৪০%, ১০–২০ গ্রেড ৫০%) + ইনক্রিমেন্ট; মহার্ঘ ভাতা বিলুপ্ত
- ০১-০১-২০২৭: ২য় কিস্তি (৩০% / ২৫%)
- ০১-০৭-২০২৭: ৩য় কিস্তি, পূর্ণ স্কেল + ইনক্রিমেন্ট
- ৩১-১২-২০২৭ পর্যন্ত সব ভাতা ৩০-০৬-২০২৬ এর অঙ্কে; ০১-০১-২০২৮ থেকে ২০২৬ আদেশের হারে
"""
import calendar
from datetime import date

OLD = {1: [78000], 2: [66000, 68480, 71050, 73720, 76490], 3: [56500, 58760, 61120, 63570, 66120, 68770, 71530, 74400], 4: [50000, 52000, 54080, 56250, 58500, 60840, 63280, 65820, 68460, 71200], 5: [43000, 44940, 46970, 49090, 51300, 53610, 56030, 58560, 61200, 63960, 66840, 69850], 6: [35500, 37280, 39150, 41110, 43170, 45330, 47600, 49980, 52480, 55110, 57870, 60770, 63810, 67010], 7: [29000, 30450, 31980, 33580, 35260, 37030, 38890, 40840, 42890, 45040, 47300, 49670, 52160, 54770, 57510, 60390, 63410], 8: [23000, 24150, 25360, 26630, 27970, 29370, 30840, 32390, 34010, 35720, 37510, 39390, 41360, 43430, 45610, 47900, 50300, 52820, 55470], 9: [22000, 23100, 24260, 25480, 26760, 28100, 29510, 30990, 32540, 34170, 35880, 37680, 39570, 41550, 43630, 45820, 48120, 50530, 53060], 10: [16000, 16800, 17640, 18530, 19460, 20440, 21470, 22550, 23680, 24870, 26120, 27430, 28810, 30260, 31780, 33370, 35040, 36800, 38640], 11: [12500, 13130, 13790, 14480, 15210, 15980, 16780, 17620, 18510, 19440, 20420, 21450, 22530, 23660, 24850, 26100, 27410, 28790, 30230], 12: [11300, 11870, 12470, 13100, 13760, 14450, 15180, 15940, 16740, 17580, 18460, 19390, 20360, 21380, 22450, 23580, 24760, 26000, 27300], 13: [11000, 11550, 12130, 12740, 13380, 14050, 14760, 15500, 16280, 17100, 17960, 18860, 19810, 20810, 21860, 22960, 24110, 25320, 26590], 14: [10200, 10710, 11250, 11820, 12420, 13050, 13710, 14400, 15120, 15880, 16680, 17520, 18400, 19320, 20290, 21310, 22380, 23500, 24680], 15: [9700, 10190, 10700, 11240, 11810, 12410, 13040, 13700, 14390, 15110, 15870, 16670, 17510, 18390, 19310, 20280, 21300, 22370, 23490], 16: [9300, 9770, 10260, 10780, 11320, 11890, 12490, 13120, 13780, 14470, 15200, 15960, 16760, 17600, 18480, 19410, 20390, 21410, 22490], 17: [9000, 9450, 9930, 10430, 10960, 11510, 12090, 12700, 13340, 14010, 14720, 15460, 16240, 17060, 17920, 18820, 19770, 20760, 21800], 18: [8800, 9240, 9710, 10200, 10710, 11250, 11820, 12420, 13050, 13710, 14400, 15120, 15880, 16680, 17520, 18400, 19320, 20290, 21310], 19: [8500, 8930, 9380, 9850, 10350, 10870, 11420, 12000, 12600, 13230, 13900, 14600, 15330, 16100, 16910, 17760, 18650, 19590, 20570], 20: [8250, 8670, 9110, 9570, 10050, 10560, 11090, 11650, 12240, 12860, 13510, 14190, 14900, 15650, 16440, 17270, 18140, 19050, 20010]}
NEW = {1: [156000], 2: [132000, 135700, 139400, 143200, 147200, 151200, 153000], 3: [113000, 117000, 121100, 125300, 129700, 134300, 139000, 143800, 148800], 4: [100000, 103500, 107200, 110900, 114800, 118800, 123000, 127300, 131700, 136300, 142400], 5: [86000, 89500, 93100, 96800, 100700, 104700, 108900, 113200, 117700, 122500, 127400, 132400, 139700], 6: [71000, 74600, 78300, 82200, 86400, 90700, 95200, 100000, 104900, 110200, 115700, 121500, 127600, 134000], 7: [58000, 60900, 64000, 67200, 70500, 74100, 77800, 81700, 85700, 90000, 94500, 99200, 104200, 109400, 114900, 120600, 126800], 8: [46000, 48300, 50800, 53300, 56000, 58800, 61700, 64800, 68000, 71400, 75000, 78700, 82700, 86800, 91100, 95700, 100500, 105500, 110800], 9: [44000, 46200, 48600, 51000, 53500, 56200, 59000, 62000, 65100, 68300, 71700, 75300, 79100, 83000, 87200, 91500, 96100, 100900, 105900], 10: [32000, 33600, 35300, 37100, 38900, 40900, 42900, 45100, 47300, 49700, 52200, 54800, 57500, 60400, 63400, 66600, 69900, 73400, 77300], 11: [25000, 26300, 27600, 29000, 30400, 32000, 33600, 35200, 37000, 38800, 40800, 42800, 44900, 47200, 49500, 52000, 54600, 57300, 60500], 12: [24300, 25600, 26800, 28200, 29600, 31100, 32600, 34200, 36000, 37700, 39600, 41600, 43700, 45900, 48200, 50600, 53100, 55700, 58700], 13: [24000, 25200, 26500, 27800, 29200, 30700, 32200, 33800, 35500, 37300, 39100, 41100, 43200, 45300, 47600, 49900, 52400, 55100, 58000], 14: [23500, 24700, 26000, 27300, 28600, 30000, 31500, 33100, 34800, 36500, 38300, 40200, 42300, 44400, 46600, 48900, 51300, 53900, 56800], 15: [22800, 24000, 25200, 26400, 27800, 29100, 30600, 32100, 33700, 35400, 37200, 39000, 41000, 43000, 45200, 47400, 49800, 52300, 55200], 16: [21900, 23000, 24200, 25400, 26700, 28000, 29400, 30900, 32400, 34000, 35700, 37500, 39400, 41300, 43400, 45600, 47900, 50200, 52900], 17: [21400, 22500, 23600, 24800, 26100, 27400, 28700, 30200, 31700, 33200, 34900, 36700, 38500, 40400, 42400, 44500, 46800, 49100, 51900], 18: [21000, 22100, 23200, 24400, 25600, 26900, 28200, 29600, 31100, 32600, 34300, 36000, 37800, 39600, 41600, 43700, 45900, 48200, 50900], 19: [20500, 21600, 22700, 23800, 25000, 26200, 27500, 28900, 30300, 31900, 33400, 35100, 36900, 38700, 40600, 42700, 44800, 47000, 49600], 20: [20000, 21000, 22100, 23200, 24400, 25600, 26900, 28200, 29600, 31100, 32600, 34300, 36000, 37800, 39600, 41600, 43700, 45900, 48400]}

# বাড়িভাড়া, ২০১৫ আদেশ অনুচ্ছেদ ১৭(৭) — নারায়ণগঞ্জ সিটি কর্পোরেশন কলাম: (মূল বেতন পর্যন্ত, হার, ন্যূনতম)
HR15 = [(9700, 0.55, 5000, "টাকা ৯,৭০০ পর্যন্ত"), (16000, 0.50, 5400, "টাকা ৯,৭০১ হইতে ১৬,০০০ পর্যন্ত"),
        (35500, 0.45, 8000, "টাকা ১৬,০০১ হইতে ৩৫,৫০০ পর্যন্ত"), (float("inf"), 0.40, 16000, "টাকা ৩৫,৫০১ ও তদূর্ধ্ব")]
# বাড়িভাড়া, ২০২৬ আদেশ অনুচ্ছেদ ১৫(৬) — ০১-০১-২০২৮ থেকে, গ্রেড অনুযায়ী, ন্যূনতম নেই: (গ্রেড হতে, পর্যন্ত, হার)
HR26 = [(16, 20, 0.50), (10, 15, 0.40), (5, 9, 0.35), (1, 4, 0.30)]

FIX_DATE = date(2026, 6, 30)      # এই তারিখের মূল বেতন থেকে পে ফিক্সেশন
ON26, ON27, ON28 = date(2026, 7, 1), date(2027, 7, 1), date(2028, 1, 1)

# মাস → কোন ধাপ (বেতন শিট ও বিবরণীতে দেখানোর জন্য)
PHASES = [
    ((2026, 7), "২০১৫ স্কেল", "২০১৫ স্কেল ও ২০১৫ আদেশের ভাতা, বিশেষ সুবিধাসহ"),
    ((2027, 1), "১ম কিস্তি", "০১-০৭-২০২৬ থেকে · ভাতা ৩০-০৬-২০২৬ এর অঙ্কে"),
    ((2027, 7), "২য় কিস্তি", "০১-০১-২০২৭ থেকে · ভাতা ৩০-০৬-২০২৬ এর অঙ্কে"),
    ((2028, 1), "পূর্ণ স্কেল", "০১-০৭-২০২৭ থেকে · ভাতা ৩০-০৬-২০২৬ এর অঙ্কে"),
    ((9999, 1), "নতুন হারে ভাতা", "০১-০১-২০২৮ থেকে · ২০২৬ আদেশের ভাতা"),
]


def r2(x):
    return round(x + 0.0, 2)


def phase_of(y, m):
    for limit, name, desc in PHASES:
        if (y, m) < limit:
            return name, desc


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    m += 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def parse_date(s):
    if not s:
        return None
    try:
        y, m, d = (int(p) for p in str(s)[:10].split("-"))
        return date(y, m, d)
    except ValueError:
        return None


def hr15(b):
    upto, rate, mn, label = next(s for s in HR15 if b <= s[0])
    by_rate = r2(b * rate)
    return max(by_rate, mn), rate, mn, label


def hr26(b, g):
    rate = next(s[2] for s in HR26 if s[0] <= g <= s[1])
    return r2(b * rate), rate


def special_rate(emp):
    v = (emp["special_rate"] or "auto") if "special_rate" in emp.keys() else "auto"
    if v == "auto":
        return 0.10 if emp["grade"] <= 9 else 0.15
    return float(v)


def special(b, rate):
    if not rate:
        return 0
    return max(r2(b * rate), 1000)


def is_new_joiner(emp):
    j = parse_date(emp["join_date"])
    return bool(j and j > FIX_DATE)


def increments(emp):
    """ইনক্রিমেন্ট প্রাপ্যতা: কর্মচারীর ঘরের মান এবং যোগদানের তারিখ থেকে ৬ মাস পূর্ণ হওয়া — দুটোই লাগবে।"""
    inc26, inc27 = bool(emp["inc_2026"]), bool(emp["inc_2027"])
    j = parse_date(emp["join_date"])
    if j:
        inc26 = inc26 and add_months(j, 6) <= ON26
        inc27 = inc27 and add_months(j, 6) <= ON27
    return inc26, inc27


def fix(g, basic, inc26=True, inc27=True):
    """২০২৬ আদেশ অনুচ্ছেদ ১(৩), ৫ ও ৯ অনুযায়ী পে ফিক্সেশন (ক্যালকুলেটরের ২ নং অংশ)।"""
    o, n = OLD.get(g), NEW.get(g)
    if not o:
        return {"error": "গ্রেড নির্বাচন করা হয়নি"}
    if not basic or basic <= 0:
        return {"error": "৩০-০৬-২০২৬ তারিখের মূল বেতন দেওয়া হয়নি"}
    if basic < o[0]:
        return {"error": f"মূল বেতন গ্রেড-{g} এর ২০১৫ স্কেলের প্রারম্ভিক ধাপ {o[0]} এর চেয়ে কম"}
    last = len(n) - 1
    diff = basic - o[0]
    total = diff + n[0]
    idx = next((i for i, v in enumerate(n) if v >= total), None)
    over = idx is None
    if over:
        idx = last
    fixed = n[idx]
    i26 = min(idx + 1, last) if inc26 else idx
    step26 = n[i26]
    inc = step26 - fixed
    raise_ = fixed - basic
    p = (0.40, 0.30, 0.30) if g <= 9 else (0.50, 0.25, 0.25)
    part1 = r2(raise_ * p[0])
    pay10 = r2(basic + part1)
    pay11 = r2(pay10 + inc)
    part2 = r2(raise_ * p[1])
    pay13 = r2(pay11 + part2)
    part3 = r2(raise_ * p[2])
    pay15 = r2(pay13 + part3)
    i27 = min(i26 + 1, last) if inc27 else i26
    pay16 = n[i27]
    below = sum(1 for v in o if v <= basic)
    old_paid = basic if (not inc26 or below >= len(o)) else o[below]
    return dict(g=g, basic=basic, old_init=o[0], old_max=o[-1], new_init=n[0], new_max=n[last], diff=diff,
                total=total, idx=idx, fixed=fixed, i26=i26, step26=step26, inc=inc, raise_=raise_, p=p,
                part1=part1, pay10=pay10, pay11=pay11, part2=part2, pay13=pay13, part3=part3, pay15=pay15,
                i27=i27, pay16=pay16, over=over, off_step=basic not in o, old_step=o.index(basic) if basic in o else -1,
                below=below, old_paid=old_paid, inc26=inc26, inc27=inc27)


def plan(emp):
    """কর্মচারীর তিন কিস্তির মূল বেতন (P1, P2, P3) — হাতে দেওয়া নতুন মূল বেতন থাকলে পার্থক্য পরের কিস্তিতেও ধরা হয়।"""
    if is_new_joiner(emp):
        return {"new_joiner": True, "basic": emp["basic"], "notes": [
            "৩০-০৬-২০২৬ এর পরে যোগদান: পে ফিক্সেশন প্রযোজ্য নয় (২০২৬ আদেশ, অনুচ্ছেদ ১০)। "
            "মূল বেতন ঘরে ২০২৬ স্কেলের মূল বেতন দিন এবং ভাতা যাচাই করুন।"]}
    inc26, inc27 = increments(emp)
    res = fix(emp["grade"], emp["basic"], inc26, inc27)
    if "error" in res:
        return {"error": res["error"]}
    override = emp["new_basic_override"]
    p1 = r2(override) if override else res["pay11"]
    adj = r2(p1 - res["pay11"])
    p3 = r2(res["pay16"] + adj) if override else res["pay16"]
    notes = []
    if res["over"]:
        notes.append("মোট অঙ্ক ২০২৬ স্কেলের সর্বোচ্চ ধাপ ছাড়িয়েছে; সর্বোচ্চ ধাপে নির্ধারণ করা হয়েছে।")
    if res["off_step"]:
        notes.append("৩০-০৬-২০২৬ এর মূল বেতন ২০১৫ স্কেলের এই গ্রেডের কোনো ধাপের সাথে মেলেনি; সংখ্যাটি যাচাই করুন।")
    if override:
        notes.append(f"নতুন মূল বেতন হাতে দেওয়া হয়েছে ({p1}); পে ফিক্সেশনের হিসাব ছিল {res['pay11']}।")
    if emp["join_date"] and not inc26:
        notes.append("০১-০৭-২০২৬ এ ইনক্রিমেন্ট ধরা হয়নি (চাকরিকাল ৬ মাস পূর্ণ হয়নি বা 'না' দেওয়া)।")
    return {"res": res, "P1": p1, "P2": r2(res["pay13"] + adj), "P3": p3, "override": bool(override), "notes": notes}


def basic_for(emp, pl, y, m):
    """নির্দিষ্ট মাসের মূল বেতন।"""
    if pl.get("new_joiner"):
        return emp["basic"]
    res = pl["res"]
    if (y, m) < (2026, 7):
        return res["basic"]
    if (y, m) < (2027, 1):
        return pl["P1"]
    if (y, m) < (2027, 7):
        return pl["P2"]
    # ০১-০৭-২০২৮ থেকে প্রতি ১ জুলাই পরবর্তী ধাপ (সর্বোচ্চ ধাপ পর্যন্ত)
    extra = max(0, y - 2028 + (1 if m >= 7 else 0))
    if not extra:
        return pl["P3"]
    steps = NEW[res["g"]]
    j = min(res["i27"] + extra, len(steps) - 1)
    return r2(pl["P3"] + steps[j] - steps[res["i27"]])


def allowance_basic(emp, pl, y, m):
    """ভাতার ভিত্তি মূল বেতন: ৩১-১২-২০২৭ পর্যন্ত ৩০-০৬-২০২৬ এর মূল বেতন, ০১-০১-২০২৮ থেকে সেই মাসের মূল বেতন।
    (উৎসব ভাতা, নববর্ষ ভাতা ও % হিসাবের অন্যান্য ভাতায় ব্যবহৃত)"""
    if (y, m) < (2028, 1) and not pl.get("new_joiner"):
        return pl["res"]["basic"]
    return basic_for(emp, pl, y, m)


def over_50(dob, on):
    d = parse_date(dob)
    if not d:
        return False
    try:
        fiftieth = date(d.year + 50, d.month, d.day)
    except ValueError:
        fiftieth = date(d.year + 50, 3, 1)
    return on > fiftieth


def allowances_2015(b0, g, emp):
    """২০১৫ আদেশের ভাতা (৩০-০৬-২০২৬ এর অঙ্ক; ৩১-১২-২০২৭ পর্যন্ত এই অঙ্কেই)।"""
    kids = min(int(emp["children"] or 0), 2)
    return {
        "house_rent": 0 if emp["govt_quarter"] else hr15(b0)[0],
        "medical": 1500,
        "tiffin": 200 if g >= 11 and emp["tiffin"] else 0,
        "transport": 300 if g >= 11 and emp["transport"] else 0,
        "education": 500 * kids,
        "mobile": 0,
        "wash": 100 if emp["wash"] else 0,
    }


def allowances_2026(b, g, emp, on):
    """২০২৬ আদেশের ভাতা (০১-০১-২০২৮ থেকে)।"""
    kids = min(int(emp["children"] or 0), 2)
    return {
        "house_rent": 0 if emp["govt_quarter"] else hr26(b, g)[0],
        "medical": 4000 if over_50(emp["dob"], on) else 3000,
        "tiffin": 500 if g >= 11 and emp["tiffin"] else 0,
        "transport": 600 if g >= 11 and emp["transport"] else 0,
        "education": 500 * kids,
        "mobile": 500 if g <= 5 else 150,
        "wash": 300 if emp["wash"] else 0,
    }


OVERRIDABLE = ["house_rent", "medical", "education", "tiffin", "transport", "mobile"]


def month_components(emp, y, m):
    """এক মাসের মূল বেতন, ভাতা, বিশেষ সুবিধা ও PF (উৎসব/বকেয়া/অতিরিক্ত আইটেম ছাড়া)।"""
    pl = plan(emp)
    if "error" in pl:
        return None, pl
    g = emp["grade"]
    b = basic_for(emp, pl, y, m)
    first = date(y, m, 1)
    if pl.get("new_joiner"):
        a = allowances_2026(b, g, emp, first) if (y, m) >= (2028, 1) else allowances_2015(b, g, emp)
        sp = 0
    else:
        b0 = pl["res"]["basic"]
        a = allowances_2015(b0, g, emp) if (y, m) < (2028, 1) else allowances_2026(b, g, emp, first)
        sp = special(b0, special_rate(emp)) if (y, m) < (2026, 7) else 0
    for k in OVERRIDABLE:
        ovr = emp["ovr_" + k] if ("ovr_" + k) in emp.keys() else None
        if ovr is not None:
            a[k] = ovr
    pf = r2(b * 0.10) if emp["pf"] else 0
    return {"basic": b, **a, "special": sp, "pf_office": pf, "pf_fund": pf, "pf_own": pf}, pl


def arrears(emp):
    """জুলাই–সেপ্টেম্বর ২০২৬ এর বকেয়া (২০১৫ স্কেলে দেওয়া হয়েছিল) — ক্যালকুলেটরের ৩ নং অংশ, স্বয়ংক্রিয় অঙ্কে।"""
    pl = plan(emp)
    if "error" in pl or pl.get("new_joiner"):
        return None
    res, p1 = pl["res"], pl["P1"]
    rate = special_rate(emp)
    paid_b = res["old_paid"]
    hr_due = 0 if emp["govt_quarter"] else hr15(res["basic"])[0]
    hr_paid = 0 if emp["govt_quarter"] else hr15(paid_b)[0]
    sp = special(paid_b, rate)
    d_b = r2(p1 - paid_b)
    hr_adj = r2(hr_due - hr_paid)
    monthly = r2(d_b + hr_adj - sp)
    pf_a = r2(r2(p1 * 0.10) - r2(paid_b * 0.10)) if emp["pf"] else 0
    return {"paid_basic": paid_b, "paid_hr": hr_paid, "paid_special": sp, "due_basic": p1, "due_hr": hr_due,
            "diff_basic": d_b, "hr_adj": hr_adj, "monthly": monthly, "pf_a": pf_a, "months": 3,
            "total": r2(monthly * 3), "pf_total": r2(pf_a * 3), "net_total": r2((monthly - pf_a) * 3)}


def comparison(emp):
    """পাঁচটি তারিখে মূল ও মোট বেতনের তুলনা (ক্যালকুলেটরের ১ নং অংশ)।"""
    cols = []
    for (y, m), label in [((2026, 6), "৩০-০৬-২০২৬"), ((2026, 7), "০১-০৭-২০২৬"), ((2027, 1), "০১-০১-২০২৭"),
                          ((2027, 7), "০১-০৭-২০২৭"), ((2028, 1), "০১-০১-২০২৮")]:
        c, _ = month_components(emp, y, m)
        if c is None:
            return []
        gross = r2(sum(c[k] for k in ["basic", "house_rent", "medical", "tiffin", "transport", "education", "mobile",
                                      "wash", "special", "pf_office"]))
        cols.append({"label": label, "phase": phase_of(y, m)[0], "c": c, "gross": gross,
                     "net": r2(gross - c["pf_fund"] - c["pf_own"])})
    return cols


def prl_date(dob):
    d = parse_date(dob)
    return add_months(d, 59 * 12) if d else None


def age_on_date(dob, on):
    d = parse_date(dob)
    if not d:
        return None
    return on.year - d.year - ((on.month, on.day) < (d.month, d.day))
