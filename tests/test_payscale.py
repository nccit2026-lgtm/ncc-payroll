"""বেতন হিসাব ক্যালকুলেটরের সাথে হুবহু মিলছে কিনা:  python tests/test_payscale.py"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from payscale import arrears, basic_for, comparison, fix, month_components, plan  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
KEYMAP = {"hr": "house_rent", "med": "medical", "tif": "tiffin", "trn": "transport", "edu": "education",
          "mob": "mobile", "wsh": "wash"}


def emp_for(c, **kw):
    e = {"grade": c["g"], "basic": c["b"], "join_date": c["join"] or None, "dob": c["dob"] or None,
         "children": c["kids"], "pf": c["pf"], "wash": c["wash"], "tiffin": c["tiffin"], "transport": c["trans"],
         "special_rate": c["rate"], "inc_2026": 1, "inc_2027": 1, "new_basic_override": None, "govt_quarter": 0,
         **{f"ovr_{k}": None for k in ["house_rent", "medical", "education", "tiffin", "transport"]}}
    e.update(kw)
    return e


def test_matches_calculator():
    data = json.load(open(os.path.join(HERE, "calculator_cases.json"), encoding="utf-8"))
    for c, exp in zip(data["cases"], data["expected"]):
        emp = emp_for(c)
        for mine, theirs in zip(comparison(emp), exp["C"]):
            assert round(mine["c"]["basic"], 2) == round(theirs["b"], 2), c
            assert round(mine["gross"], 2) == round(theirs["gross"], 2), c
            assert round(mine["net"], 2) == round(theirs["net"], 2), c
            assert round(mine["c"]["special"], 2) == round(theirs["sp"], 2), c
            for k, v in theirs["A"].items():
                assert round(mine["c"][KEYMAP[k]], 2) == round(v, 2), (c, k)
        a = arrears(emp)
        assert round(a["total"], 2) == round(exp["totalN"], 2), c
        assert round(a["net_total"], 2) == round(exp["netTotal"], 2), c


def test_iBAS_sample_and_phases():
    # ক্যালকুলেটরের নমুনা: গ্রেড ১৩, মূল বেতন ১৯,৮১০
    r = fix(13, 19810)
    assert (r["pay11"], r["pay13"], r["pay16"]) == (28505, 32002.5, 37300)
    emp = emp_for({"g": 13, "b": 19810, "join": "", "dob": "", "kids": 0, "pf": 0, "wash": 0, "tiffin": 1,
                   "trans": 1, "rate": "auto"})
    pl = plan(emp)
    months = {(2026, 6): 19810, (2026, 10): 28505, (2027, 3): 32002.5, (2027, 9): 37300, (2028, 3): 37300,
              (2028, 7): 39100, (2029, 7): 41100}
    for (y, m), b in months.items():
        assert basic_for(emp, pl, y, m) == b, (y, m)
    # ২০২৮ থেকে নতুন হারে ভাতা
    c, _ = month_components(emp, 2028, 1)
    assert (c["house_rent"], c["medical"], c["tiffin"], c["transport"], c["mobile"]) == (14920, 3000, 500, 600, 150)
    c, _ = month_components(emp, 2027, 12)
    assert (c["house_rent"], c["medical"], c["tiffin"], c["transport"], c["mobile"]) == (8914.5, 1500, 200, 300, 0)


def test_override_and_increment_flag():
    emp = emp_for({"g": 13, "b": 19810, "join": "", "dob": "", "kids": 0, "pf": 0, "wash": 0, "tiffin": 1,
                   "trans": 1, "rate": "auto"}, new_basic_override=30000)
    pl = plan(emp)
    assert (pl["P1"], pl["P2"], pl["P3"]) == (30000, 33497.5, 38795)
    no_inc = plan(emp_for({"g": 14, "b": 10200, "join": "", "dob": "", "kids": 0, "pf": 0, "wash": 0, "tiffin": 1,
                           "trans": 1, "rate": "auto"}, inc_2026=0))
    assert no_inc["res"]["inc26"] is False and no_inc["res"]["step26"] == no_inc["res"]["fixed"]


if __name__ == "__main__":
    test_matches_calculator()
    test_iBAS_sample_and_phases()
    test_override_and_increment_flag()
    print("বেতন হিসাবের সব টেস্ট সফল")
