"""ডেমো ডাটা: ২০০ জন নমুনা কর্মচারী ও কয়েক মাসের চূড়ান্ত বেতন শিট।

ব্যবহার:  python seed_demo.py            (instance/payroll.db তে)
শুধু পরীক্ষার জন্য — আসল ডাটাবেসে চালাবেন না।
"""
import random
from datetime import date

from app import create_app, save_employee
from db import get_db
from payscale import OLD
from payroll import generate_run

FIRST = ["মোঃ রহিম", "মোঃ করিম", "আব্দুল হালিম", "মোছাঃ রোকেয়া", "ফাতেমা", "নাসরিন", "মোঃ জসিম", "শাহানা", "মোঃ আলমগীর",
         "রফিকুল", "মোঃ সাইফুল", "তাহমিনা", "মোঃ কামরুল", "শারমিন", "মোঃ মিজানুর", "রুবিনা", "মোঃ শফিকুল", "আয়েশা",
         "মোঃ হাবিবুর", "সুমাইয়া", "মোঃ নজরুল", "দিলরুবা", "অনিল", "সুমন", "প্রদীপ", "মোঃ আনোয়ার", "মাহমুদা", "মোঃ ফারুক"]
LAST = ["ইসলাম", "হোসেন", "রহমান", "আহমেদ", "খাতুন", "বেগম", "উদ্দিন", "মিয়া", "সরকার", "চৌধুরী", "দাস", "সাহা", "আক্তার"]
DEPTS = {
    "প্রশাসন": [(9, "সহকারী সচিব"), (13, "প্রশাসনিক কর্মকর্তা"), (14, "উচ্চমান সহকারী"), (16, "অফিস সহকারী"), (20, "অফিস সহায়ক")],
    "হিসাব": [(9, "হিসাবরক্ষণ কর্মকর্তা"), (14, "হিসাব সহকারী"), (16, "অফিস সহকারী কাম কম্পিউটার অপারেটর"), (20, "অফিস সহায়ক")],
    "প্রকৌশল": [(6, "নির্বাহী প্রকৌশলী"), (9, "সহকারী প্রকৌশলী"), (10, "উপ-সহকারী প্রকৌশলী"), (16, "কার্যসহকারী"), (18, "ইলেকট্রিশিয়ান")],
    "স্বাস্থ্য": [(6, "প্রধান স্বাস্থ্য কর্মকর্তা"), (9, "মেডিকেল অফিসার"), (14, "স্যানিটারি ইন্সপেক্টর"), (17, "টিকাদানকারী"), (20, "পরিচ্ছন্নতা পরিদর্শক")],
    "রাজস্ব": [(9, "কর নির্ধারক"), (13, "লাইসেন্স পরিদর্শক"), (15, "কর আদায়কারী"), (16, "অফিস সহকারী"), (20, "অফিস সহায়ক")],
    "বর্জ্য ব্যবস্থাপনা": [(9, "বর্জ্য ব্যবস্থাপনা কর্মকর্তা"), (16, "সুপারভাইজার"), (19, "গাড়িচালক"), (20, "পরিচ্ছন্নতা কর্মী")],
    "পরিবহন": [(10, "পরিবহন পরিদর্শক"), (16, "মেকানিক"), (19, "গাড়িচালক"), (20, "হেলপার")],
    "শিক্ষা ও সংস্কৃতি": [(9, "শিক্ষা কর্মকর্তা"), (13, "লাইব্রেরিয়ান"), (16, "অফিস সহকারী"), (20, "অফিস সহায়ক")],
}
BANKS = ["সোনালী ব্যাংক", "জনতা ব্যাংক", "অগ্রণী ব্যাংক", "রূপালী ব্যাংক"]


def main():
    random.seed(2026)
    app = create_app()
    with app.app_context():
        conn = get_db()
        if conn.execute("SELECT COUNT(*) FROM employees").fetchone()[0]:
            print("ডাটাবেসে আগে থেকেই কর্মচারী আছে — ডেমো ডাটা যোগ করা হলো না।")
            return
        for i in range(1, 201):
            dept = random.choice(list(DEPTS))
            grade, desig = random.choice(DEPTS[dept])
            basic = random.choice(OLD[grade])
            dob = date(random.randint(1968, 1998), random.randint(1, 12), random.randint(1, 28))
            name = f"{random.choice(FIRST)} {random.choice(LAST)}"
            emp_id = save_employee({
                "code": f"NCC-{i:04d}", "name_bn": name, "name_en": None, "father_name": None, "mother_name": None,
                "gender": "মহিলা" if any(w in name for w in ["মোছাঃ", "ফাতেমা", "নাসরিন", "শাহানা", "তাহমিনা", "শারমিন", "রুবিনা", "আয়েশা", "সুমাইয়া", "দিলরুবা", "মাহমুদা", "রোকেয়া"]) else "পুরুষ",
                "dob": dob.isoformat(), "nid": str(random.randint(10**9, 10**10 - 1)),
                "mobile": f"01{random.choice('3456789')}{random.randint(10**7, 10**8 - 1)}", "email": None,
                "address": "নারায়ণগঞ্জ", "designation": desig, "department": dept,
                "employment_type": random.choice(["স্থায়ী"] * 9 + ["চুক্তিভিত্তিক"]),
                "join_date": date(max(dob.year + 22, 1995), random.randint(1, 12), 1).isoformat(), "prl_date": None,
                "status": "active", "grade": grade, "basic": basic, "children": random.choice([0, 1, 1, 2, 2]),
                "govt_quarter": 1 if random.random() < 0.05 else 0, "gpf_no": None, "tin": None,
                "inc_2026": 1, "inc_2027": 1, "new_basic_override": None, "pf": 1 if random.random() < 0.3 else 0,
                "special_rate": "auto", "tiffin": 1, "transport": 1, "wash": 1 if grade == 20 else 0,
                "bank_name": random.choice(BANKS), "bank_branch": "নারায়ণগঞ্জ",
                "account_no": str(random.randint(10**12, 10**13 - 1)),
                "ovr_house_rent": None, "ovr_medical": None, "ovr_education": None, "ovr_tiffin": None, "ovr_transport": None,
            })
            if random.random() < 0.15:
                conn.execute("INSERT INTO employee_items (employee_id, kind, name, amount) VALUES (?, 'deduction', ?, ?)",
                             (emp_id, "গৃহ নির্মাণ ঋণ কিস্তি", random.choice([1500, 2000, 3000])))
            if grade <= 9 and random.random() < 0.3:
                conn.execute("INSERT INTO employee_items (employee_id, kind, name, amount) VALUES (?, 'allowance', ?, ?)",
                             (emp_id, "মোবাইল ভাতা", 1500))
        conn.commit()

        # জুলাই–সেপ্টেম্বর ২০২৬ চূড়ান্ত শিট (সেপ্টেম্বরে উৎসব ভাতা সহ)
        for month, fest in [(7, False), (8, False), (9, True)]:
            run_id, _ = generate_run(2026, month, fest, False, "ঈদ/পূজা উৎসব ভাতা সহ" if fest else "")
            conn.execute("UPDATE payroll_runs SET status = 'final', finalized_at = CURRENT_TIMESTAMP WHERE id = ?", (run_id,))
        conn.commit()
        print("ডেমো ডাটা তৈরি হয়েছে: ২০০ জন কর্মচারী, জুলাই–সেপ্টেম্বর ২০২৬ এর বেতন শিট।")


if __name__ == "__main__":
    main()
