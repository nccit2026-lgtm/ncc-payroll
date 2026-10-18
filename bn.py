"""বাংলা সংখ্যা, মাসের নাম ও টাকার ফরম্যাট।"""

BN_DIGITS = str.maketrans("0123456789", "০১২৩৪৫৬৭৮৯")

MONTHS_BN = [
    "", "জানুয়ারি", "ফেব্রুয়ারি", "মার্চ", "এপ্রিল", "মে", "জুন",
    "জুলাই", "আগস্ট", "সেপ্টেম্বর", "অক্টোবর", "নভেম্বর", "ডিসেম্বর",
]


def bn_num(value):
    if value is None or value == "":
        return ""
    return str(value).translate(BN_DIGITS)


def bn_taka(value):
    """দক্ষিণ এশীয় পদ্ধতিতে কমা (১২,৩৪,৫৬৭)।"""
    if value is None or value == "":
        return ""
    n = int(value)
    sign = "-" if n < 0 else ""
    s = str(abs(n))
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        s = ",".join(groups) + "," + tail
    return (sign + s).translate(BN_DIGITS)


def month_label(year, month):
    return f"{MONTHS_BN[month]} {bn_num(year)}"


def bn_date(value):
    """YYYY-MM-DD → দিন/মাস/বছর (বাংলা সংখ্যায়)।"""
    if not value:
        return ""
    y, m, d = str(value)[:10].split("-")
    return bn_num(f"{d}/{m}/{y}")


def bn_pct(value):
    v = float(value or 0)
    return bn_num(int(v) if v.is_integer() else v)
