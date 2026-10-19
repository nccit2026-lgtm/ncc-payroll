PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- পুরনো সংস্করণের টেবিল (এখন ব্যবহৃত হয় না; বেতন স্কেল payscale.py তে)
CREATE TABLE IF NOT EXISTS pay_grades (
    grade       INTEGER PRIMARY KEY,
    start_basic INTEGER NOT NULL,
    max_basic   INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS employees (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    code             TEXT NOT NULL UNIQUE,          -- কর্মচারী আইডি
    name_bn          TEXT NOT NULL,
    name_en          TEXT,
    father_name      TEXT,
    mother_name      TEXT,
    gender           TEXT,
    religion         TEXT,                          -- ইসলাম / হিন্দু / বৌদ্ধ / খ্রিস্টান / অন্যান্য
    dob              TEXT,                          -- YYYY-MM-DD
    nid              TEXT,
    mobile           TEXT,
    email            TEXT,
    address          TEXT,
    designation      TEXT,
    department       TEXT,
    employment_type  TEXT DEFAULT 'স্থায়ী',
    join_date        TEXT,
    prl_date         TEXT,
    status           TEXT DEFAULT 'active',         -- active / inactive
    separation_reason TEXT,                         -- নিষ্ক্রিয়ের কারণ: অবসর / পদত্যাগ / বদলি ...
    separation_date  TEXT,
    grade            INTEGER,
    basic            REAL    NOT NULL DEFAULT 0,    -- ৩০-০৬-২০২৬ এর মূল বেতন (২০১৫ স্কেল)
    children         INTEGER NOT NULL DEFAULT 0,    -- শিক্ষা সহায়ক ভাতার জন্য
    govt_quarter     INTEGER NOT NULL DEFAULT 0,    -- সরকারি বাসায় থাকলে বাড়ি ভাড়া ভাতা নেই
    gpf_pct          REAL    NOT NULL DEFAULT 0,    -- ব্যবহৃত হয় না (PF দেখুন)
    inc_2026         INTEGER NOT NULL DEFAULT 1,    -- ০১-০৭-২০২৬ এ ইনক্রিমেন্ট প্রাপ্য
    inc_2027         INTEGER NOT NULL DEFAULT 1,    -- ০১-০৭-২০২৭ এ ইনক্রিমেন্ট প্রাপ্য
    new_basic_override REAL,                        -- আইবাস++/দপ্তরে ভিন্ন হলে ০১-০৭-২০২৬ এর মূল বেতন
    pf               INTEGER NOT NULL DEFAULT 0,    -- ভবিষ্য তহবিল প্রযোজ্য
    special_rate     TEXT    NOT NULL DEFAULT 'auto', -- মহার্ঘ ভাতা: auto / 0.10 / 0.15 / 0
    tiffin           INTEGER NOT NULL DEFAULT 1,
    transport        INTEGER NOT NULL DEFAULT 1,
    wash             INTEGER NOT NULL DEFAULT 0,    -- ধোলাই ভাতা
    gpf_no           TEXT,
    tin              TEXT,
    bank_name        TEXT,
    bank_branch      TEXT,
    account_no       TEXT,
    -- খালি (NULL) থাকলে নিয়ম অনুযায়ী স্বয়ংক্রিয় হিসাব হবে, মান দিলে সেটাই বসবে
    ovr_house_rent   INTEGER,
    ovr_medical      INTEGER,
    ovr_education    INTEGER,
    ovr_tiffin       INTEGER,
    ovr_transport    INTEGER,
    ovr_mobile       REAL,
    created_at       TEXT DEFAULT CURRENT_TIMESTAMP
);

-- শাখা/বিভাগের তালিকা (কর্মচারী না থাকলেও শাখা আগে থেকে রাখা যায়)
CREATE TABLE IF NOT EXISTS departments (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

-- কর্মচারীভিত্তিক অতিরিক্ত ভাতা / কর্তন (প্রতি মাসে বসবে)
CREATE TABLE IF NOT EXISTS employee_items (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    kind        TEXT NOT NULL CHECK (kind IN ('allowance', 'deduction')),
    name        TEXT NOT NULL,
    amount      INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS users (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role          TEXT NOT NULL CHECK (role IN ('admin', 'employee')),
    employee_id   INTEGER REFERENCES employees(id) ON DELETE CASCADE,
    must_change   INTEGER NOT NULL DEFAULT 0
);

-- প্রতি মাসের বেতন শিট
CREATE TABLE IF NOT EXISTS payroll_runs (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    year              INTEGER NOT NULL,
    month             INTEGER NOT NULL,
    kind              TEXT NOT NULL DEFAULT 'salary',  -- salary / festival / boishakhi / arrears / other
    title             TEXT,                            -- যেমন: ঈদুল ফিতর উৎসব ভাতা
    religion          TEXT,                            -- উৎসব ভাতা: যে ধর্মের কর্মচারীরা পাবেন
    multiplier        REAL,                            -- উৎসব ভাতা: মূল বেতনের কত গুণ
    pct               REAL,                            -- অন্যান্য ভাতা: মূল বেতনের %
    amount            REAL,                            -- অন্যান্য ভাতা: জনপ্রতি নির্দিষ্ট টাকা
    status            TEXT NOT NULL DEFAULT 'draft',   -- draft / final
    include_festival  INTEGER NOT NULL DEFAULT 0,
    include_boishakhi INTEGER NOT NULL DEFAULT 0,
    include_arrears   INTEGER NOT NULL DEFAULT 0,
    manual            INTEGER NOT NULL DEFAULT 0,      -- ১ = সফটওয়্যারের আগের মাস, হাতে এন্ট্রি (তালিকা/রিপোর্টে আসে না)
    note              TEXT,
    created_at        TEXT DEFAULT CURRENT_TIMESTAMP,
    finalized_at      TEXT
);
CREATE INDEX IF NOT EXISTS idx_runs_month ON payroll_runs(year, month);

-- প্রত্যেক কর্মচারীর মাসিক বেতন বিবরণী (তৈরির সময়ের তথ্যের স্থায়ী কপি)
CREATE TABLE IF NOT EXISTS payslips (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id          INTEGER NOT NULL REFERENCES payroll_runs(id) ON DELETE CASCADE,
    employee_id     INTEGER NOT NULL REFERENCES employees(id),
    code            TEXT,
    name_bn         TEXT,
    designation     TEXT,
    department      TEXT,
    grade           INTEGER,
    account_no      TEXT,
    basic           INTEGER NOT NULL DEFAULT 0,
    house_rent      INTEGER NOT NULL DEFAULT 0,
    medical         INTEGER NOT NULL DEFAULT 0,
    education       INTEGER NOT NULL DEFAULT 0,
    tiffin          INTEGER NOT NULL DEFAULT 0,
    transport       INTEGER NOT NULL DEFAULT 0,
    festival        INTEGER NOT NULL DEFAULT 0,
    boishakhi       INTEGER NOT NULL DEFAULT 0,
    mobile          REAL NOT NULL DEFAULT 0,
    wash            REAL NOT NULL DEFAULT 0,
    special         REAL NOT NULL DEFAULT 0,
    pf_office       REAL NOT NULL DEFAULT 0,
    pf_fund         REAL NOT NULL DEFAULT 0,
    pf_own          REAL NOT NULL DEFAULT 0,
    phase           TEXT,
    ref_basic       REAL,                           -- বোনাস/ভাতার শিটে হিসাবের ভিত্তি মূল বেতন
    other_allowance TEXT NOT NULL DEFAULT '[]',     -- JSON [{name, amount}]
    gpf             INTEGER NOT NULL DEFAULT 0,
    revenue_stamp   INTEGER NOT NULL DEFAULT 0,
    other_deduction TEXT NOT NULL DEFAULT '[]',     -- JSON [{name, amount}]
    gross           INTEGER NOT NULL DEFAULT 0,
    total_deduction INTEGER NOT NULL DEFAULT 0,
    net             INTEGER NOT NULL DEFAULT 0,
    UNIQUE (run_id, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_payslips_employee ON payslips(employee_id);

-- কর্মচারী নিজে প্রোফাইল বদলালে তার রেকর্ড (অ্যাডমিন দেখতে পাবেন)
CREATE TABLE IF NOT EXISTS employee_changes (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    field       TEXT NOT NULL,
    old_value   TEXT,
    new_value   TEXT,
    changed_at  TEXT DEFAULT CURRENT_TIMESTAMP
);

-- ঋণ (গৃহ নির্মাণ ঋণ, ভবিষ্য তহবিল অগ্রিম, অন্যান্য ঋণ): প্রতি মাসের বেতন শিট থেকে কিস্তি কাটা হয়
CREATE TABLE IF NOT EXISTS loans (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id        INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    kind               TEXT NOT NULL,
    installment        REAL NOT NULL,            -- প্রতি কিস্তির টাকা
    total_installments INTEGER,                  -- মোট কিস্তি (জানা না থাকলে খালি)
    paid_before        INTEGER NOT NULL DEFAULT 0, -- এই সফটওয়্যার শুরুর আগে পরিশোধিত কিস্তি
    principal          REAL,                     -- মোট ঋণ (ঐচ্ছিক)
    start_ym           TEXT,                     -- যে মাস থেকে কাটা শুরু (YYYY-MM)
    status             TEXT NOT NULL DEFAULT 'active',  -- active / closed
    note               TEXT,
    created_at         TEXT DEFAULT CURRENT_TIMESTAMP
);

-- কোন শিটে কোন ঋণের কিস্তি কাটা হয়েছে (শিট চূড়ান্ত হলে কিস্তি পরিশোধিত ধরা হয়)
CREATE TABLE IF NOT EXISTS loan_payments (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    loan_id INTEGER NOT NULL REFERENCES loans(id) ON DELETE CASCADE,
    run_id  INTEGER NOT NULL REFERENCES payroll_runs(id) ON DELETE CASCADE,
    amount  REAL NOT NULL,
    UNIQUE (loan_id, run_id)
);
