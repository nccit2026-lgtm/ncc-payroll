PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- জাতীয় বেতন স্কেল: প্রতিটি গ্রেডের প্রারম্ভিক ও সর্বোচ্চ মূল বেতন
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
    grade            INTEGER REFERENCES pay_grades(grade),
    basic            INTEGER NOT NULL DEFAULT 0,
    children         INTEGER NOT NULL DEFAULT 0,    -- শিক্ষা সহায়ক ভাতার জন্য
    govt_quarter     INTEGER NOT NULL DEFAULT 0,    -- সরকারি বাসায় থাকলে বাড়ি ভাড়া ভাতা নেই
    gpf_pct          REAL    NOT NULL DEFAULT 5,
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
    created_at       TEXT DEFAULT CURRENT_TIMESTAMP
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
    status            TEXT NOT NULL DEFAULT 'draft',   -- draft / final
    include_festival  INTEGER NOT NULL DEFAULT 0,
    include_boishakhi INTEGER NOT NULL DEFAULT 0,
    note              TEXT,
    created_at        TEXT DEFAULT CURRENT_TIMESTAMP,
    finalized_at      TEXT,
    UNIQUE (year, month)
);

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
