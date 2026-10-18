"""অ্যাডমিন পাসওয়ার্ড রিসেট (সার্ভারের কনসোল থেকে):

    .venv/bin/python reset_admin.py নতুন-পাসওয়ার্ড

পাসওয়ার্ড বদলায়, ভুল-চেষ্টার লক খুলে দেয় এবং 'admin' অ্যাকাউন্ট না থাকলে তৈরি করে।
প্রথম লগইনে আবার পাসওয়ার্ড পরিবর্তন করতে বলা হবে।
"""
import sys

from werkzeug.security import generate_password_hash

from app import create_app
from db import get_db


def main():
    if len(sys.argv) != 2 or len(sys.argv[1]) < 6:
        print("ব্যবহার: python reset_admin.py <নতুন-পাসওয়ার্ড (কমপক্ষে ৬ অক্ষর)>")
        sys.exit(1)
    app = create_app()
    with app.app_context():
        conn = get_db()
        pw_hash = generate_password_hash(sys.argv[1])
        if conn.execute("SELECT 1 FROM users WHERE username = 'admin'").fetchone():
            conn.execute("UPDATE users SET password_hash = ?, role = 'admin', must_change = 1 WHERE username = 'admin'",
                         (pw_hash,))
        else:
            conn.execute("INSERT INTO users (username, password_hash, role, must_change) VALUES ('admin', ?, 'admin', 1)",
                         (pw_hash,))
        conn.commit()
        print(f"অ্যাডমিন পাসওয়ার্ড রিসেট হয়েছে। ডাটাবেস: {app.config['DATABASE']}")
        print("লগইন আইডি: admin — এখন ওয়েবসাইটের Web ট্যাবে Reload চাপুন, তারপর লগইন করুন।")


if __name__ == "__main__":
    main()
