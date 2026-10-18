#!/bin/bash
# PythonAnywhere এ নতুন আপডেট চালু করার জন্য:  bash ~/ncc-payroll/update.sh
# তথ্য (instance/payroll.db) অপরিবর্তিত থাকে — শুধু কোড হালনাগাদ হয়।
set -e
cd "$(dirname "$0")"
mkdir -p backups
[ -f instance/payroll.db ] && cp instance/payroll.db "backups/payroll-before-update-$(date +%F-%H%M).db"
git pull --ff-only
.venv/bin/pip install -q -r requirements.txt
# WSGI ফাইল touch করলে PythonAnywhere সাইটটি নিজে Reload করে
touch /var/www/*_pythonanywhere_com_wsgi.py 2>/dev/null && echo "সাইট রিলোড হয়েছে।" || echo "Web ট্যাব থেকে Reload চাপুন।"
echo "আপডেট সম্পন্ন ✔"
