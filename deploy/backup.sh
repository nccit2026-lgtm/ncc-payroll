#!/bin/sh
# প্রতিদিনের ব্যাকআপ (crontab: 0 2 * * * /opt/ncc-payroll/deploy/backup.sh)
set -e
DIR=/opt/ncc-payroll
mkdir -p "$DIR/backups"
sqlite3 "$DIR/instance/payroll.db" ".backup '$DIR/backups/payroll-$(date +%F).db'"
find "$DIR/backups" -name 'payroll-*.db' -mtime +60 -delete
