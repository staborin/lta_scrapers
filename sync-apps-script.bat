@echo off
cd /d "%~dp0apps-script"
clasp pull
cd /d "%~dp0"
git add .
git commit -m "Update Apps Script files"
git push