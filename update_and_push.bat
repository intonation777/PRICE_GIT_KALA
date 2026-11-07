@echo off
cd /d "%~dp0"
python update_data.py
if errorlevel 1 (
  echo Update failed. Nothing was pushed.
  exit /b 1
)
git add docs
git commit -m "update prices"
git push
