@echo off
REM Weekly refresh for Windows. Builds output\Sales_Pack_<date>.xlsx and the Tableau feed.
REM Schedule this file in Windows Task Scheduler (see README, step 6).

REM Change this line if your data file has a different name.
set INPUT=data\raw\superstore.csv

cd /d "%~dp0.."
call .venv\Scripts\activate.bat
python -m sales_pack --input "%INPUT%" --out output
