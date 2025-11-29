Attendance Automation

A Python tool to automate daily attendance updates in an Excel-based master sheet.

Instead of manually marking P/A for every student each day, you just provide:

A master Excel file with all students

A daily CSV file with roll numbers of present students

The script:

Adds/updates a column for the given date

Marks P (present) / A (absent)

Recalculates Total_Present and Percentage

Creates a timestamped backup before overwriting the master file

Features

Command-line arguments for:

--date

--daily

Auto-detects default daily file based on date

Automatically creates backup files

Supports unlimited date columns

Works with simple Excel/CSV formats

Folder Structure
attendance-automation/
├─ data/
│  ├─ master_attendance.xlsx     
│  └─ daily_YYYY-MM-DD.csv       
├─ src/
│  └─ main.py                    
├─ backups/                      
├─ requirements.txt
├─ .gitignore
└─ README.md

Requirements

Python 3.10+

Install dependencies:

pip install -r requirements.txt


Packages used:

pandas

openpyxl

Master File Format (master_attendance.xlsx)

Sheet name must be: Attendance

Required columns:

Roll_No	Name

Attendance date columns must be in: YYYY-MM-DD format

Script will manage:

Total_Present	Percentage
Example:
Roll_No	Name	2025-11-25	2025-11-26	Total_Present	Percentage
1	Student One	P	A	1	50.0
2	Student Two	A	P	1	50.0
Daily File Format (daily_YYYY-MM-DD.csv)

CSV file

Required column:

Roll_No
Example:
Roll_No
1
3
5


This means roll numbers 1, 3, 5 are present.

How to Run
1. Default mode (uses today’s date)
python src/main.py


Script will look for:

data/daily_<today>.csv

2. Specify custom date
python src/main.py --date 2025-11-28


Uses:

data/daily_2025-11-28.csv

3. Specify custom daily file
python src/main.py --daily data/my_custom_file.csv

4. Specify both
python src/main.py --date 2025-11-28 --daily data/daily_2025-11-28.csv

Backups

Before updating the master file, the script creates a backup in:

backups/


Backup filename format:

master_attendance_2025-11-28_134501.xlsx


This protects you from accidental errors.

Example Run (Terminal Output)
Using date: 2025-11-28
Using daily file: data/daily_2025-11-28.csv
Loading master attendance...
Loading today's present students...
Creating backup of master file...
Backup created at: backups/master_attendance_2025-11-28_134501.xlsx
Updating attendance for 2025-11-28 ...
Saving updated master file...
Done. Master sheet updated.

Future Improvements

Validation and cleaner error messages

Web UI or Google Sheets integration

Auto-fetching attendance from biometric/QR systems

Summary report generation

Unit tests for attendance logic