\# Attendance Automation



A Python script that automatically updates a master attendance Excel sheet using daily CSV files of present students.



\## Features



\- Reads master attendance sheet (Excel)

\- Reads daily present-student list (CSV)

\- Marks \*\*P\*\* (present) or \*\*A\*\* (absent)

\- Creates a new date column if missing

\- Updates total presents

\- Calculates attendance percentage

\- Saves updated Excel file



\## How to Run



1\. Put your master file in `data/master\_attendance.xlsx`

2\. Put your daily file in `data/daily\_YYYY-MM-DD.csv`

3\. Open `src/main.py` and update:

&nbsp;  - `DAILY\_FILE`

&nbsp;  - `TODAY\_STR`

4\. Run:



```bash

python src/main.py



