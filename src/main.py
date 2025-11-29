import pandas as pd
from pathlib import Path
from datetime import date, datetime

# ---------- CONFIG ----------
MASTER_FILE = Path("data/master_attendance.xlsx")
SHEET_NAME = "Attendance"

# Change these two when you run for a new day
TODAY_STR = "2025-11-27"  # e.g. "2025-11-29"
DAILY_FILE = Path("data/daily_2025-11-27.csv")
# ----------------------------


def normalize_column_name(col) -> str:
    """
    Make sure column names are clean strings like '2025-11-27',
    not '2025-11-27 00:00:00' or Timestamp objects.
    """
    # If it's a pandas Timestamp or datetime/date → format as YYYY-MM-DD
    if isinstance(col, (pd.Timestamp, datetime, date)):
        return col.strftime("%Y-%m-%d")

    col_str = str(col).strip()

    # If it looks like 'YYYY-MM-DD 00:00:00' → keep only date part
    if " " in col_str and col_str[:10].count("-") == 2:
        # try to parse first part as a date
        first_part = col_str.split(" ")[0]
        try:
            datetime.strptime(first_part, "%Y-%m-%d")
            return first_part
        except ValueError:
            pass

    return col_str


def load_master(path: Path, sheet_name: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Master file not found: {path}")

    df = pd.read_excel(path, sheet_name=sheet_name)

    # Normalize all column names once when loading
    df.columns = [normalize_column_name(c) for c in df.columns]

    required_cols = {"Roll_No", "Name"}
    if not required_cols.issubset(df.columns):
        raise ValueError(f"Master sheet must contain columns: {required_cols}")
    return df

from openpyxl import load_workbook

def autofit_columns(excel_file, sheet_name):
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter  # A, B, C...

        for cell in column_cells:
            try:
                cell_value = str(cell.value)
                if cell_value:
                    max_length = max(max_length, len(cell_value))
            except:
                pass

        adjusted_width = max_length + 2  # extra padding
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(excel_file)


def load_daily(path: Path) -> set:
    if not path.exists():
        raise FileNotFoundError(f"Daily file not found: {path}")

    df = pd.read_csv(path)
    if "Roll_No" not in df.columns:
        raise ValueError("Daily file must contain column 'Roll_No'")
    return set(df["Roll_No"].tolist())


def update_attendance(master_df: pd.DataFrame, present_set: set, date_col: str) -> pd.DataFrame:
    # Normalize the date column name
    date_col = normalize_column_name(date_col)

    # Add date column if it doesn't exist
    if date_col not in master_df.columns:
        master_df[date_col] = ""

    # Mark P / A
    def mark(row):
        return "P" if row["Roll_No"] in present_set else "A"

    master_df[date_col] = master_df.apply(mark, axis=1)

    # Identify date columns (exclude fixed ones)
    base_columns = ["Roll_No", "Name", "Total_Present", "Percentage"]
    date_columns = [col for col in master_df.columns if col not in base_columns]

    # Count P for each row across all date columns
    master_df["Total_Present"] = (master_df[date_columns] == "P").sum(axis=1)

    # Total number of classes = number of date columns
    total_classes = len(date_columns)
    if total_classes > 0:
        master_df["Percentage"] = (master_df["Total_Present"] / total_classes * 100).round(2)
    else:
        master_df["Percentage"] = 0.0

    return master_df


def save_master(df: pd.DataFrame, path: Path, sheet_name: str):
    # Normalize columns again before saving (in case new ones were added)
    df.columns = [normalize_column_name(c) for c in df.columns]

    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)


def main():
    print("Loading master attendance...")
    master_df = load_master(MASTER_FILE, SHEET_NAME)

    print("Loading today's present students...")
    present_set = load_daily(DAILY_FILE)

    print(f"Updating attendance for {TODAY_STR} ...")
    updated_df = update_attendance(master_df, present_set, TODAY_STR)

    print("Saving updated master file...")
    save_master(updated_df, MASTER_FILE, SHEET_NAME)

    print("Auto-fitting column widths...")
    autofit_columns(MASTER_FILE, SHEET_NAME)

    print("Done. Master sheet updated.")


if __name__ == "__main__":
    main()
