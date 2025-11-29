import argparse
from pathlib import Path
from datetime import date, datetime
import shutil
import pandas as pd

# ---------- CONFIG ----------
MASTER_FILE = Path("data/master_attendance.xlsx")
SHEET_NAME = "Attendance"
BACKUP_DIR = Path("backups")
# ----------------------------


def load_master(path: Path, sheet_name: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Master file not found: {path}")

    df = pd.read_excel(path, sheet_name=sheet_name)
    required_cols = {"Roll_No", "Name"}
    if not required_cols.issubset(df.columns):
        raise ValueError(f"Master sheet must contain columns: {required_cols}")
    return df


def load_daily(path: Path) -> set:
    if not path.exists():
        raise FileNotFoundError(f"Daily file not found: {path}")

    df = pd.read_csv(path)
    if "Roll_No" not in df.columns:
        raise ValueError("Daily file must contain column 'Roll_No'")
    return set(df["Roll_No"].tolist())


def update_attendance(master_df: pd.DataFrame, present_set: set, date_col: str) -> pd.DataFrame:
    # Add date column if it doesn't exist
    if date_col not in master_df.columns:
        master_df[date_col] = ""

    # Mark P / A
    def mark(row):
        return "P" if row["Roll_No"] in present_set else "A"

    master_df[date_col] = master_df.apply(mark, axis=1)

    # Identify date columns (exclude fixed ones)
    date_columns = [
        col for col in master_df.columns
        if col not in ["Roll_No", "Name", "Total_Present", "Percentage"]
    ]

    # Count P for each row
    master_df["Total_Present"] = (master_df[date_columns] == "P").sum(axis=1)

    # Total number of classes = number of date columns
    total_classes = len(date_columns)
    if total_classes > 0:
        master_df["Percentage"] = (master_df["Total_Present"] / total_classes * 100).round(2)
    else:
        master_df["Percentage"] = 0.0

    return master_df


def backup_master(master_path: Path, backup_dir: Path):
    """
    Create a timestamped backup of the master file in backup_dir.
    Example: backups/master_attendance_2025-11-28_134501.xlsx
    """
    if not master_path.exists():
        print(f"No master file found at {master_path}, skipping backup.")
        return

    backup_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_name = f"{master_path.stem}_{timestamp}{master_path.suffix}"
    backup_path = backup_dir / backup_name

    shutil.copy2(master_path, backup_path)
    print(f"Backup created at: {backup_path}")


def save_master(df: pd.DataFrame, path: Path, sheet_name: str):
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Automate attendance updates in an Excel master sheet."
    )
    parser.add_argument(
        "--date",
        required=False,
        help="Date for attendance in YYYY-MM-DD format. Default: today.",
    )
    parser.add_argument(
        "--daily",
        required=False,
        help="Path to daily CSV file containing present students (column: Roll_No). "
             "Default: data/daily_<date>.csv",
    )
    args = parser.parse_args()

    # Handle default date = today
    if args.date is None:
        today = date.today()
        date_str = today.isoformat()  # YYYY-MM-DD
    else:
        date_str = args.date

    # Handle default daily path = data/daily_<date>.csv
    if args.daily is None:
        daily_path = Path(f"data/daily_{date_str}.csv")
    else:
        daily_path = Path(args.daily)

    return date_str, daily_path


def main():
    # Parse CLI arguments
    date_str, daily_file = parse_args()

    print(f"Using date: {date_str}")
    print(f"Using daily file: {daily_file}")

    print("Loading master attendance...")
    master_df = load_master(MASTER_FILE, SHEET_NAME)

    print("Loading today's present students...")
    present_set = load_daily(daily_file)

    print("Creating backup of master file...")
    backup_master(MASTER_FILE, BACKUP_DIR)

    print(f"Updating attendance for {date_str} ...")
    updated_df = update_attendance(master_df, present_set, date_str)

    print("Saving updated master file...")
    save_master(updated_df, MASTER_FILE, SHEET_NAME)

    print("Done. Master sheet updated.")


if __name__ == "__main__":
    main()
