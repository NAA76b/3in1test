"""Utility to combine History and Employee ID Excel sheets.

Reads a pay period "History" Excel file and an "Employee ID" Excel file,
matches employees by name, and outputs a single Excel sheet that retains the
columns from the history data while inserting the corresponding Employee ID.

The output also duplicates the timesheet owner name next to the new Employee ID
column for easy visual verification.

Usage::

    python merge_history_employee.py --history History.xlsx --employees Employees.xlsx --output Combined.xlsx

Both input files are expected to have a column for the employee name.  The
history file typically contains a column named "Timesheet Owner Name" while the
employee file may use either "Timesheet Owner Name" or "Employee Name".
"""

from __future__ import annotations

import argparse
from pathlib import Path
import pandas as pd


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip surrounding whitespace from column names."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def merge_history_employee(
    history_path: Path, employee_path: Path, output_path: Path
) -> None:
    """Merge history and employee sheets and write a single consolidated file.

    Parameters
    ----------
    history_path: Path
        Path to the history Excel file.
    employee_path: Path
        Path to the employee ID Excel file.
    output_path: Path
        Destination for the combined Excel sheet.
    """

    history_df = _normalize_columns(pd.read_excel(history_path))
    employee_df = _normalize_columns(pd.read_excel(employee_path))

    # Determine which column in the employee sheet contains the name.
    name_col_emp = (
        "Timesheet Owner Name"
        if "Timesheet Owner Name" in employee_df.columns
        else "Employee Name"
    )
    if name_col_emp not in employee_df.columns:
        raise KeyError(
            "Employee file must contain a 'Timesheet Owner Name' or 'Employee Name' column"
        )

    # Clean up name columns for matching.
    history_df["Timesheet Owner Name"] = (
        history_df["Timesheet Owner Name"].astype(str).str.strip()
    )
    employee_df[name_col_emp] = employee_df[name_col_emp].astype(str).str.strip()

    merged = history_df.merge(
        employee_df[[name_col_emp, "Employee ID"]],
        left_on="Timesheet Owner Name",
        right_on=name_col_emp,
        how="left",
    )

    # Drop extra name column from merge and insert Employee ID next to the name.
    merged.drop(columns=[name_col_emp], inplace=True)
    name_index = merged.columns.get_loc("Timesheet Owner Name")
    id_col = merged.pop("Employee ID")
    merged.insert(name_index + 1, "Employee ID", id_col)

    # Duplicate the Timesheet Owner Name column for verification.
    merged.insert(name_index + 2, "Timesheet Owner Name Duplicate", merged["Timesheet Owner Name"])

    merged.to_excel(output_path, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--history", required=True, type=Path, help="History Excel file")
    parser.add_argument(
        "--employees", required=True, type=Path, help="Employee ID Excel file"
    )
    parser.add_argument("--output", required=True, type=Path, help="Output Excel file")
    args = parser.parse_args()

    merge_history_employee(args.history, args.employees, args.output)


if __name__ == "__main__":  # pragma: no cover - direct execution only
    main()
