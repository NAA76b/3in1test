# 3in1test

This repository contains experimental tools for merging pay period reports.

## merge_history_employee.py

`merge_history_employee.py` combines a pay period **History** sheet with an
**Employee ID** sheet. The resulting Excel file keeps the original columns from
the History data while inserting the matching **Employee ID** and a duplicate of
the **Timesheet Owner Name** for quick verification.

Run the tool:

```
python merge_history_employee.py --history History.xlsx --employees EmployeeIDs.xlsx --output Combined.xlsx
```

The output is a single Excel file with all data merged onto one sheet.
