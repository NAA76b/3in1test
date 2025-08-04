"""MASTER 3-in-1 Tool v2

Light refactor of the original `MASTER_3in1Tool_v1.py` focusing on:
1. Centralised helper utilities (`normalize_name`, `format_output_filename`).
2. Hardened name matching logic using the shared normaliser.
3. Centralised output-file naming.

The UI and the majority of the business logic remain unchanged so that
existing users feel no disruption.  Further modern-UI work will be added
incrementally (see project TODO list).
"""

from __future__ import annotations

import os
import threading
from datetime import datetime

import pandas as pd
from tkinter import messagebox

# Re-use all classes/functions from v1 as a starting point
from MASTER_3in1Tool_v1 import (
    FormatterFrame,
    EmailerFrame,
    get_default_output_path,
    log_message_util,
)

# Import the shared helpers we just created
from utils_threeinone import normalize_name, format_output_filename

# -----------------------------------------------------
# ReporterFrame (V2) – subclass the original to inject fixes
# -----------------------------------------------------

from MASTER_3in1Tool_v1 import ReporterFrame as _BaseReporterFrame


class ReporterFrame(_BaseReporterFrame):
    """Drop-in replacement for the original ReporterFrame with
    improved name matching & output file naming.
    """

    # --- Improved Name Matching ------------------------------------------------
    def validate_employee_ids(self, df1: pd.DataFrame, df2: pd.DataFrame):  # type: ignore[override]
        """Match two dataframes on normalised owner names.

        Returns: (matched, only_in_df1, only_in_df2)
        """
        df1 = df1.copy()
        df2 = df2.copy()
        df1["Name_Standard"] = df1["Time Sheet Owner Name"].apply(normalize_name)
        df2["Name_Standard"] = df2["Time Sheet Owner Name"].apply(normalize_name)

        matched = pd.merge(
            df1,
            df2,
            on="Name_Standard",
            how="inner",
            suffixes=("_hist", "_emp"),
        )
        only_in_file1 = df1[~df1["Name_Standard"].isin(df2["Name_Standard"])]
        only_in_file2 = df2[~df2["Name_Standard"].isin(df1["Name_Standard"])]
        return matched, only_in_file1, only_in_file2

    # --- Centralised Output File Naming ---------------------------------------
    def process_files_thread(self):  # type: ignore[override]
        try:
            history_df = self.clean_report_data(self.history_file.get())
            employee_df = self.clean_report_data(self.employee_file.get())

            matched_df, unmatched_hist, unmatched_emp = self.validate_employee_ids(
                history_df, employee_df
            )
            log_message_util(
                self.app.root,
                self.log_text,
                f"Reporter: Found {len(matched_df)} matched employees.",
            )

            improved, declined, unchanged = self.track_compliance_changes(matched_df)
            log_message_util(
                self.app.root,
                self.log_text,
                f"Reporter: {len(improved)} employees improved compliance.",
            )

            output_dir = self.reporter_output_location.get()

            matched_filename = os.path.join(
                output_dir, format_output_filename("matched_report")
            )
            unmatched_hist.to_excel(
                os.path.join(output_dir, format_output_filename("unmatched_history")),
                index=False,
            )
            unmatched_emp.to_excel(
                os.path.join(output_dir, format_output_filename("unmatched_employee")),
                index=False,
            )

            with pd.ExcelWriter(matched_filename) as writer:
                matched_df.to_excel(writer, sheet_name="Matched_Employees", index=False)
                improved.to_excel(writer, sheet_name="Improved_Compliance", index=False)
                declined.to_excel(writer, sheet_name="Declined_Compliance", index=False)

            self.app.root.after(0, self.process_complete, output_dir)
        except Exception as e:  # pragma: no-cover – Swallow & log all
            self.app.root.after(0, self.process_error, e)


# -----------------------------------------------------
# Application shell that swaps in the new ReporterFrame
# -----------------------------------------------------

import tkinter as tk
from tkinter import ttk


class MasterWorkflowApp:  # noqa: D101 – top-level app class
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Master Workflow Tool v2")
        self.root.geometry("950x850")
        self.root.configure(bg="#2E2E2E")

        # --- Style & theme setup (re-use v1 helpers) --------------------------
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.BG_COLOR = "#2E2E2E"
        self.FG_COLOR = "#FFFFFF"
        self.ACCENT_COLOR_ORANGE = "#FF6600"
        self.ACCENT_COLOR_BLUE = "#007ACC"
        self.ACCENT_COLOR_TEAL = "#2a9d8f"
        self.WIDGET_BG = "#3C3C3C"

        self.style.configure(".", background=self.BG_COLOR, foreground=self.FG_COLOR)
        self.style.configure("TFrame", background=self.BG_COLOR)
        self.style.configure(
            "TLabel",
            background=self.BG_COLOR,
            foreground=self.FG_COLOR,
            font=("Helvetica", 12),
        )
        self.style.configure(
            "Info.TLabel",
            background=self.BG_COLOR,
            foreground="#CCCCCC",
            font=("Helvetica", 9),
        )
        self.style.configure("Header.TLabel", font=("Helvetica", 18, "bold"))
        self.style.configure(
            "TButton",
            background=self.WIDGET_BG,
            foreground=self.FG_COLOR,
            font=("Helvetica", 11),
            borderwidth=1,
        )
        self.style.map("TButton", background=[("active", "#4a4a4a")])

        self.style.configure(
            "Orange.TButton",
            background=self.ACCENT_COLOR_ORANGE,
            foreground=self.FG_COLOR,
            font=("Helvetica", 14, "bold"),
        )
        self.style.map("Orange.TButton", background=[("active", "#FF8533")])
        self.style.configure(
            "Blue.TButton",
            background=self.ACCENT_COLOR_BLUE,
            foreground=self.FG_COLOR,
            font=("Helvetica", 14, "bold"),
        )
        self.style.map("Blue.TButton", background=[("active", "#0099FF")])
        self.style.configure(
            "Teal.TButton",
            background=self.ACCENT_COLOR_TEAL,
            foreground=self.FG_COLOR,
            font=("Helvetica", 14, "bold"),
        )
        self.style.map("Teal.TButton", background=[("active", "#34c2b2")])

        self.style.configure(
            "TCombobox",
            fieldbackground=self.WIDGET_BG,
            background=self.WIDGET_BG,
            foreground=self.FG_COLOR,
            arrowcolor=self.FG_COLOR,
        )
        self.style.configure(
            "TEntry",
            fieldbackground=self.WIDGET_BG,
            foreground=self.FG_COLOR,
            borderwidth=1,
        )
        self.style.configure(
            "Orange.Horizontal.TProgressbar",
            troughcolor=self.WIDGET_BG,
            background=self.ACCENT_COLOR_ORANGE,
            thickness=20,
        )
        self.style.configure(
            "Blue.Horizontal.TProgressbar",
            troughcolor=self.WIDGET_BG,
            background=self.ACCENT_COLOR_BLUE,
            thickness=20,
        )
        self.style.configure(
            "Teal.Horizontal.TProgressbar",
            troughcolor=self.WIDGET_BG,
            background=self.ACCENT_COLOR_TEAL,
            thickness=20,
        )
        self.style.configure("TLabelframe", background=self.BG_COLOR, bordercolor=self.WIDGET_BG)
        self.style.configure(
            "TLabelframe.Label",
            background=self.BG_COLOR,
            foreground=self.FG_COLOR,
            font=("Helvetica", 12, "bold"),
        )
        self.style.configure("TNotebook", background=self.BG_COLOR, borderwidth=0)
        self.style.configure(
            "TNotebook.Tab",
            background=self.WIDGET_BG,
            foreground=self.FG_COLOR,
            padding=[10, 5],
            font=("Helvetica", 11, "bold"),
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", self.ACCENT_COLOR_BLUE)],
            foreground=[("selected", self.FG_COLOR)],
        )
        self.style.configure(
            "TRadiobutton",
            background=self.BG_COLOR,
            foreground=self.FG_COLOR,
            font=("Helvetica", 11),
        )
        self.style.map(
            "TRadiobutton",
            indicatorcolor=[("selected", self.ACCENT_COLOR_TEAL)],
            background=[("active", self.BG_COLOR)],
        )

        # --- Shared Data ------------------------------------------------------
        self.default_output_path = get_default_output_path()
        from MASTER_3in1Tool_v1 import MasterWorkflowApp as _LegacyApp  # for office map

        self.office_name_map = _LegacyApp().office_name_map  # type: ignore[attr-defined]

        # --- Notebook & Tabs ---------------------------------------------------
        notebook = ttk.Notebook(self.root, style="TNotebook")
        notebook.pack(expand=True, fill="both", padx=10, pady=10)

        reporter_tab = ReporterFrame(notebook, self, padding="20")
        formatter_tab = FormatterFrame(notebook, self, padding="20")
        emailer_tab = EmailerFrame(notebook, self, padding="20")

        notebook.add(reporter_tab, text="Pay Period Reporter")
        notebook.add(formatter_tab, text="Helper Sheet Formatter")
        notebook.add(emailer_tab, text="Office Emailer")

    # ---------------------------------------------------------------------
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = MasterWorkflowApp()
    app.run()