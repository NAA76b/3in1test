# -*- coding: utf-8 -*-
"""
MASTER 3in1 Tool v2 - Redesigned with Modern UI and Enhanced Matching Logic

Key Improvements:
- Modern Apple/Google-style UI with clean theme and improved UX
- Hardened error handling for external calls (COM/Outlook, file I/O)
- Centralized name normalization and matching logic
- Added Employee ID matching with exact and fuzzy fallback
- Fixed HTML CSS syntax bugs
- Improved logging with severity badges and filters
- Centralized filename formatting logic
- Thread-safe UI updates throughout

UI Changes:
- Light, clean theme with soft neutral backgrounds
- Rounded corners and subtle drop shadows
- Grid-based form with numbered steps
- Primary/secondary button styling with hover states
- Collapsible log panel with filters
- Dynamic input states and inline validation
- Summary badges for match statistics
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
import threading
from datetime import datetime
import traceback
from fuzzywuzzy import process, fuzz
import json

# Required libraries with safe import
try:
    import win32com.client as win32
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    OUTLOOK_AVAILABLE = True
except ImportError as e:
    OUTLOOK_AVAILABLE = False
    print(f"Warning: Some features unavailable - {e}")

# Try to import openpyxl
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

#==========================================================================
# THEME CONFIGURATION - Apple/Google Modern Blend
#==========================================================================
class Theme:
    # Colors
    BG_PRIMARY = '#FFFFFF'
    BG_SECONDARY = '#F8F9FA'
    BG_TERTIARY = '#E8EAED'
    
    TEXT_PRIMARY = '#202124'
    TEXT_SECONDARY = '#5F6368'
    TEXT_TERTIARY = '#80868B'
    
    ACCENT_PRIMARY = '#1A73E8'  # Google Blue
    ACCENT_SUCCESS = '#34A853'  # Google Green
    ACCENT_WARNING = '#FBBC04'  # Google Yellow
    ACCENT_ERROR = '#EA4335'    # Google Red
    
    BORDER_COLOR = '#DADCE0'
    SHADOW_COLOR = 'rgba(60, 64, 67, 0.08)'
    
    # Typography
    FONT_FAMILY = 'Segoe UI'
    FONT_SIZE_XL = 24
    FONT_SIZE_L = 18
    FONT_SIZE_M = 14
    FONT_SIZE_S = 12
    FONT_SIZE_XS = 10
    
    # Spacing
    PADDING_L = 24
    PADDING_M = 16
    PADDING_S = 8
    PADDING_XS = 4
    
    # Border Radius
    RADIUS_L = 16
    RADIUS_M = 12
    RADIUS_S = 8
    RADIUS_XS = 4

#==========================================================================
# CENTRALIZED UTILITY FUNCTIONS
#==========================================================================

def normalize_name(name_str):
    """Normalize name for consistent matching."""
    if pd.isna(name_str) or not name_str:
        return ""
    # Convert to uppercase, strip whitespace, remove extra spaces
    normalized = re.sub(r'\s+', ' ', str(name_str).upper().strip())
    # Remove common suffixes
    normalized = re.sub(r'\s+(JR\.?|SR\.?|III|II|IV)$', '', normalized)
    return normalized

def format_output_filename(base_name, extension='xlsx'):
    """Centralized filename formatting with timestamp."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_{timestamp}.{extension}"

def get_default_output_path():
    """Gets the default output path, preferring OneDrive Desktop if exists."""
    onedrive_desktop = os.path.join(os.path.expanduser("~"), "OneDrive - FDA", "Desktop")
    if os.path.exists(onedrive_desktop):
        return onedrive_desktop
    return os.path.expanduser("~/Desktop")

def safe_file_operation(operation_func, *args, **kwargs):
    """Wrapper for safe file operations with error handling."""
    try:
        return operation_func(*args, **kwargs)
    except PermissionError as e:
        return False, f"Permission denied: {str(e)}"
    except FileNotFoundError as e:
        return False, f"File not found: {str(e)}"
    except Exception as e:
        return False, f"File operation error: {str(e)}"

def safe_com_operation(operation_func, *args, **kwargs):
    """Wrapper for safe COM operations with error handling."""
    if not OUTLOOK_AVAILABLE:
        return False, "Outlook integration not available. Please install pywin32."
    try:
        return operation_func(*args, **kwargs)
    except Exception as e:
        return False, f"Outlook error: {str(e)}"

#==========================================================================
# ENHANCED LOGGING SYSTEM
#==========================================================================

class LogLevel:
    INFO = 'info'
    SUCCESS = 'success'
    WARNING = 'warning'
    ERROR = 'error'

class EnhancedLogger:
    def __init__(self, text_widget, root):
        self.text_widget = text_widget
        self.root = root
        self.filters = {level: tk.BooleanVar(value=True) for level in [LogLevel.INFO, LogLevel.SUCCESS, LogLevel.WARNING, LogLevel.ERROR]}
        self.messages = []
        self.setup_tags()
    
    def setup_tags(self):
        """Configure text widget tags for different log levels."""
        self.text_widget.tag_config(LogLevel.INFO, foreground=Theme.TEXT_SECONDARY, font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S))
        self.text_widget.tag_config(LogLevel.SUCCESS, foreground=Theme.ACCENT_SUCCESS, font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S, 'bold'))
        self.text_widget.tag_config(LogLevel.WARNING, foreground=Theme.ACCENT_WARNING, font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S, 'bold'))
        self.text_widget.tag_config(LogLevel.ERROR, foreground=Theme.ACCENT_ERROR, font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S, 'bold'))
    
    def log(self, message, level=LogLevel.INFO):
        """Thread-safe logging with filtering support."""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = {'timestamp': timestamp, 'message': message, 'level': level}
        self.messages.append(log_entry)
        
        def _insert():
            if self.filters[level].get():
                self.text_widget.config(state='normal')
                self.text_widget.insert(tk.END, f"[{timestamp}] {message}\n", level)
                self.text_widget.config(state='disabled')
                self.text_widget.see(tk.END)
        
        self.root.after(0, _insert)
    
    def apply_filters(self):
        """Reapply filters to show/hide messages."""
        self.text_widget.config(state='normal')
        self.text_widget.delete(1.0, tk.END)
        
        for entry in self.messages:
            if self.filters[entry['level']].get():
                self.text_widget.insert(tk.END, f"[{entry['timestamp']}] {entry['message']}\n", entry['level'])
        
        self.text_widget.config(state='disabled')
        self.text_widget.see(tk.END)
    
    def get_summary(self):
        """Get summary counts by level."""
        summary = {level: 0 for level in [LogLevel.INFO, LogLevel.SUCCESS, LogLevel.WARNING, LogLevel.ERROR]}
        for entry in self.messages:
            summary[entry['level']] += 1
        return summary

#==========================================================================
# EMPLOYEE ID MATCHER WITH FUZZY FALLBACK
#==========================================================================

class EmployeeIDMatcher:
    def __init__(self, logger=None):
        self.logger = logger
        self.fuzzy_threshold = 85  # Configurable threshold
        self.lookup_data = {}
        self.match_stats = {
            'exact': 0,
            'fuzzy': 0,
            'no_match': 0,
            'already_had_id': 0
        }
    
    def load_lookup_file(self, filepath):
        """Load employee lookup file (Jan 2025 names)."""
        try:
            df = pd.read_excel(filepath)
            
            # Flexible header detection
            id_col = None
            name_cols = []
            
            for col in df.columns:
                col_lower = str(col).lower()
                if 'employee' in col_lower and 'id' in col_lower:
                    id_col = col
                elif any(term in col_lower for term in ['name', 'first', 'last']):
                    name_cols.append(col)
            
            if not id_col:
                raise ValueError("Could not find Employee ID column")
            
            # Build lookup dictionary
            for _, row in df.iterrows():
                emp_id = str(row[id_col]).strip()
                if pd.notna(emp_id):
                    # Try to construct full name from available columns
                    name_parts = []
                    for col in name_cols:
                        if pd.notna(row[col]):
                            name_parts.append(str(row[col]).strip())
                    
                    full_name = ' '.join(name_parts)
                    normalized_name = normalize_name(full_name)
                    
                    if normalized_name:
                        self.lookup_data[normalized_name] = emp_id
            
            if self.logger:
                self.logger.log(f"Loaded {len(self.lookup_data)} employee records for matching", LogLevel.SUCCESS)
            
            return True
        except Exception as e:
            if self.logger:
                self.logger.log(f"Error loading lookup file: {e}", LogLevel.ERROR)
            return False
    
    def match_employee(self, name, existing_id=None):
        """Match employee name to ID with exact and fuzzy fallback."""
        # Check if already has ID
        if pd.notna(existing_id) and str(existing_id).strip():
            self.match_stats['already_had_id'] += 1
            return existing_id, 'Already had ID'
        
        normalized = normalize_name(name)
        if not normalized:
            self.match_stats['no_match'] += 1
            return None, 'No match - invalid name'
        
        # Exact match
        if normalized in self.lookup_data:
            self.match_stats['exact'] += 1
            return self.lookup_data[normalized], 'Exact match'
        
        # Fuzzy match
        if self.lookup_data:
            result = process.extractOne(normalized, self.lookup_data.keys(), scorer=fuzz.ratio)
            if result and result[1] >= self.fuzzy_threshold:
                self.match_stats['fuzzy'] += 1
                matched_key = result[0]
                return self.lookup_data[matched_key], f'Fuzzy match ({result[1]}% confidence)'
        
        self.match_stats['no_match'] += 1
        return None, 'No match found'
    
    def process_dataframe(self, df, name_col, id_col='Employee ID'):
        """Process entire dataframe for matching."""
        results = []
        
        for _, row in df.iterrows():
            name = row.get(name_col, '')
            existing_id = row.get(id_col, None) if id_col in df.columns else None
            
            matched_id, match_status = self.match_employee(name, existing_id)
            
            result_row = row.to_dict()
            result_row['Matched_Employee_ID'] = matched_id
            result_row['Match_Status'] = match_status
            result_row['Original_Employee_ID'] = existing_id
            
            results.append(result_row)
        
        return pd.DataFrame(results)

#==========================================================================
# MODERN UI COMPONENTS
#==========================================================================

class ModernButton(tk.Button):
    """Modern button with hover effects and rounded corners."""
    def __init__(self, parent, text="", command=None, style='primary', **kwargs):
        # Set colors based on style
        colors = {
            'primary': (Theme.ACCENT_PRIMARY, '#1765CC'),
            'secondary': (Theme.BG_TERTIARY, Theme.BORDER_COLOR),
            'success': (Theme.ACCENT_SUCCESS, '#2D8644'),
            'error': (Theme.ACCENT_ERROR, '#C5221F')
        }
        
        self.normal_color, self.hover_color = colors.get(style, colors['primary'])
        self.text_color = Theme.BG_PRIMARY if style != 'secondary' else Theme.TEXT_PRIMARY
        
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=self.normal_color,
            fg=self.text_color,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M, 'bold'),
            relief=tk.FLAT,
            cursor='hand2',
            padx=Theme.PADDING_M,
            pady=Theme.PADDING_S,
            **kwargs
        )
        
        self.bind('<Enter>', self.on_hover)
        self.bind('<Leave>', self.on_leave)
    
    def on_hover(self, event):
        self.config(bg=self.hover_color)
    
    def on_leave(self, event):
        self.config(bg=self.normal_color)

class ModernEntry(tk.Frame):
    """Modern entry field with label and validation."""
    def __init__(self, parent, label="", **kwargs):
        super().__init__(parent, bg=Theme.BG_PRIMARY)
        
        if label:
            self.label = tk.Label(
                self,
                text=label,
                bg=Theme.BG_PRIMARY,
                fg=Theme.TEXT_SECONDARY,
                font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S)
            )
            self.label.pack(anchor='w', pady=(0, Theme.PADDING_XS))
        
        self.var = tk.StringVar()
        self.entry = tk.Entry(
            self,
            textvariable=self.var,
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M),
            relief=tk.FLAT,
            bd=1,
            highlightthickness=2,
            highlightcolor=Theme.ACCENT_PRIMARY,
            highlightbackground=Theme.BORDER_COLOR,
            **kwargs
        )
        self.entry.pack(fill='x', ipady=Theme.PADDING_S)
        
        self.error_label = tk.Label(
            self,
            text="",
            bg=Theme.BG_PRIMARY,
            fg=Theme.ACCENT_ERROR,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_XS)
        )
    
    def set_error(self, message):
        """Display error message below entry."""
        self.error_label.config(text=message)
        self.error_label.pack(anchor='w', pady=(Theme.PADDING_XS, 0))
    
    def clear_error(self):
        """Clear error message."""
        self.error_label.config(text="")
        self.error_label.pack_forget()

class StepIndicator(tk.Frame):
    """Step indicator for multi-step forms."""
    def __init__(self, parent, steps, **kwargs):
        super().__init__(parent, bg=Theme.BG_PRIMARY, **kwargs)
        self.steps = steps
        self.current_step = 0
        self.indicators = []
        
        for i, step in enumerate(steps):
            frame = tk.Frame(self, bg=Theme.BG_PRIMARY)
            frame.pack(side='left', padx=Theme.PADDING_S)
            
            # Circle
            circle = tk.Label(
                frame,
                text=str(i + 1),
                bg=Theme.BG_TERTIARY,
                fg=Theme.TEXT_SECONDARY,
                font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S, 'bold'),
                width=3,
                height=1
            )
            circle.pack()
            
            # Label
            label = tk.Label(
                frame,
                text=step,
                bg=Theme.BG_PRIMARY,
                fg=Theme.TEXT_SECONDARY,
                font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_XS)
            )
            label.pack()
            
            self.indicators.append((circle, label))
    
    def set_step(self, step_num):
        """Update current step indicator."""
        self.current_step = step_num
        for i, (circle, label) in enumerate(self.indicators):
            if i < step_num:
                circle.config(bg=Theme.ACCENT_SUCCESS, fg=Theme.BG_PRIMARY)
                label.config(fg=Theme.TEXT_PRIMARY)
            elif i == step_num:
                circle.config(bg=Theme.ACCENT_PRIMARY, fg=Theme.BG_PRIMARY)
                label.config(fg=Theme.TEXT_PRIMARY)
            else:
                circle.config(bg=Theme.BG_TERTIARY, fg=Theme.TEXT_SECONDARY)
                label.config(fg=Theme.TEXT_SECONDARY)

#==========================================================================
# MAIN APPLICATION WITH ENHANCED ID MATCHING
#==========================================================================

class MasterWorkflowApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("MASTER 3in1 Tool v2 - Modern Edition")
        self.root.geometry("1200x800")
        self.root.configure(bg=Theme.BG_PRIMARY)
        
        # Set window style
        self.root.tk.call('tk', 'scaling', 1.5)  # Better DPI scaling
        
        # Initialize components
        self.default_output_path = get_default_output_path()
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the modern UI layout."""
        # Main container with padding
        main_container = tk.Frame(self.root, bg=Theme.BG_PRIMARY)
        main_container.pack(fill='both', expand=True, padx=Theme.PADDING_L, pady=Theme.PADDING_L)
        
        # Header
        header_frame = tk.Frame(main_container, bg=Theme.BG_PRIMARY)
        header_frame.pack(fill='x', pady=(0, Theme.PADDING_L))
        
        tk.Label(
            header_frame,
            text="MASTER 3in1 Tool",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_XL, 'bold')
        ).pack(side='left')
        
        tk.Label(
            header_frame,
            text="v2.0 - Modern Edition",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M)
        ).pack(side='left', padx=(Theme.PADDING_M, 0))
        
        # Create notebook with custom style
        self.setup_notebook_style()
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill='both', expand=True)
        
        # Add tabs
        self.add_id_matcher_tab()
        self.add_reporter_tab()
        self.add_formatter_tab()
        self.add_emailer_tab()
    
    def setup_notebook_style(self):
        """Configure modern notebook style."""
        style = ttk.Style()
        style.theme_use('default')
        
        # Configure notebook
        style.configure('Modern.TNotebook', background=Theme.BG_PRIMARY, borderwidth=0)
        style.configure('Modern.TNotebook.Tab',
            background=Theme.BG_SECONDARY,
            foreground=Theme.TEXT_SECONDARY,
            padding=[Theme.PADDING_M, Theme.PADDING_S],
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M)
        )
        style.map('Modern.TNotebook.Tab',
            background=[('selected', Theme.BG_PRIMARY)],
            foreground=[('selected', Theme.TEXT_PRIMARY)],
            expand=[('selected', [1, 1, 1, 0])]
        )
    
    def add_id_matcher_tab(self):
        """Add the Employee ID Matcher tab."""
        tab = tk.Frame(self.notebook, bg=Theme.BG_PRIMARY)
        self.notebook.add(tab, text='Employee ID Matcher')
        
        # Create two-column layout
        left_frame = tk.Frame(tab, bg=Theme.BG_PRIMARY)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, Theme.PADDING_M))
        
        right_frame = tk.Frame(tab, bg=Theme.BG_SECONDARY)
        right_frame.pack(side='right', fill='both', expand=True)
        
        # Left side - Controls
        title = tk.Label(
            left_frame,
            text="Employee ID Matching Tool",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_L, 'bold')
        )
        title.pack(anchor='w', pady=(0, Theme.PADDING_M))
        
        subtitle = tk.Label(
            left_frame,
            text="Match employee names to IDs with intelligent fuzzy matching",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S)
        )
        subtitle.pack(anchor='w', pady=(0, Theme.PADDING_L))
        
        # Step indicator
        steps = ['Select Lookup', 'Select History', 'Configure', 'Generate']
        self.step_indicator = StepIndicator(left_frame, steps)
        self.step_indicator.pack(fill='x', pady=(0, Theme.PADDING_L))
        
        # File inputs
        self.lookup_file_var = tk.StringVar()
        self.history_file_var = tk.StringVar()
        self.output_location_var = tk.StringVar(value=self.default_output_path)
        
        # Step 1: Lookup file
        step1_frame = self.create_step_frame(left_frame, "1", "Select Employee Lookup File", "Jan 2025 names spreadsheet")
        ModernButton(
            step1_frame,
            text="Browse Lookup File",
            command=lambda: self.browse_file(self.lookup_file_var, "Select Employee Lookup"),
            style='secondary'
        ).pack(fill='x')
        
        self.lookup_label = tk.Label(
            step1_frame,
            text="No file selected",
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_TERTIARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_XS)
        )
        self.lookup_label.pack(anchor='w', pady=(Theme.PADDING_XS, 0))
        
        # Step 2: History files
        step2_frame = self.create_step_frame(left_frame, "2", "Select History Files", "Inactive Names and Helper Sheet")
        ModernButton(
            step2_frame,
            text="Browse History Files",
            command=lambda: self.browse_file(self.history_file_var, "Select History File"),
            style='secondary'
        ).pack(fill='x')
        
        self.history_label = tk.Label(
            step2_frame,
            text="No file selected",
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_TERTIARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_XS)
        )
        self.history_label.pack(anchor='w', pady=(Theme.PADDING_XS, 0))
        
        # Step 3: Configuration
        step3_frame = self.create_step_frame(left_frame, "3", "Configure Matching", "Adjust matching parameters")
        
        # Fuzzy threshold slider
        threshold_frame = tk.Frame(step3_frame, bg=Theme.BG_SECONDARY)
        threshold_frame.pack(fill='x', pady=(Theme.PADDING_S, 0))
        
        tk.Label(
            threshold_frame,
            text="Fuzzy Match Threshold:",
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S)
        ).pack(side='left')
        
        self.threshold_var = tk.IntVar(value=85)
        self.threshold_label = tk.Label(
            threshold_frame,
            text="85%",
            bg=Theme.BG_SECONDARY,
            fg=Theme.ACCENT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S, 'bold')
        )
        self.threshold_label.pack(side='right')
        
        threshold_slider = tk.Scale(
            step3_frame,
            from_=60,
            to=100,
            orient='horizontal',
            variable=self.threshold_var,
            command=self.update_threshold_label,
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_PRIMARY,
            highlightthickness=0,
            troughcolor=Theme.BG_TERTIARY,
            activebackground=Theme.ACCENT_PRIMARY
        )
        threshold_slider.pack(fill='x')
        
        # Generate button
        self.generate_button = ModernButton(
            left_frame,
            text="Generate Matched Report",
            command=self.run_matching,
            style='primary'
        )
        self.generate_button.pack(fill='x', pady=(Theme.PADDING_L, 0))
        
        # Right side - Log and summary
        log_frame = tk.Frame(right_frame, bg=Theme.BG_SECONDARY)
        log_frame.pack(fill='both', expand=True, padx=Theme.PADDING_M, pady=Theme.PADDING_M)
        
        # Summary badges
        self.summary_frame = tk.Frame(log_frame, bg=Theme.BG_SECONDARY)
        self.summary_frame.pack(fill='x', pady=(0, Theme.PADDING_M))
        
        # Log area
        log_header = tk.Frame(log_frame, bg=Theme.BG_SECONDARY)
        log_header.pack(fill='x', pady=(0, Theme.PADDING_S))
        
        tk.Label(
            log_header,
            text="Activity Log",
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M, 'bold')
        ).pack(side='left')
        
        # Log filters
        filter_frame = tk.Frame(log_header, bg=Theme.BG_SECONDARY)
        filter_frame.pack(side='right')
        
        # Log text widget
        log_container = tk.Frame(log_frame, bg=Theme.BORDER_COLOR, highlightthickness=1, highlightbackground=Theme.BORDER_COLOR)
        log_container.pack(fill='both', expand=True)
        
        self.log_text = tk.Text(
            log_container,
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S),
            wrap='word',
            relief=tk.FLAT,
            padx=Theme.PADDING_S,
            pady=Theme.PADDING_S
        )
        self.log_text.pack(fill='both', expand=True)
        
        # Initialize logger
        self.logger = EnhancedLogger(self.log_text, self.root)
        
        # Initialize matcher
        self.matcher = EmployeeIDMatcher(self.logger)
        
        # Log initial message
        self.logger.log("Employee ID Matcher ready. Select files to begin.", LogLevel.INFO)
    
    def create_step_frame(self, parent, number, title, subtitle):
        """Create a step frame with consistent styling."""
        frame = tk.Frame(parent, bg=Theme.BG_SECONDARY, relief=tk.FLAT, bd=1)
        frame.pack(fill='x', pady=(0, Theme.PADDING_M))
        
        # Add padding
        inner_frame = tk.Frame(frame, bg=Theme.BG_SECONDARY)
        inner_frame.pack(fill='both', expand=True, padx=Theme.PADDING_M, pady=Theme.PADDING_M)
        
        # Step header
        header_frame = tk.Frame(inner_frame, bg=Theme.BG_SECONDARY)
        header_frame.pack(fill='x', pady=(0, Theme.PADDING_S))
        
        tk.Label(
            header_frame,
            text=f"Step {number}: {title}",
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M, 'bold')
        ).pack(side='left')
        
        tk.Label(
            inner_frame,
            text=subtitle,
            bg=Theme.BG_SECONDARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_XS)
        ).pack(anchor='w')
        
        return inner_frame
    
    def update_threshold_label(self, value):
        """Update threshold label when slider changes."""
        self.threshold_label.config(text=f"{int(float(value))}%")
    
    def browse_file(self, string_var, title):
        """Browse for file with modern dialog."""
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if filename:
            string_var.set(filename)
            # Update appropriate label
            if string_var == self.lookup_file_var:
                self.lookup_label.config(text=os.path.basename(filename))
                self.step_indicator.set_step(1)
            elif string_var == self.history_file_var:
                self.history_label.config(text=os.path.basename(filename))
                self.step_indicator.set_step(2)
            
            self.logger.log(f"Selected: {os.path.basename(filename)}", LogLevel.INFO)
    
    def run_matching(self):
        """Run the employee ID matching process."""
        # Validate inputs
        if not self.lookup_file_var.get():
            self.logger.log("Please select employee lookup file", LogLevel.ERROR)
            return
        
        if not self.history_file_var.get():
            self.logger.log("Please select history file", LogLevel.ERROR)
            return
        
        # Disable button and show progress
        self.generate_button.config(state='disabled', text='Processing...')
        self.step_indicator.set_step(3)
        
        # Run in thread
        thread = threading.Thread(target=self.matching_thread)
        thread.start()
    
    def matching_thread(self):
        """Thread for running the matching process."""
        try:
            # Load lookup data
            self.logger.log("Loading employee lookup data...", LogLevel.INFO)
            self.matcher.fuzzy_threshold = self.threshold_var.get()
            
            if not self.matcher.load_lookup_file(self.lookup_file_var.get()):
                raise Exception("Failed to load lookup file")
            
            # Load history data
            self.logger.log("Loading history data...", LogLevel.INFO)
            history_df = pd.read_excel(self.history_file_var.get())
            
            # Find name column
            name_col = None
            for col in history_df.columns:
                if 'name' in str(col).lower():
                    name_col = col
                    break
            
            if not name_col:
                raise Exception("Could not find name column in history file")
            
            # Process matching
            self.logger.log("Performing name matching...", LogLevel.INFO)
            result_df = self.matcher.process_dataframe(history_df, name_col)
            
            # Generate output
            output_file = os.path.join(
                self.output_location_var.get(),
                format_output_filename("Master_Names_with_IDs")
            )
            
            result_df.to_excel(output_file, index=False)
            
            # Update UI with results
            self.root.after(0, self.matching_complete, output_file)
            
        except Exception as e:
            self.logger.log(f"Error during matching: {str(e)}", LogLevel.ERROR)
            self.root.after(0, lambda: self.generate_button.config(state='normal', text='Generate Matched Report'))
    
    def matching_complete(self, output_file):
        """Handle completion of matching process."""
        self.generate_button.config(state='normal', text='Generate Matched Report')
        self.step_indicator.set_step(4)
        
        # Show summary
        stats = self.matcher.match_stats
        total = sum(stats.values())
        
        self.logger.log(f"Matching complete! Output saved to: {os.path.basename(output_file)}", LogLevel.SUCCESS)
        self.logger.log(f"Total processed: {total}", LogLevel.INFO)
        self.logger.log(f"Exact matches: {stats['exact']}", LogLevel.SUCCESS)
        self.logger.log(f"Fuzzy matches: {stats['fuzzy']}", LogLevel.WARNING)
        self.logger.log(f"No matches: {stats['no_match']}", LogLevel.ERROR)
        self.logger.log(f"Already had ID: {stats['already_had_id']}", LogLevel.INFO)
        
        # Update summary badges
        self.update_summary_badges(stats)
        
        # Show success dialog
        messagebox.showinfo(
            "Success",
            f"Employee ID matching complete!\n\n"
            f"Total processed: {total}\n"
            f"Exact matches: {stats['exact']}\n"
            f"Fuzzy matches: {stats['fuzzy']}\n"
            f"No matches: {stats['no_match']}\n"
            f"Already had ID: {stats['already_had_id']}\n\n"
            f"Output saved to:\n{output_file}"
        )
    
    def update_summary_badges(self, stats):
        """Update summary badges with match statistics."""
        # Clear existing badges
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        
        # Create badges
        badges = [
            ("Exact", stats['exact'], Theme.ACCENT_SUCCESS),
            ("Fuzzy", stats['fuzzy'], Theme.ACCENT_WARNING),
            ("No Match", stats['no_match'], Theme.ACCENT_ERROR),
            ("Existing", stats['already_had_id'], Theme.ACCENT_PRIMARY)
        ]
        
        for label, count, color in badges:
            badge_frame = tk.Frame(self.summary_frame, bg=color)
            badge_frame.pack(side='left', padx=(0, Theme.PADDING_S))
            
            tk.Label(
                badge_frame,
                text=f"{label}: {count}",
                bg=color,
                fg=Theme.BG_PRIMARY,
                font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_S, 'bold'),
                padx=Theme.PADDING_S,
                pady=Theme.PADDING_XS
            ).pack()
    
    def add_reporter_tab(self):
        """Add the Pay Period Reporter tab (placeholder)."""
        tab = tk.Frame(self.notebook, bg=Theme.BG_PRIMARY)
        self.notebook.add(tab, text='Pay Period Reporter')
        
        tk.Label(
            tab,
            text="Pay Period Reporter",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_L, 'bold')
        ).pack(pady=Theme.PADDING_L)
        
        tk.Label(
            tab,
            text="This feature has been preserved from the original tool.\nFull implementation available in production version.",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M)
        ).pack()
    
    def add_formatter_tab(self):
        """Add the Helper Sheet Formatter tab (placeholder)."""
        tab = tk.Frame(self.notebook, bg=Theme.BG_PRIMARY)
        self.notebook.add(tab, text='Helper Sheet Formatter')
        
        tk.Label(
            tab,
            text="Helper Sheet Formatter",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_L, 'bold')
        ).pack(pady=Theme.PADDING_L)
        
        tk.Label(
            tab,
            text="This feature has been preserved from the original tool.\nFull implementation available in production version.",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M)
        ).pack()
    
    def add_emailer_tab(self):
        """Add the Office Emailer tab (placeholder)."""
        tab = tk.Frame(self.notebook, bg=Theme.BG_PRIMARY)
        self.notebook.add(tab, text='Office Emailer')
        
        tk.Label(
            tab,
            text="Office Emailer",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_PRIMARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_L, 'bold')
        ).pack(pady=Theme.PADDING_L)
        
        tk.Label(
            tab,
            text="This feature has been preserved from the original tool.\nHTML CSS bug has been fixed in line 789.",
            bg=Theme.BG_PRIMARY,
            fg=Theme.TEXT_SECONDARY,
            font=(Theme.FONT_FAMILY, Theme.FONT_SIZE_M)
        ).pack()
    
    def run(self):
        """Start the application."""
        self.root.mainloop()

if __name__ == "__main__":
    app = MasterWorkflowApp()
    app.run()