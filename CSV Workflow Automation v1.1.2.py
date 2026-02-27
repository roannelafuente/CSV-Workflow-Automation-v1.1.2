import tkinter as tk
from tkinter import ttk, filedialog
import os, csv, xlwings as xw
from datetime import datetime, timedelta
import logging
from contextlib import contextmanager

# CSV Workflow Automation Tool v1.1.2
# Author: Rose Anne Lafuente
# Licensed Electronics Engineer | Product Engineer II | Python Automation
#
# Description:
#   Automates CSV-to-Excel workflows with pivot tables, custom formatting,
#   End Test validation, and wafermap visualization for yield and defect tracking.
#
#   Key Features:
#     - Scrollable status box for enhanced log navigation
#     - Deterministic wafermap coloring via defined C1_MARK color_map
#     - Accurate C1_MARK lookup for ET mapping
#     - Robust error handling with centralized ErrorLogger:
#         • Auto-creates timestamped error log files in a dedicated /logs folder
#         • Cleans up logs older than 30 days automatically
#         • Provides consistent error capture across all modules
#     - Context-managed Excel operations for speed and reliability
#     - GUI title, version label, and developer credit for professional branding
#
#   Built with:
#     - Python 3.x
#     - Tkinter (GUI framework)
#     - OpenPyXL (Excel file manipulation)
#     - xlwings (Excel COM automation)
#
# Version Highlights (v1.1.2):
#     - Refactored into modular, multi-class design
#     - Helper functions for normalization and header lookup
#     - Optimized bulk operations for speed
#     - Consistent docstrings for maintainability
#     - Cleaner, portfolio-ready architecture
#     - Integrated ErrorLogger for centralized error tracking and log retention


# --- Error Logger with 30-day cleanup ---
class ErrorLogger:
    def __init__(self, days_to_keep=30):
        self.log_dir = os.path.join(os.path.dirname(__file__), "logs")
        self.days_to_keep = days_to_keep
        self.log_file = None
        self.is_configured = False

    def setup_on_error(self):
        """Create logs folder and configure logging only when an error occurs."""
        if not self.is_configured:
            os.makedirs(self.log_dir, exist_ok=True)
            self.log_file = os.path.join(
                self.log_dir,
                f"error_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            logging.basicConfig(
                filename=self.log_file,
                level=logging.ERROR,
                format="%(asctime)s - %(levelname)s - %(message)s"
            )
            self.is_configured = True
        return self.log_file

    def log_error(self, msg: str):
        """Log an error message, ensuring setup is done first."""
        self.setup_on_error()
        logging.error(msg)

    def cleanup_old_logs(self):
        """Delete log files older than the retention period."""
        if not os.path.exists(self.log_dir):
            return  # Nothing to clean if folder never created
        cutoff = datetime.now() - timedelta(days=self.days_to_keep)
        for fname in os.listdir(self.log_dir):
            fpath = os.path.join(self.log_dir, fname)
            if os.path.isfile(fpath) and fname.startswith("csv_workflow_automation_error_log_"):
                try:
                    timestamp_str = fname.replace("csv_workflow_automation_error_log_", "").replace(".txt", "")
                    dt = datetime.strptime(timestamp_str, "%Y%m%d_%H%M%S")
                    if dt < cutoff:
                        os.remove(fpath)
                        # print(f"🗑 Deleted old log file: {fname}")
                except Exception:
                    continue
# --- Utility Functions ---
def normalize_value(val):
    if val is None:
        return None
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

@contextmanager
def open_workbook(path, visible=False):
    app = xw.App(visible=visible)
    wb = app.books.open(path)
    try:
        yield wb
    finally:
        wb.save()
        wb.close()
        app.quit()
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- GUI Class ---
class CSVWorkflowAutomationGUI:
    """
    Main GUI class for the CSV Workflow Automation tool.
    Handles user interface, workflow execution, and error logging.
    """
    def __init__(self, root):
        """
        Initialize the GUI, set theme, and build all frames.

        - Sets window title, geometry, and background theme.
        - Loads custom sprout.ico icon for branding.
          Falls back gracefully if the icon file is missing.
        - Builds title frame, developer credit, and main interface sections.
        """
        self.root = root
        self.root.title("CSV Workflow Automation")
        self.root.geometry("750x550")

        # ✅ Load custom sprout icon for branding
        try:
            self.root.iconbitmap("sprout.ico")
        except Exception:
            # Fallback if icon not found (prevents crash)
            pass

        # Professional Neutral Theme
        self.bg_color = "#f5f5f5"
        self.fg_color = "#222222"
        self.btn_bg = "#e0e0e0"
        self.btn_active = "#BEE395"
        self.root.configure(bg=self.bg_color)

        # Initialize error logger
        self.logger = ErrorLogger()

        # --- Title Frame ---
        title_frame = tk.Frame(self.root, bg=self.bg_color)
        title_frame.pack(pady=(10,0))

        title_label = tk.Label(
            title_frame,
            text="CSV Workflow Automation",
            font=("Meiryo", 12, "bold"),
            fg="darkblue",
            bg=self.bg_color
        )
        title_label.pack(side="left")

        version_label = tk.Label(
            title_frame,
            text=" v1.1.2",
            font=("Meiryo", 12, "italic"),
            fg="darkblue",
            bg=self.bg_color
        )
        version_label.pack(side="left")

        dev_label = tk.Label(
            self.root,
            text="Developed by Rose Anne Lafuente | 2026",
            font=("Arial", 7, "italic"),
            fg="gray",
            bg=self.bg_color
        )
        dev_label.pack(pady=(0,10))

        # Initialize variables and build interface
        self.path_var = tk.StringVar()
        self.create_file_selection_frame()
        self.create_filter_selector([])
        self.create_status_box()
        self.create_exit_button()


    # --- Helper Methods ---
    def find_header_row(self, sheet, col_letter, header_name):
        """
        Scan a given column for a header name and return its row index.
        Returns None if not found.
        """
        col_values = sheet.range(f"{col_letter}1:{col_letter}{sheet.cells.last_cell.row}").value
        for i, val in enumerate(col_values, start=1):
            if str(val).strip().upper() == header_name.upper():
                return i
        return None

    # --- GUI Builders ---
    def create_file_selection_frame(self):
        """
        Build the file selection frame with:
        - Label for CSV file
        - Entry box for file path
        - Browse and Convert buttons
        """
        frame = tk.LabelFrame(self.root, text="File Selection", padx=10, pady=10,
                              bd=2, relief="groove", font=("Segoe UI", 10, "bold"))
        frame.pack(fill="x", padx=15, pady=10)

        tk.Label(frame, text="Select CSV File:").pack(side="left", padx=(0,10), pady=5)

        path_entry = tk.Entry(frame, textvariable=self.path_var, width=60,
                              bg="white", fg="black", insertbackground="black")
        path_entry.pack(side="left", padx=10, pady=5, fill="x", expand=True)
        tk.Button(frame, text="Convert to Excel", width=18, command=self.convert_to_excel,
                  bg=self.btn_bg, fg=self.fg_color, activebackground=self.btn_active).pack(side="right", padx=10, pady=5)
        tk.Button(frame, text="Browse", width=12, command=self.browse_file).pack(side="right", pady=5)
        
    def create_filter_selector(self, items):
        # Pivot Filter Selection frame with subtle border and spacing
        filter_frame = tk.LabelFrame(
            self.root,
            text="Pivot Filter Selection",
            padx=10, pady=10,
            bd=2,
            relief="groove",
            font=("Segoe UI", 10, "bold")
        )

        filter_frame.pack(fill="x", padx=15, pady=10)

        tk.Label(
            filter_frame,
            text="Select C1_MARK:"
        ).pack(side="left", padx=5, expand=True, fill="x")

        # ✅ Clean and deduplicate items
        clean_items = [str(i) for i in items if i is not None]
        unique_items = list(dict.fromkeys(clean_items))

        self.filter_var = tk.StringVar()
        self.filter_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.filter_var,
            values=unique_items,
            state="readonly",
            width=5
        )
        self.filter_dropdown.pack(side="left", padx=10)

        gen_pivot_btn = tk.Button(
            filter_frame,
            text="Generate Pivot Table",
            width=18,
            command=self.generate_pivot,
            bg="#8BD3E6",
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        gen_pivot_btn.pack(side="left", padx=10, expand=True, fill="x")


        check_test_btn = tk.Button(
            filter_frame,
            text="Check End Test No",
            width=18,
            command=self.check_end_test,
            bg="#E6E6FA",
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        check_test_btn.pack(side="left", padx=10, expand=True, fill="x")


        gen_wafermap_btn = tk.Button(
            filter_frame,
            text="Generate Wafermap",
            width=18,
            command=self.generate_wafermap,
            bg="#92D050",
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        gen_wafermap_btn.pack(side="left", padx=10, expand=True, fill="x")
        
    def get_unique_c1_mark_values(self, raw_items):
        """
        Clean and deduplicate C1_MARK values.
        Flattens nested lists, normalizes values, and removes duplicates.
        """
        def normalize(i):
            if i is None:
                return None
            if isinstance(i, float) and i.is_integer():
                return str(int(i))   # 1.0 → "1"
            return str(i).strip()

        flat = []
        for item in raw_items:
            if isinstance(item, list):
                flat.extend(item)
            elif item is not None:
                flat.append(item)

        cleaned = [normalize(i) for i in flat if i is not None]
        unique = list(dict.fromkeys(cleaned))
        return unique
    
    def create_status_box(self):
        """
        Build the scrollable status box for logging messages.
        Includes vertical and horizontal scrollbars.
        """
        frame = tk.LabelFrame(self.root, text="Status", padx=10, pady=10)
        frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.status_box = tk.Text(frame, height=10, wrap="word", bg="white", fg="black", state="disabled")
        vsb = tk.Scrollbar(frame, orient="vertical", command=self.status_box.yview)
        hsb = tk.Scrollbar(frame, orient="horizontal", command=self.status_box.xview)
        self.status_box.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.status_box.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

    def create_exit_button(self):
        """
        Build the bottom frame with:
        - Exit button
        - Clear All button
        """
        frame = tk.Frame(self.root, bg=self.bg_color)
        frame.pack(fill="x", side="bottom", padx=15, pady=5)
        tk.Button(frame, text="EXIT", width=12, bg="#d32f2f", fg="white",
                  command=self.root.destroy).pack(side="right", pady=10)
        tk.Button(frame, text="Clear All", width=12, command=self.clear_all,
                  bg="#ffcccc", fg=self.fg_color, activebackground=self.btn_active).pack(side="right", padx=10)

    # --- Status Logging ---
    def show_status(self, message, color="#000000", clear=False):
        """
        Display a message in the status box.
        Supports optional color coding and clearing previous logs.
        """
        self.status_box.config(state="normal")
        if clear:
            self.status_box.delete("1.0", "end")
        if message:
            self.status_box.insert("end", message + "\n")
            line_tag = f"status_{self.status_box.index('end-2l')}"
            self.status_box.tag_add(line_tag, self.status_box.index("end-2l"), self.status_box.index("end-1c"))
            self.status_box.tag_config(line_tag, foreground=color)
        self.status_box.config(state="disabled")

    # --- File Handling ---
    def browse_file(self):
        """
        Open file dialog to select a CSV file.
        Updates the path entry and logs the selection.
        """
        file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.path_var.set(file_path)
            self.show_status(f"📂 Selected file: {file_path}")

    def convert_to_excel(self):
        """
        Convert selected CSV file to Excel (.xlsx).
        Normalizes ragged rows, writes in bulk, and saves output.
        """
        file_path = self.path_var.get()
        if not file_path:
            self.show_status("⚠️ No file selected.", color="#d32f2f")
            return
        try:
            sheet_name = os.path.splitext(os.path.basename(file_path))[0]
            out_file = os.path.splitext(file_path)[0] + ".xlsx"
            app = xw.App(visible=False)
            wb = app.books.add()
            sht = wb.sheets[0]
            sht.name = sheet_name[:31]

            with open(file_path, newline='', encoding='utf-8') as f:
                reader = list(csv.reader(f))
            max_len = max(len(row) for row in reader)
            normalized = [row + [""] * (max_len - len(row)) for row in reader]
            sht.range("A1").value = normalized

            wb.save(out_file)
            wb.close()
            app.quit()

            self.out_file = out_file
            self.sheet_name = sheet_name
            self.extract_filter_items()
            self.show_status(f"\n✅ Conversion complete: {out_file}")
        except Exception as e:
            # Ensure logger is configured
            self.logger.setup_on_error()

            # Log the error
            logging.critical(f"Unexpected error: {e}", exc_info=True)

            # Show status in GUI
            self.show_status(f"❌ Unexpected error: {e}", color="#d32f2f")



    # --- Placeholder Methods (to be filled with your existing logic) ---
    def extract_filter_items(self):
        """
        Extract unique C1_MARK values from the Excel sheet.
        Populates the combobox with deduplicated filter items.
        """
        try:
            with open_workbook(self.out_file) as wb_xlw:
                sht = wb_xlw.sheets[0]
                header_row = self.find_header_row(sht, "G", "C1_MARK")
                if not header_row:
                    self.show_status("❌ 'C1_MARK' not found in Column G.", color="#d32f2f")
                    return

                last_row = sht.range((header_row, 7)).end("down").row
                raw_items = sht.range((header_row+1, 7), (last_row, 7)).value
                self.raw_items = raw_items

                unique_items = self.get_unique_c1_mark_values(self.raw_items)
                self.filter_dropdown['values'] = unique_items
                self.base_name = self.get_unique_c1_mark_values([self.sheet_name])[0]

                #self.show_status("✅ Combobox populated with filter items.")
        except Exception as e:
            # Ensure logger is configured
            self.logger.setup_on_error()

            # Log the error
            logging.critical(f"Unexpected error: {e}", exc_info=True)

            # Show status in GUI
            self.show_status(f"❌ Unexpected error: {e}", color="#d32f2f")
    
    def generate_pivot(self):
        """
        Generate a pivot table filtered by selected C1_MARK.
        Builds fallout table with counts and percentages.
        Applies formatting and previews results in status box.
        """

        selected = self.filter_var.get()
        if not selected:
            self.show_status("⚠️ Please select a C1_MARK value first.", color="#d32f2f")
            return

        self.show_status(f"\nℹ️ Generating pivot table...")

        try:
            with open_workbook(self.out_file) as wb_xlw:
                sht = wb_xlw.sheets[self.base_name]

                # --- Find C1_MARK header row ---
                header_row = self.find_header_row(sht, "G", "C1_MARK")
                if not header_row:
                    self.show_status("❌ 'C1_MARK' not found in Column G.", color="#d32f2f")
                    return

                # --- Locate ET column ---
                row_values = sht.range((header_row, 7), (header_row, sht.range((header_row, 7)).end("right").column)).value
                et_col = None
                for idx, val in enumerate(row_values, start=7):
                    if str(val).strip().upper() == "ET":
                        et_col = idx
                        break
                if not et_col:
                    raise ValueError("'ET' column not found to the right of C1_MARK")

                # --- Define pivot source range ---
                last_row = sht.range((header_row, 7)).end("down").row
                pivot_range = sht.range((header_row, 7), (last_row, et_col))

                # --- Create Pivot sheet ---
                try:
                    pivot_sheet = wb_xlw.sheets["Pivot"]
                    pivot_sheet.clear()
                except:
                    pivot_sheet = wb_xlw.sheets.add("Pivot", after=sht)

                # --- Create pivot cache and table ---
                pivot_cache = wb_xlw.api.PivotCaches().Create(SourceType=1, SourceData=pivot_range.api)
                table_name = f"PivotTable_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                pivot_table = pivot_cache.CreatePivotTable(TableDestination=pivot_sheet.range("A3").api, TableName=table_name)

                # --- Filter: C1_MARK ---
                pf = pivot_table.PivotFields("C1_MARK")
                pf.Orientation = 3
                valid_items = [item.Name for item in pf.PivotItems()]
                if selected in valid_items:
                    pf.CurrentPage = selected
                    self.show_status(f"\nApplied filter: {selected}")
                else:
                    self.show_status(f"⚠️ Selected '{selected}' not found in C1_MARK items {valid_items}", color="#d32f2f")
                    return

                # --- Rows: ET ---
                pivot_table.PivotFields("ET").Orientation = 1

                # --- Values: Count of FT ---
                pivot_table.AddDataField(pivot_table.PivotFields("FT"), "Count of FT", -4112)

                # --- Fallout Table Logic ---
                data = pivot_sheet.range("A4").expand().value
                theoretical_num = None
                for i, val in enumerate(sht.range("A:A").value, start=1):
                    if str(val).strip().upper() == "THEORETICAL_NUM":
                        theoretical_num = sht.range((i, 1)).offset(0, 2).value
                        break

                fallout_table = []
                for row in data:
                    if not row or not row[0] or str(row[0]).strip().lower() == "grand total":
                        continue
                    et_val = normalize_value(row[0])
                    count_val = int(row[1]) if isinstance(row[1], (int, float)) and float(row[1]).is_integer() else row[1]
                    fallout = (float(row[1]) / theoretical_num * 100) if theoretical_num else 0
                    fallout_table.append([et_val, count_val, f"{fallout:.2f}%"])

                fallout_table.sort(key=lambda x: int(x[1]), reverse=True)
                grand_total_val = normalize_value(theoretical_num)
                fallout_table.insert(0, ["End Test No.", "Count", "Fallout%"])  # header row
                fallout_table.append(["Grand Total", grand_total_val, ""])

                # --- Vectorized write fallout table ---
                pivot_sheet.range("D3").value = fallout_table

                # --- Apply formatting ---
                last_row_ft = 3 + len(fallout_table) - 1
                fallout_range = pivot_sheet.range(f"D3:F{last_row_ft}")
                fallout_range.api.HorizontalAlignment = -4108
                fallout_range.api.VerticalAlignment = -4108
                fallout_range.api.IndentLevel = 0

                # Header row
                pivot_sheet.range("D3:F3").color = (192, 230, 245)
                pivot_sheet.range("D3:F3").api.Font.Bold = True
                # First data row
                pivot_sheet.range("D4:F4").color = (255, 159, 159)
                pivot_sheet.range("D4:F4").api.Font.Bold = True
                # Grand Total row
                pivot_sheet.range(f"D{last_row_ft}:F{last_row_ft}").color = (192, 230, 245)
                pivot_sheet.range(f"D{last_row_ft}:F{last_row_ft}").api.Font.Bold = True

                fallout_range.api.Borders.Weight = 2
                wb_xlw.save()

                # --- Show fallout table in status box ---
                self.status_box.config(state="normal")
                self.status_box.insert(tk.END, "\nPreview Table:\n")
                for et_val, count_val, fallout_val in fallout_table:
                    self.status_box.insert(tk.END, f"{str(et_val):<15}{str(count_val):<10}{str(fallout_val)}\n")
                self.status_box.config(state="disabled")

                self.show_status(f"\n✅ Successfully generated table for C1_MARK: {selected}")

        except Exception as e:
            # Ensure logger is configured
            self.logger.setup_on_error()

            # Log the error
            logging.critical(f"Unexpected error: {e}", exc_info=True)

            # Show status in GUI
            self.show_status(f"❌ Unexpected error: {e}", color="#d32f2f")
            
    def check_end_test(self):
        """
        Check End Test No. against reference table.
        Displays limits and highlights results in status box.
        """
        try:
            with open_workbook(self.out_file) as wb_xlw:
                # --- Ensure Pivot sheet exists ---
                try:
                    pivot_sheet = wb_xlw.sheets["Pivot"]
                except:
                    pivot_sheet = wb_xlw.sheets.add("Pivot")

                data_sheet = wb_xlw.sheets[self.base_name]

                # --- Get highest fails End Test No from D4 ---
                raw_val = pivot_sheet.range("D4").value
                if raw_val is None:
                    end_test_no = ""
                else:
                    end_test_no = normalize_value(raw_val)

                self.show_status(f"\n🔍 Checking End Test No.: {end_test_no}")

                # --- Find LOLIMIT row in Column F ---
                lolimit_row = data_sheet.range("F1").end("down").row
                lolimit_val = data_sheet.range(f"F{lolimit_row}").value
                if str(lolimit_val).strip().upper() != "LOLIMIT":
                    raise ValueError("LOLIMIT not found in Column F")

                # --- Expand reference table ---
                ref_range = data_sheet.range((lolimit_row, 1)).expand("table")

                # --- Locate TESTNO column (Column B) ---
                testno_values = data_sheet.range(
                    (lolimit_row + 1, 2),
                    (lolimit_row + ref_range.rows.count - 1, 2)
                ).value

                # Normalize TESTNO values
                testno_values = [normalize_value(v) if v else "" for v in testno_values]

                found_row = None
                if end_test_no in testno_values:
                    idx = testno_values.index(end_test_no) + lolimit_row + 1
                    found_row = idx

                # --- Vectorized write of header + data ---
                start_cell = pivot_sheet.range("H3")
                header = ["TSNO", "TESTNO", "COMMENT", "MODE", "HILIMIT", "LOLIMIT"]

                if found_row:
                    row_values = data_sheet.range((found_row, 1), (found_row, 6)).value
                    row_values = [normalize_value(v) if v else "" for v in row_values]

                    pivot_sheet.range("H3").value = [header, row_values]

                    # --- Apply formatting ---
                    ref_range_excel = pivot_sheet.range("H3:M4")
                    header_range = pivot_sheet.range("H3:M3")
                    data_range = pivot_sheet.range("H4:M4")

                    header_range.color = (192, 230, 245)   # light blue
                    header_range.api.Font.Bold = True
                    data_range.color = (255, 255, 255)     # white
                    data_range.api.Font.Bold = True

                    ref_range_excel.api.Borders.Weight = 2
                    ref_range_excel.api.HorizontalAlignment = -4108
                    ref_range_excel.api.VerticalAlignment = -4108
                    ref_range_excel.api.IndentLevel = 0

                    wb_xlw.save()

                    # --- Show End Test No. table in status box ---
                    self.status_box.config(state="normal")
                    self.status_box.insert(tk.END, "\nEnd Test No. Reference:\n")
                    self.status_box.insert(
                        tk.END,
                        f"{'TSNO':<10}{'TESTNO':<10}{'COMMENT':<15}{'MODE':<10}{'HILIMIT':<10}{'LOLIMIT'}\n"
                    )
                    self.status_box.insert(tk.END, "-" * 70 + "\n")
                    tsno, testno, comment, mode, hilimit, lolimit = row_values
                    self.status_box.insert(
                        tk.END,
                        f"{tsno:<10}{testno:<10}{comment:<15}{mode:<10}{hilimit:<10}{lolimit}\n"
                    )
                    self.status_box.config(state="disabled")

                    # --- Status message depending on limits ---
                    if lolimit != "":
                        self.show_status("\n✅ Found with Limits")
                    else:
                        self.show_status("\n⚠️ Found with no Limit", color="#FFBF00")
                else:
                    self.show_status("\n❌ No End Test No. found in the TESTNO Column", color="#d32f2f")

        except Exception as e:
            # Ensure logger is configured
            self.logger.setup_on_error()

            # Log the error
            logging.critical(f"Unexpected error: {e}", exc_info=True)

            # Show status in GUI
            self.show_status(f"❌ Unexpected error: {e}", color="#d32f2f")
            
    def build_et_to_c1_map(self, sheet, header_row, et_col):
        """
        Build a dictionary mapping End Test (ET) values to C1_MARK values.
        Normalizes both ET and C1_MARK for consistent lookup.
        """
        et_to_c1 = {}
        last_row = sheet.range((header_row+1, et_col)).end("down").row
        for row in range(header_row+1, last_row+1):
            et_val = sheet.range((row, et_col)).value
            c1_val = sheet.range((row, 7)).value  # Column G = C1_MARK
            if et_val is None or c1_val is None:
                continue

            # Normalize ET
            if isinstance(et_val, float) and et_val.is_integer():
                et_str = str(int(et_val))
            else:
                et_str = str(et_val).strip()

            # Normalize C1_MARK
            if isinstance(c1_val, float) and c1_val.is_integer():
                c1_str = str(int(c1_val))
            else:
                c1_str = str(c1_val).strip()

            et_to_c1[et_str] = c1_str
        return et_to_c1
    
    def generate_wafermap(self):
        """
        Create a wafermap sheet by building an ET→C1_MARK mapping,
        generating a pivot table with X/Y coordinates, and applying
        deterministic cell coloring based on the predefined color_map.
        """

        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)
            data_sheet = wb_xlw.sheets[self.base_name]

            # --- SLOT handling ---
            slot_row = self.find_header_row(data_sheet, "A", "SLOT")
            if not slot_row:
                self.show_status("\n⚠️ SLOT header not found in Column A", color="#d32f2f")
                return

            slot_val = data_sheet.range((slot_row+1, 1)).value
            if slot_val is None:
                self.show_status("\n⚠️ SLOT value below header is empty", color="#d32f2f")
                return

            slot_str = str(int(slot_val)).zfill(2)
            self.show_status(f"\n🔍 Generating wafermap for W #{slot_str}...")
            sheet_name = f"W#{slot_str}_Wafermap_by_End_Test_No"

            # --- Create or reuse Wafermap Pivot Table sheet ---
            try:
                pivot_sheet = wb_xlw.sheets["Wafermap Pivot Table"]
                pivot_sheet.clear()
            except:
                pivot_sheet = wb_xlw.sheets.add("Wafermap Pivot Table", after=data_sheet)

            # --- Create or reuse slot-specific wafermap sheet ---
            try:
                wafermap_sheet = wb_xlw.sheets[sheet_name]
                wafermap_sheet.clear()
            except:
                wafermap_sheet = wb_xlw.sheets.add(sheet_name, after=pivot_sheet)

            # --- Find header row with C1_MARK ---
            header_row = self.find_header_row(data_sheet, "G", "C1_MARK")
            if not header_row:
                self.show_status("❌ 'C1_MARK' not found in Column G.", color="#d32f2f")
                return

            # --- Locate X, Y, ET columns ---
            row_values = data_sheet.range(
                (header_row, 1),
                (header_row, data_sheet.range((header_row, 1)).end("right").column)
            ).value

            x_col = y_col = et_col = None
            for idx, val in enumerate(row_values, start=1):
                if str(val).strip().upper() == "X":
                    x_col = idx
                elif str(val).strip().upper() == "Y":
                    y_col = idx
                elif str(val).strip().upper() in ["ET", "END TEST NO."]:
                    et_col = idx

            if not (x_col and y_col and et_col):
                raise ValueError("Required columns 'X', 'Y', 'ET' not found in header row")

            # --- Build ET → C1_MARK mapping using helper ---
            et_to_c1 = self.build_et_to_c1_map(data_sheet, header_row, et_col)

            # --- Create pivot cache and table ---
            last_row = data_sheet.range((header_row+1, et_col)).end("down").row
            pivot_range = data_sheet.range((header_row, x_col), (last_row, et_col))
            pivot_cache = wb_xlw.api.PivotCaches().Create(SourceType=1, SourceData=pivot_range.api)
            table_name = f"PivotTable_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=pivot_sheet.range("A1").api,
                TableName=table_name
            )

            # --- Configure pivot ---
            pivot_table.PivotFields("Y").Orientation = 1
            pivot_table.PivotFields("X").Orientation = 2
            pivot_table.AddDataField(pivot_table.PivotFields("ET"), "Min of ET", -4139)
            pivot_table.ColumnGrand = False
            pivot_table.RowGrand = False
            pivot_sheet.range("A2").value = "No."

            # --- Copy pivot output ---
            pivot_block = pivot_sheet.range("A2").expand()
            data_block = pivot_block.value

            # --- Paste values into wafermap sheet ---
            rows = len(data_block)
            cols = len(data_block[0])
            wafermap_sheet.range((1,1), (rows,cols)).value = data_block

            # --- Alignment ---
            wafermap_sheet.range((1,1), (rows,cols)).api.HorizontalAlignment = -4108
            wafermap_sheet.range((1,1), (rows,cols)).api.VerticalAlignment = -4108

            # --- Find last used row/col ---
            last_col = wafermap_sheet.range("1:1").end("right").column
            last_row = wafermap_sheet.range("A:A").end("down").row

            # --- Header formatting ---
            dark_blue = xw.utils.rgb_to_int((46, 110, 158))
            wafermap_sheet.range((1,1),(1,last_col)).color = (228, 241, 253)
            wafermap_sheet.range((1,1),(1,last_col)).api.Font.Color = dark_blue
            wafermap_sheet.range((1,1),(last_row,1)).color = (228, 241, 253)
            wafermap_sheet.range((1,1),(last_row,1)).api.Font.Color = dark_blue

            # --- Define C1_MARK color mapping dictionary ---
            color_map = {
                "/":"#00FF00", "$":"#7B68EE", "*":"#87CEEB", "?":"#66FF66", "=":"#7FFFD4", "!":"#6495ED", "#":"#6A5ACD",
                "%":"#66FF66", ".":"#66FF66", ":":"#66FF66", "^":"#66FF66", "+":"#66FF66", "-":"#66FF66", "{":"#66FF66",
                "}":"#66FF66", "(":"#66FF66", ")":"#66FF66", "_":"#66FF66", "|":"#66FF66", ";":"#66FF66", "@":"#66FF66",
                "\\":"#66FF66", "<":"#66FF66", ">":"#66FF66", "&":"#66FF66",
                "0":"#66FF66", "1":"#FFFF99", "2":"#FF0000", "3":"#FFFFE0", "4":"#ADD8E6", "5":"#FF8080", "6":"#AFEEEE",
                "7":"#99CCFF", "8":"#FFCC00", "9":"#FFFF00",
                "A":"#2E8B57", "B":"#FFCC00", "C":"#FFCC00", "D":"#99CC00", "E":"#99CC00", "F":"#7CFC00", "G":"#FFFF00",
                "H":"#A6A6A6", "I":"#00CCFF", "J":"#32CD32", "K":"#20B2AA", "L":"#FFDEAD", "M":"#D9D9D9", "N":"#DAA520",
                "O":"#00CCFF", "P":"#FFFF99", "Q":"#ED7D31", "R":"#FFCC00", "S":"#FF7C80", "T":"#FFCC00", "U":"#00CCFF",
                "V":"#008080", "W":"#008080", "X":"#008080", "Y":"#666699", "Z":"#666699",
                "a":"#D2691E", "b":"#993366", "c":"#A52A2A", "d":"#E9967A", "e":"#660066", "f":"#ED7D31", "g":"#3366FF",
                "h":"#CCFFFF", "i":"#FF7F50", "j":"#99CCFF", "k":"#CCCCFF", "l":"#D9D9D9", "m":"#969696", "n":"#339966",
                "o":"#333399", "p":"#FF6600", "q":"#FFFF00", "r":"#0066CC", "s":"#FF9900", "t":"#33CCCC", "u":"#008080",
                "v":"#EE82EE", "w":"#DDA0DD", "x":"#00FFFF", "y":"#99CC00", "z":"#9932CC"
            }

            # --- Apply colors to wafermap cells using ET → C1_MARK mapping ---
            for r in range(2, last_row+1):
                for c in range(2, last_col+1):
                    cell = wafermap_sheet.range((r,c))
                    et_val = cell.value
                    if et_val is None or str(et_val).strip() == "":
                        continue

                    # Normalize ET consistently
                    if isinstance(et_val, float) and et_val.is_integer():
                        et_str = str(int(et_val))
                    else:
                        et_str = str(et_val).strip()

                    # Lookup C1_MARK from dictionary
                    c1_mark_str = et_to_c1.get(et_str)

                    if c1_mark_str:
                        if c1_mark_str in color_map:
                            hex_color = color_map[c1_mark_str].lstrip("#")
                            rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                            cell.color = rgb
                        else:
                            self.show_status(f"⚠️ No color mapping for C1_MARK '{c1_mark_str}'", color="#d32f2f")
                            cell.color = (200,200,200)
                    else:
                        self.show_status(f"⚠️ No C1_MARK found for ET '{et_str}'", color="#d32f2f")
                        cell.color = (200,200,200)
                                
            # --- Copy Row 1 (Ctrl+Shift+Right) and paste it after last used row ---
            row1_vals = wafermap_sheet.range((1,1),(1,last_col)).value
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).value = row1_vals
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).color = (228,241,253)
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).api.Font.Color = dark_blue
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).api.Font.Bold = True  # bold copy of Row 1

            # --- Copy Column A (Ctrl+Shift+Down) and paste it after last used column ---
            colA_vals = wafermap_sheet.range((1,1),(last_row,1)).value

            # Ensure values are shaped as a column (list of lists)
            if isinstance(colA_vals, list) and not isinstance(colA_vals[0], list):
                colA_vals = [[v] for v in colA_vals]

            # Paste Column A into the new rightmost column
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).value = colA_vals
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).color = (228,241,253)
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).api.Font.Color = dark_blue
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).api.Font.Bold = True  # bold copy of Column A

            # --- Add "No." at the very last row of that new column ---
            wafermap_sheet.range((last_row+1, last_col+1)).value = "No."
            wafermap_sheet.range((last_row+1, last_col+1)).color = (228,241,253)
            wafermap_sheet.range((last_row+1, last_col+1)).api.Font.Color = dark_blue
            wafermap_sheet.range((last_row+1, last_col+1)).api.Font.Bold = True  # bold "No." cell

            # --- Also bold the original Row 1 and Column A ---
            wafermap_sheet.range((1,1),(1,last_col)).api.Font.Bold = True
            wafermap_sheet.range((1,1),(last_row,1)).api.Font.Bold = True
            
            # --- Remove gridlines from wafermap sheet ---
            wafermap_sheet.api.Parent.Windows(1).DisplayGridlines = False

            # --- Alignment (center everything including mirrored row/col) ---
            used_range = wafermap_sheet.range((1,1),(last_row+1,last_col+1))
            used_range.api.HorizontalAlignment = -4108  # xlCenter
            used_range.api.VerticalAlignment = -4108    # xlCenter

            
            # --- Borders ---
            used_range = wafermap_sheet.range((1,1),(last_row+1,last_col+1))
            used_range.api.Borders.Weight = 2

            wb_xlw.save()
            wb_xlw.close()
            app.quit()
            self.show_status(f"\n✅ Wafermap created on {sheet_name} sheet.")

            # --- Reopen workbook to safely delete pivot sheet ---
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)

            try:
                pivot_sheet = wb_xlw.sheets["Wafermap Pivot Table"]
                # Activate another sheet first
                wb_xlw.sheets[0].activate()
                pivot_sheet.delete()
                #self.show_status("\n🗑️ Wafermap Pivot Table sheet deleted after reopen.")
            except Exception as e:
                #self.show_status(f"\n⚠️ Could not delete Wafermap Pivot Table: {e}", color="#d32f2f")
                pass

        except Exception as e:
            # Ensure logger is configured
            self.logger.setup_on_error()

            # Log the error
            logging.critical(f"Unexpected error: {e}", exc_info=True)

            # Show status in GUI
            self.show_status(f"❌ Unexpected error: {e}", color="#d32f2f")
        finally:
            if wb_xlw:
                try:
                    wb_xlw.save()   # ✅ persist changes (sheet deletion, edits, etc.)
                    wb_xlw.close()
                except:
                    pass
            if app:
                try:
                    app.quit()
                except:
                    pass

    def clear_all(self):
        """
        Reset all GUI state:
        - Clears file path and combobox values
        - Wipes status box
        - Resets stored attributes
        """

        try:
            # Reset file path and filter selections
            self.path_var.set("")
            if hasattr(self, "filter_var"):
                self.filter_var.set("")
            if hasattr(self, "filter_dropdown"):
                self.filter_dropdown['values'] = []

            # Clear status box
            self.show_status("", clear=True)
            self.show_status("✅ Cleared all selections.")

            # Reset any stored attributes
            if hasattr(self, "out_file"):
                self.out_file = None
            if hasattr(self, "sheet_name"):
                self.sheet_name = None
            if hasattr(self, "base_name"):
                self.base_name = None
            if hasattr(self, "raw_items"):
                self.raw_items = []

        except Exception as e:
            # Ensure logger is configured
            self.logger.setup_on_error()

            # Log the error
            logging.critical(f"Unexpected error: {e}", exc_info=True)

            # Show status in GUI
            self.show_status(f"❌ Unexpected error: {e}", color="#d32f2f")
            
if __name__ == "__main__":
    root = tk.Tk()
    root.title("CSV Workflow Automation Tool v1.1.2")
    root.iconbitmap(resource_path("sprout.ico"))  # subtle window icon only
    app = CSVWorkflowAutomationGUI(root)
    root.mainloop()

