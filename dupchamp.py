"""
DupChamp - Duplicate Row Finder
Scans folder files (CSV/XLSX) against a test file to find duplicate rows
based on user-defined column mappings.
"""

import os
import json
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import pandas as pd

# --- Constants ---
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dupchamp_config.json")
SCAN_EXTENSIONS = (".csv", ".xlsx", ".xls")
TEST_EXTENSIONS = (".csv", ".xlsx", ".xls")


# --- Config Persistence ---
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# --- File Reading Helpers ---
def read_file_to_df(filepath, sheet_name=None):
    ext = Path(filepath).suffix.lower()
    if ext == ".csv":
        # Try multiple encodings for Hebrew CSV files
        for enc in ("utf-8-sig", "utf-8", "cp1255", "iso-8859-8"):
            try:
                return pd.read_csv(filepath, dtype=str, keep_default_na=False, encoding=enc)
            except (UnicodeDecodeError, UnicodeError):
                continue
        return pd.read_csv(filepath, dtype=str, keep_default_na=False, encoding="latin-1")
    elif ext in (".xlsx", ".xls"):
        kwargs = {"dtype": str, "keep_default_na": False, "engine": "openpyxl"}
        if sheet_name:
            kwargs["sheet_name"] = sheet_name
        return pd.read_excel(filepath, **kwargs)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def get_sheet_names(filepath):
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls"):
        xls = pd.ExcelFile(filepath, engine="openpyxl")
        return xls.sheet_names
    return []


def get_folder_files(folder_path):
    files = []
    for f in os.listdir(folder_path):
        if Path(f).suffix.lower() in SCAN_EXTENSIONS:
            files.append(os.path.join(folder_path, f))
    return files


def col_index_to_letter(index):
    """Convert 0-based column index to Excel column letter (A, B, ..., Z, AA, AB, ...)."""
    result = ""
    while True:
        result = chr(ord('A') + index % 26) + result
        index = index // 26 - 1
        if index < 0:
            break
    return result


def format_col_display(columns):
    """Return list of display strings like 'A - ColumnName' for dropdown."""
    return [f"{col_index_to_letter(i)} - {col}" for i, col in enumerate(columns)]


def parse_col_display(display_value):
    """Extract actual column name from display string like 'A - ColumnName'."""
    if " - " in display_value:
        return display_value.split(" - ", 1)[1]
    return display_value


def get_scan_columns(folder_path):
    """Read column names from all files in folder, return union of columns."""
    all_cols = []
    seen = set()
    for fpath in get_folder_files(folder_path):
        try:
            df = read_file_to_df(fpath)
            for c in df.columns:
                if c not in seen:
                    all_cols.append(c)
                    seen.add(c)
        except Exception:
            continue
    return all_cols


# --- Main Application ---
class DupChampApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DupChamp - Duplicate Row Finder")
        self.root.geometry("1100x800")
        self.root.minsize(900, 650)

        self.config = load_config()
        self.test_df = None
        self.test_columns = []
        self.scan_columns = []
        self.mapping_rows = []  # list of (test_combo, scan_combo, row_frame)
        self.is_running = False

        self._build_ui()
        self._restore_from_config()

    def _build_ui(self):
        # --- Top Frame: File/Folder Selection ---
        top_frame = ttk.LabelFrame(self.root, text="הגדרות", padding=10)
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        # Row 0: Scan folder (CSV/XLSX)
        ttk.Label(top_frame, text="קבצים נבדקים (תיקייה):").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))
        self.folder_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.folder_var, width=70).grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(top_frame, text="בחר תיקייה...", command=self._select_folder).grid(row=0, column=0, padx=5)

        # Row 1: Test file (CSV or Excel)
        ttk.Label(top_frame, text="קובץ נבדק (CSV/Excel):").grid(row=1, column=2, sticky=tk.E, padx=(0, 5), pady=(5, 0))
        self.testfile_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.testfile_var, width=70).grid(row=1, column=1, sticky=tk.EW, padx=5, pady=(5, 0))
        ttk.Button(top_frame, text="בחר קובץ...", command=self._select_test_file).grid(row=1, column=0, padx=5, pady=(5, 0))

        # Row 2: Sheet selection (for Excel test files only)
        ttk.Label(top_frame, text="גיליון בקובץ נבדק:").grid(row=2, column=2, sticky=tk.E, padx=(0, 5), pady=(5, 0))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(top_frame, textvariable=self.sheet_var, state="readonly", width=40)
        self.sheet_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=(5, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)
        ttk.Label(top_frame, text="(רלוונטי רק לקובץ Excel)", foreground="gray").grid(row=2, column=0, sticky=tk.W, padx=5, pady=(5, 0))

        top_frame.columnconfigure(1, weight=1)

        # --- Column Mapping Frame ---
        map_frame = ttk.LabelFrame(self.root, text="מיפוי עמודות - הגדרת כפילויות", padding=10)
        map_frame.pack(fill=tk.X, padx=10, pady=5)

        # Header row
        header_frame = ttk.Frame(map_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        header_frame.columnconfigure(0, weight=0, minsize=40)
        header_frame.columnconfigure(1, weight=1)
        header_frame.columnconfigure(2, weight=0, minsize=40)
        header_frame.columnconfigure(3, weight=1)

        ttk.Label(header_frame, text="קבצים נבדקים", font=("Segoe UI", 10, "bold"), anchor=tk.CENTER).grid(row=0, column=0, columnspan=2, sticky=tk.EW)
        ttk.Label(header_frame, text="↔", font=("Segoe UI", 12, "bold"), anchor=tk.CENTER).grid(row=0, column=2)
        ttk.Label(header_frame, text="קובץ נבדק", font=("Segoe UI", 10, "bold"), anchor=tk.CENTER).grid(row=0, column=3, sticky=tk.EW)

        # Scrollable mapping area
        self.map_canvas = tk.Canvas(map_frame, height=130, highlightthickness=0)
        map_scrollbar = ttk.Scrollbar(map_frame, orient=tk.VERTICAL, command=self.map_canvas.yview)
        self.map_inner_frame = ttk.Frame(self.map_canvas)

        self.map_inner_frame.bind("<Configure>", lambda e: self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all")))
        self._canvas_window = self.map_canvas.create_window((0, 0), window=self.map_inner_frame, anchor=tk.NW)
        self.map_canvas.configure(yscrollcommand=map_scrollbar.set)
        self.map_canvas.bind("<Configure>", lambda e: self.map_canvas.itemconfig(self._canvas_window, width=e.width))

        self.map_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        map_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Placeholder
        self.map_placeholder = ttk.Label(self.map_inner_frame, text="טען קובץ בדיקה ובחר תיקיית סריקה כדי להגדיר מיפוי עמודות", foreground="gray")
        self.map_placeholder.pack(anchor=tk.CENTER, pady=20)

        # Add mapping button
        btn_frame = ttk.Frame(map_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        self.add_map_btn = ttk.Button(btn_frame, text="+ הוסף מיפוי עמודות", command=self._add_mapping_row, state=tk.DISABLED)
        self.add_map_btn.pack(side=tk.RIGHT)

        # --- Run Button & Progress ---
        run_frame = ttk.Frame(self.root, padding=(10, 5))
        run_frame.pack(fill=tk.X, padx=10)

        self.run_btn = ttk.Button(run_frame, text="▶  הרץ בדיקת כפילויות", command=self._run_check)
        self.run_btn.pack(side=tk.RIGHT, padx=(10, 0))

        self.stats_label = ttk.Label(run_frame, text="")
        self.stats_label.pack(side=tk.RIGHT, padx=10)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(run_frame, variable=self.progress_var, maximum=100, length=300)
        self.progress_bar.pack(side=tk.RIGHT, padx=5)

        self.progress_label = ttk.Label(run_frame, text="")
        self.progress_label.pack(side=tk.RIGHT)

        # --- Results Frame ---
        results_frame = ttk.LabelFrame(self.root, text="תוצאות", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

        tree_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(
            results_frame,
            columns=("test_row", "dup_file", "dup_row", "values"),
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
        )
        self.tree.heading("test_row", text="שורה בקובץ בדיקה", anchor=tk.CENTER)
        self.tree.heading("dup_file", text="קובץ עם כפילות", anchor=tk.CENTER)
        self.tree.heading("dup_row", text="שורה בקובץ", anchor=tk.CENTER)
        self.tree.heading("values", text="ערכים כפולים", anchor=tk.CENTER)

        self.tree.column("test_row", width=120, anchor=tk.CENTER)
        self.tree.column("dup_file", width=350, anchor=tk.W)
        self.tree.column("dup_row", width=100, anchor=tk.CENTER)
        self.tree.column("values", width=400, anchor=tk.W)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # --- Export Button ---
        export_frame = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        export_frame.pack(fill=tk.X)
        self.export_btn = ttk.Button(export_frame, text="ייצוא תוצאות לקובץ CSV", command=self._export_results, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT)

        self.result_count_label = ttk.Label(export_frame, text="")
        self.result_count_label.pack(side=tk.RIGHT, padx=10)

    # --- Restore ---
    def _restore_from_config(self):
        if "folder_path" in self.config:
            path = self.config["folder_path"]
            self.folder_var.set(path)
            if os.path.isdir(path):
                self.scan_columns = get_scan_columns(path)

        if "test_file" in self.config:
            path = self.config["test_file"]
            self.testfile_var.set(path)
            if os.path.exists(path):
                self._populate_sheets(path)
                saved_sheet = self.config.get("sheet_name", "")
                if saved_sheet and saved_sheet in (self.sheet_combo["values"] or []):
                    self.sheet_var.set(saved_sheet)
                self._load_test_file_silent(path)

        # Restore saved mappings
        self._try_restore_mappings()

    def _try_restore_mappings(self):
        saved_mappings = self.config.get("column_mappings", [])
        if self.test_columns and self.scan_columns:
            self.add_map_btn.config(state=tk.NORMAL)
            if self.map_placeholder.winfo_exists():
                self.map_placeholder.destroy()
            if saved_mappings:
                for m in saved_mappings:
                    test_col = m.get("test", "")
                    scan_col = m.get("scan", "")
                    self._add_mapping_row(test_col, scan_col)
            if not self.mapping_rows:
                self._add_mapping_row()

    # --- Folder ---
    def _select_folder(self):
        initial = self.folder_var.get() or None
        folder = filedialog.askdirectory(title="בחר תיקיית סריקה", initialdir=initial)
        if folder:
            self.folder_var.set(folder)
            self.config["folder_path"] = folder
            save_config(self.config)
            self.scan_columns = get_scan_columns(folder)
            self._refresh_mapping_combos()

    # --- Test File ---
    def _select_test_file(self):
        initial_dir = os.path.dirname(self.testfile_var.get()) if self.testfile_var.get() else None
        filepath = filedialog.askopenfilename(
            title="בחר קובץ בדיקה",
            initialdir=initial_dir,
            filetypes=[("CSV / Excel", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if filepath:
            self.testfile_var.set(filepath)
            self.config["test_file"] = filepath
            save_config(self.config)
            self._populate_sheets(filepath)
            self._load_test_file(filepath)

    def _populate_sheets(self, filepath):
        sheets = get_sheet_names(filepath)
        if sheets:
            self.sheet_combo["values"] = sheets
            saved_sheet = self.config.get("sheet_name", "")
            if saved_sheet and saved_sheet in sheets:
                self.sheet_var.set(saved_sheet)
            elif "אקסל חשבונות מפורטים" in sheets:
                self.sheet_var.set("אקסל חשבונות מפורטים")
                self.config["sheet_name"] = "אקסל חשבונות מפורטים"
                save_config(self.config)
            else:
                self.sheet_var.set(sheets[0])
            self.sheet_combo.config(state="readonly")
        else:
            self.sheet_combo["values"] = []
            self.sheet_var.set("")
            self.sheet_combo.config(state="disabled")

    def _on_sheet_selected(self, event=None):
        sheet = self.sheet_var.get()
        self.config["sheet_name"] = sheet
        save_config(self.config)
        filepath = self.testfile_var.get()
        if filepath and os.path.exists(filepath):
            self._load_test_file(filepath)

    def _load_test_file_silent(self, filepath):
        try:
            sheet = self.sheet_var.get() or None
            self.test_df = read_file_to_df(filepath, sheet_name=sheet)
            self.test_columns = list(self.test_df.columns)
        except Exception:
            pass

    def _load_test_file(self, filepath):
        try:
            sheet = self.sheet_var.get() or None
            self.test_df = read_file_to_df(filepath, sheet_name=sheet)
            self.test_columns = list(self.test_df.columns)
            self._refresh_mapping_combos()
        except Exception as e:
            messagebox.showerror("שגיאה בטעינת קובץ", str(e))

    # --- Column Mapping UI ---
    def _refresh_mapping_combos(self):
        """Update all combo boxes with current column lists."""
        if self.test_columns and self.scan_columns:
            self.add_map_btn.config(state=tk.NORMAL)
            if self.map_placeholder.winfo_exists():
                self.map_placeholder.destroy()
            test_display = format_col_display(self.test_columns)
            scan_display = format_col_display(self.scan_columns)
            # Update existing comboboxes
            for test_combo, scan_combo, _ in self.mapping_rows:
                current_test = test_combo.get()
                current_scan = scan_combo.get()
                test_combo["values"] = test_display
                scan_combo["values"] = scan_display
                if current_test and parse_col_display(current_test) not in self.test_columns:
                    test_combo.set("")
                if current_scan and parse_col_display(current_scan) not in self.scan_columns:
                    scan_combo.set("")
            # If no mapping rows exist, add one
            if not self.mapping_rows:
                self._add_mapping_row()
        else:
            self.add_map_btn.config(state=tk.DISABLED)

    def _add_mapping_row(self, default_test="", default_scan=""):
        row_frame = ttk.Frame(self.map_inner_frame)
        row_frame.pack(fill=tk.X, pady=2)
        row_frame.columnconfigure(0, weight=0, minsize=40)
        row_frame.columnconfigure(1, weight=1)
        row_frame.columnconfigure(2, weight=0, minsize=40)
        row_frame.columnconfigure(3, weight=1)

        test_display = format_col_display(self.test_columns)
        scan_display = format_col_display(self.scan_columns)

        # Remove button (leftmost)
        remove_btn = ttk.Button(row_frame, text="✕", width=3,
                                command=lambda rf=row_frame: self._remove_mapping_row(rf))
        remove_btn.grid(row=0, column=0, padx=(0, 5))

        # Scan column combo (left side)
        scan_combo = ttk.Combobox(row_frame, values=scan_display, state="readonly", width=35)
        scan_combo.grid(row=0, column=1, sticky=tk.EW, padx=5)
        if default_scan:
            for dv in scan_display:
                if parse_col_display(dv) == default_scan:
                    scan_combo.set(dv)
                    break

        # Arrow label (center)
        ttk.Label(row_frame, text="↔", font=("Segoe UI", 12, "bold")).grid(row=0, column=2, padx=5)

        # Test column combo (right side)
        test_combo = ttk.Combobox(row_frame, values=test_display, state="readonly", width=35)
        test_combo.grid(row=0, column=3, sticky=tk.EW, padx=5)
        if default_test:
            for dv in test_display:
                if parse_col_display(dv) == default_test:
                    test_combo.set(dv)
                    break

        self.mapping_rows.append((test_combo, scan_combo, row_frame))

        # Save on selection change
        test_combo.bind("<<ComboboxSelected>>", lambda e: self._save_mappings())
        scan_combo.bind("<<ComboboxSelected>>", lambda e: self._save_mappings())

    def _remove_mapping_row(self, row_frame):
        self.mapping_rows = [(t, s, f) for t, s, f in self.mapping_rows if f != row_frame]
        row_frame.destroy()
        self._save_mappings()

    def _save_mappings(self):
        mappings = []
        for test_combo, scan_combo, _ in self.mapping_rows:
            t = parse_col_display(test_combo.get()) if test_combo.get() else ""
            s = parse_col_display(scan_combo.get()) if scan_combo.get() else ""
            if t or s:
                mappings.append({"test": t, "scan": s})
        self.config["column_mappings"] = mappings
        save_config(self.config)

    def _get_mappings(self):
        """Return list of (test_col, scan_col) tuples for valid mappings (actual column names)."""
        mappings = []
        for test_combo, scan_combo, _ in self.mapping_rows:
            t = parse_col_display(test_combo.get()) if test_combo.get() else ""
            s = parse_col_display(scan_combo.get()) if scan_combo.get() else ""
            if t and s:
                mappings.append((t, s))
        return mappings

    # --- Run ---
    def _run_check(self):
        folder = self.folder_var.get()
        test_file = self.testfile_var.get()

        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("שגיאה", "יש לבחור תיקיית סריקה תקינה")
            return
        if not test_file or not os.path.isfile(test_file):
            messagebox.showwarning("שגיאה", "יש לבחור קובץ בדיקה תקין")
            return

        mappings = self._get_mappings()
        if not mappings:
            messagebox.showwarning("שגיאה", "יש להגדיר לפחות מיפוי עמודות אחד (שני הצדדים חייבים להיות מלאים)")
            return

        test_cols = [m[0] for m in mappings]
        scan_cols = [m[1] for m in mappings]

        scan_files = get_folder_files(folder)
        test_file_abs = os.path.abspath(test_file)
        scan_files = [f for f in scan_files if os.path.abspath(f) != test_file_abs]

        if not scan_files:
            messagebox.showwarning("שגיאה", "לא נמצאו קבצי CSV/XLSX בתיקיית הסריקה")
            return

        # Reload test file
        try:
            sheet = self.sheet_var.get() or None
            self.test_df = read_file_to_df(test_file, sheet_name=sheet)
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בקריאת קובץ הבדיקה:\n{e}")
            return

        # Verify test columns exist
        missing = [c for c in test_cols if c not in self.test_df.columns]
        if missing:
            messagebox.showwarning("שגיאה", f"העמודות הבאות לא נמצאו בקובץ הבדיקה:\n{', '.join(missing)}")
            return

        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.export_btn.config(state=tk.DISABLED)
        self.result_count_label.config(text="")

        self.is_running = True
        self.run_btn.config(state=tk.DISABLED)

        test_rows = len(self.test_df)
        self.stats_label.config(text=f"שורות בקובץ בדיקה: {test_rows} | קבצים לסריקה: {len(scan_files)}")

        thread = threading.Thread(
            target=self._do_scan,
            args=(self.test_df, scan_files, test_cols, scan_cols, mappings),
            daemon=True,
        )
        thread.start()

    def _do_scan(self, test_df, scan_files, test_cols, scan_cols, mappings):
        results = []
        test_keys = test_df[test_cols].astype(str)

        total_scan_rows = 0
        scan_dfs = []
        for fpath in scan_files:
            try:
                df = read_file_to_df(fpath)
                scan_dfs.append((fpath, df))
                total_scan_rows += len(df)
            except Exception:
                scan_dfs.append((fpath, None))

        total_work = len(test_df) * len(scan_dfs)
        processed = 0

        for file_idx, (fpath, scan_df) in enumerate(scan_dfs):
            if scan_df is None:
                processed += len(test_df)
                pct = (processed / total_work) * 100 if total_work > 0 else 100
                self.root.after(0, self._update_progress, pct)
                continue

            # Check which scan columns exist in this file
            missing_scan = [c for c in scan_cols if c not in scan_df.columns]
            if missing_scan:
                processed += len(test_df)
                pct = (processed / total_work) * 100 if total_work > 0 else 100
                self.root.after(0, self._update_progress, pct)
                continue

            scan_keys = scan_df[scan_cols].astype(str)

            # Build index from scan file: key_tuple -> list of row numbers
            scan_index = {}
            for scan_row_idx in range(len(scan_keys)):
                key = tuple(scan_keys.iloc[scan_row_idx])
                if key not in scan_index:
                    scan_index[key] = []
                scan_index[key].append(scan_row_idx + 2)  # +2: header + 0-based

            # Match test rows
            for test_row_idx in range(len(test_keys)):
                test_key = tuple(test_keys.iloc[test_row_idx])
                if test_key in scan_index:
                    for scan_row_num in scan_index[test_key]:
                        values_str = " | ".join(
                            f"{tc}={v} → {sc}"
                            for (tc, sc), v in zip(mappings, test_key)
                        )
                        results.append({
                            "test_row": test_row_idx + 2,
                            "dup_file": fpath,
                            "dup_file_name": os.path.basename(fpath),
                            "dup_row": scan_row_num,
                            "values": values_str,
                        })

                processed += 1
                if processed % 500 == 0 or processed == total_work:
                    pct = (processed / total_work) * 100 if total_work > 0 else 100
                    self.root.after(0, self._update_progress, pct)

        self.root.after(0, self._scan_complete, results, total_scan_rows)

    def _update_progress(self, pct):
        self.progress_var.set(pct)
        self.progress_label.config(text=f"{pct:.1f}%")

    def _scan_complete(self, results, total_scan_rows):
        self.is_running = False
        self.run_btn.config(state=tk.NORMAL)
        self.progress_var.set(100)
        self.progress_label.config(text="100%")

        stats_text = self.stats_label.cget("text")
        self.stats_label.config(text=f"{stats_text} | סה\"כ שורות בתיקייה: {total_scan_rows}")

        if not results:
            self.result_count_label.config(text="לא נמצאו כפילויות ✓")
            messagebox.showinfo("תוצאות", "לא נמצאו כפילויות!")
            return

        unique_test_rows = len(set(r["test_row"] for r in results))
        self.result_count_label.config(text=f"נמצאו {len(results)} כפילויות ב-{unique_test_rows} שורות")

        for r in results:
            self.tree.insert("", tk.END, values=(
                r["test_row"],
                r["dup_file_name"],
                r["dup_row"],
                r["values"],
            ), tags=(r["dup_file"],))

        self.export_btn.config(state=tk.NORMAL)
        self._results_data = results

    def _on_tree_double_click(self, event):
        item = self.tree.selection()
        if not item:
            return
        tags = self.tree.item(item[0], "tags")
        if tags:
            filepath = tags[0]
            try:
                os.startfile(filepath)
            except Exception:
                subprocess.Popen(["explorer", "/select,", filepath])

    def _export_results(self):
        if not hasattr(self, "_results_data") or not self._results_data:
            return

        filepath = filedialog.asksaveasfilename(
            title="שמור תוצאות",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not filepath:
            return

        df = pd.DataFrame(self._results_data)
        df = df[["test_row", "dup_file", "dup_row", "values"]]
        df.columns = ["שורה בקובץ בדיקה", "קובץ עם כפילות", "שורה בקובץ", "ערכים כפולים"]
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
        messagebox.showinfo("ייצוא", f"התוצאות יוצאו בהצלחה ל:\n{filepath}")


def main():
    root = tk.Tk()
    default_font = ("Segoe UI", 10)
    root.option_add("*Font", default_font)

    style = ttk.Style()
    style.configure("TLabel", font=default_font)
    style.configure("TButton", font=default_font)
    style.configure("TCheckbutton", font=default_font)
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("Treeview", font=default_font, rowheight=25)

    app = DupChampApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
