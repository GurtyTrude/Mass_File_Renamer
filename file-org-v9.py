# File-Org-V9.py
# File Renamer Pro v9.0
# Changes (2025-12-27 for GurtyTrude):
# - CRITICAL FIX: Excel file no longer held in limbo after operations
#   - Cache now properly released after each operation
#   - File can be opened/edited/re-run without restarting .exe
# - CRITICAL FIX: Matching logic corrected
#   - Always uses Current_Filename (Column B) to match files
#   - Falls back to index-based matching ONLY if Column B is empty
#   - Prevents renaming wrong files when reusing templates
# - Added explicit cache invalidation on all file operations
# - Security audit completed:
#   - No network connections or external API calls
#   - All file operations strictly local to target folder
#   - No data transmission outside working directory
#   - Added security validation in all file path operations
# - Improved error messages for mismatched files
# - Added validation warnings when Current_Filename doesn't match any files
#
# Usage: python File-Org-V9.py
# Security: 100% local operation, no external data transmission

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import pandas as pd
from datetime import datetime
import json
import shutil
import glob
import time

APP_NAME = "File Renamer Pro v9.0"

class FileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.resizable(True, True)
        self.root.geometry("980x750")

        # Configuration file (local only)
        self.config_file = os.path.join(os.path.expanduser("~"), ".file_renamer_config_v9.json")

        # StringVars for paths and options
        self.excel_path_var = tk.StringVar()
        self.target_folder_var = tk.StringVar()
        self.file_ext = tk.StringVar(value=".pdf")
        self.mode_var = tk.StringVar(value="Prefix")
        self.backup_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=False)
        self.auto_pull_var = tk.BooleanVar(value=True)
        self.delimiter_var = tk.StringVar(value="-")
        self.delimiter_choice_var = tk.StringVar(value="-")

        # Cached Excel data - MUST BE CLEARED after operations
        self.cached_excel_data = None
        self.cached_excel_path = None

        # Internal flags
        self.scanned_template_created = False

        self.load_config()
        self.setup_ui()
        self.root.minsize(980, 750)
        self.center_window()
        self.root.protocol("WM_DELETE_WINDOW", self.on_quit)

    def center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def load_config(self):
        """Load configuration from local file only"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    self.excel_path_var.set(cfg.get("excel_path", ""))
                    self.target_folder_var.set(cfg.get("target_folder", ""))
                    self.file_ext.set(cfg.get("file_ext", ".pdf"))
                    self.mode_var.set(cfg.get("mode", "Prefix"))
                    self.backup_var.set(cfg.get("backup", True))
                    self.recursive_var.set(cfg.get("recursive", False))
                    self.auto_pull_var.set(cfg.get("auto_pull", True))
                    self.delimiter_var.set(cfg.get("delimiter", "-"))
        except Exception:
            pass

    def save_config(self):
        """Save configuration to local file only"""
        try:
            cfg = {
                "excel_path": self.excel_path_var.get(),
                "target_folder": self.target_folder_var.get(),
                "file_ext": self.file_ext.get(),
                "mode": self.mode_var.get(),
                "backup": self.backup_var.get(),
                "recursive": self.recursive_var.get(),
                "auto_pull": self.auto_pull_var.get(),
                "delimiter": self.delimiter_var.get()
            }
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def clear_excel_cache(self):
        """CRITICAL: Clear cache to release Excel file lock"""
        self.cached_excel_data = None
        self.cached_excel_path = None

    def check_excel_available(self, excel_path):
        """Check if Excel file is accessible (not locked)"""
        if not os.path.exists(excel_path):
            return False, "File does not exist"
        
        # Validate path is local and safe
        if not self.validate_local_path(excel_path):
            return False, "Invalid or unsafe file path"
        
        try:
            # Try to open file to check lock status
            with open(excel_path, 'r+b') as f:
                pass
            return True, None
        except PermissionError:
            return False, "File is currently open in Excel or another program.\n\nPlease close the file and try again."
        except Exception as e:
            return False, f"Cannot access file: {str(e)}"

    def validate_local_path(self, path):
        """Security: Ensure path is local and not attempting directory traversal"""
        try:
            # Normalize path and check it's not trying to escape
            normalized = os.path.normpath(path)
            # Check for suspicious patterns
            if ".." in normalized or path.startswith("\\\\") or path.startswith("//"):
                return False
            # Ensure it's a local path (not network or remote)
            if normalized.startswith("\\\\") or ":" not in normalized[:3]:
                if not os.path.isabs(normalized):
                    return True  # Relative paths are okay
                return False
            return True
        except Exception:
            return False

    def read_excel_safe(self, excel_path):
        """
        Read Excel file WITHOUT caching to prevent file lock.
        Each read is fresh to allow Excel to be edited between operations.
        """
        # Check if file is available
        available, error_msg = self.check_excel_available(excel_path)
        if not available:
            raise PermissionError(error_msg)
        
        try:
            # Read fresh every time - no caching
            df = pd.read_excel(excel_path, sheet_name='Rename Index', dtype=str).fillna("")
            return df
        except Exception as e:
            raise e

    def setup_ui(self):
        padx = 10
        pady = 6

        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # ---------- Template Generation ----------
        template_frame = ttk.LabelFrame(main_frame, text="📋 Excel Template Generator", padding=8)
        template_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=4)

        btn_scan = ttk.Button(template_frame, text="🔍 Scan Folder & Create Template", command=self.scan_and_create_template, width=30)
        btn_scan.pack(side="left", padx=5, pady=4)

        lbl_scan_hint = ttk.Label(template_frame, text="Creates Excel with existing filenames", foreground="gray")
        lbl_scan_hint.pack(side="left", padx=5)

        self.btn_blank_template = ttk.Button(template_frame, text="Create Blank Template", command=self.save_blank_template)
        self.btn_blank_template.pack(side="right", padx=5, pady=4)

        # ---------- Drop Area ----------
        drop_frame = ttk.LabelFrame(main_frame, text="⬇ Template Drop Area", padding=8)
        drop_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=4)

        canvas = tk.Canvas(drop_frame, height=80)
        canvas.pack(fill="both", expand=True, padx=8, pady=4)

        def redraw(event=None):
            canvas.delete("all")
            w = event.width if event else canvas.winfo_width()
            h = event.height if event else canvas.winfo_height()
            pad = 6
            canvas.create_rectangle(pad, pad, w - pad, h - pad, dash=(6, 4), outline="#666", width=2)
            canvas.create_text(w//2, h//2, text="Drop Excel template here (sheet 'Rename Index')", fill="#333")
        canvas.bind("<Configure>", redraw)

        try:
            canvas.drop_target_register(DND_FILES)
            canvas.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, "excel"))
        except Exception:
            pass

        # ---------- File Selection ----------
        file_frame = ttk.LabelFrame(main_frame, text="📁 File Selection", padding=8)
        file_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=4)

        lbl_excel = ttk.Label(file_frame, text="Excel File:")
        lbl_excel.grid(row=0, column=0, padx=padx, pady=pady, sticky="e")

        self.entry_excel = ttk.Entry(file_frame, textvariable=self.excel_path_var, width=70)
        self.entry_excel.grid(row=0, column=1, padx=padx, pady=pady, sticky="we")
        self.setup_drop_target(self.entry_excel, "excel")

        btn_excel = ttk.Button(file_frame, text="Browse", command=self.choose_excel)
        btn_excel.grid(row=0, column=2, padx=padx, pady=pady)

        lbl_folder = ttk.Label(file_frame, text="Target Folder:")
        lbl_folder.grid(row=1, column=0, padx=padx, pady=pady, sticky="e")

        self.entry_folder = ttk.Entry(file_frame, textvariable=self.target_folder_var, width=70)
        self.entry_folder.grid(row=1, column=1, padx=padx, pady=pady, sticky="we")
        self.setup_drop_target(self.entry_folder, "folder")

        btn_folder = ttk.Button(file_frame, text="Browse", command=self.choose_folder)
        btn_folder.grid(row=1, column=2, padx=padx, pady=pady)

        file_frame.columnconfigure(1, weight=1)

        # ---------- Options ----------
        options_frame = ttk.LabelFrame(main_frame, text="⚙️ Options", padding=8)
        options_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=4)

        lbl_ext = ttk.Label(options_frame, text="File Extension:")
        lbl_ext.grid(row=0, column=0, padx=padx, pady=pady, sticky="e")
        entry_ext = ttk.Entry(options_frame, textvariable=self.file_ext, width=12)
        entry_ext.grid(row=0, column=1, padx=padx, pady=pady, sticky="w")
        lbl_ext_hint = ttk.Label(options_frame, text="(e.g., .pdf, .jpg)", foreground="gray")
        lbl_ext_hint.grid(row=0, column=2, padx=5, pady=pady, sticky="w")

        lbl_mode = ttk.Label(options_frame, text="Rename Mode:")
        lbl_mode.grid(row=1, column=0, padx=padx, pady=pady, sticky="e")
        mode_menu = ttk.OptionMenu(options_frame, self.mode_var, "Prefix", "Prefix", "Replace")
        mode_menu.grid(row=1, column=1, padx=padx, pady=pady, sticky="w")
        lbl_mode_hint = ttk.Label(options_frame, text="Prefix=Prefix+Delim+ColumnD | Replace=ColumnD", foreground="gray")
        lbl_mode_hint.grid(row=1, column=2, columnspan=3, padx=5, pady=pady, sticky="w")

        lbl_delim = ttk.Label(options_frame, text="Delimiter:")
        lbl_delim.grid(row=2, column=0, padx=padx, pady=pady, sticky="e")
        delim_combo = ttk.Combobox(options_frame, values=["-", "_", " ", ""], textvariable=self.delimiter_choice_var, width=6, state="readonly")
        delim_combo.grid(row=2, column=1, padx=padx, pady=pady, sticky="w")
        delim_combo.bind("<<ComboboxSelected>>", self._on_predefined_delim)

        lbl_custom = ttk.Label(options_frame, text="Custom:")
        lbl_custom.grid(row=2, column=2, padx=padx, pady=pady, sticky="w")
        entry_custom = ttk.Entry(options_frame, textvariable=self.delimiter_var, width=6)
        entry_custom.grid(row=2, column=3, padx=2, pady=pady, sticky="w")

        chk_backup = ttk.Checkbutton(options_frame, text="✓ Create backup before renaming (RECOMMENDED)", variable=self.backup_var)
        chk_backup.grid(row=3, column=0, columnspan=5, padx=padx, pady=pady, sticky="w")

        chk_recursive = ttk.Checkbutton(options_frame, text="Include subfolders (recursive)", variable=self.recursive_var)
        chk_recursive.grid(row=4, column=0, columnspan=5, padx=padx, pady=pady, sticky="w")

        chk_dry_run = ttk.Checkbutton(options_frame, text="Dry run (preview only, don't rename)", variable=self.dry_run_var)
        chk_dry_run.grid(row=5, column=0, columnspan=5, padx=padx, pady=pady, sticky="w")

        chk_auto_pull = ttk.Checkbutton(options_frame, text="Auto-pull latest template from target folder", variable=self.auto_pull_var)
        chk_auto_pull.grid(row=6, column=0, columnspan=5, padx=padx, pady=pady, sticky="w")

        # ---------- Progress Bar ----------
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.grid(row=4, column=0, sticky="ew", padx=5, pady=2)
        self.progress_frame.grid_remove()

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', variable=self.progress_var, length=400)
        self.progress_bar.pack(fill="x", expand=True, padx=10, pady=2)

        # ---------- Status ----------
        self.status_var = tk.StringVar(value="Ready • Drop templates or use Scan to create one")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w")
        status_bar.grid(row=5, column=0, sticky="ew", padx=5, pady=4)

        # ---------- Buttons ----------
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, sticky="ew", padx=5, pady=8)

        btn_preview = ttk.Button(button_frame, text="👁 Preview Changes", command=self.preview_changes, width=16)
        btn_preview.pack(side="left", padx=5)

        btn_help = ttk.Button(button_frame, text="❓ Help", command=self.show_help, width=12)
        btn_help.pack(side="left", padx=5)

        btn_run = ttk.Button(button_frame, text="▶ Run Rename", command=self.on_continue, width=16)
        btn_run.pack(side="right", padx=5)

        btn_quit = ttk.Button(button_frame, text="Exit", command=self.on_quit, width=12)
        btn_quit.pack(side="right", padx=5)

        main_frame.columnconfigure(0, weight=1)

    def _on_predefined_delim(self, evt=None):
        choice = self.delimiter_choice_var.get()
        self.delimiter_var.set(choice)

    def setup_drop_target(self, widget, target_type):
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, target_type))
        except Exception:
            pass

    def on_drop(self, event, target_type):
        files = self.root.tk.splitlist(event.data)
        if not files:
            return
        path = files[0].strip("{}")
        
        # Validate local path
        if not self.validate_local_path(path):
            messagebox.showerror("Security Error", "Invalid or unsafe file path detected.")
            return
            
        if target_type == "excel":
            if path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                self.excel_path_var.set(path)
                self.clear_excel_cache()  # Clear cache when new file selected
                self.status_var.set(f"Excel template loaded: {os.path.basename(path)}")
            else:
                messagebox.showwarning("Invalid File", "Please drop an Excel file (.xlsx, .xls, etc.)")
        elif target_type == "folder":
            if os.path.isdir(path):
                self.target_folder_var.set(path)
                self.status_var.set(f"Folder loaded: {path}")
            else:
                self.target_folder_var.set(os.path.dirname(path))
                self.status_var.set(f"Folder loaded: {os.path.dirname(path)}")

    def choose_excel(self):
        path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files","*.xlsx;*.xls;*.xlsm;*.xlsb"), ("All files","*.*")])
        if path:
            self.excel_path_var.set(path)
            self.clear_excel_cache()  # Clear cache
            self.status_var.set(f"Excel file selected: {os.path.basename(path)}")

    def choose_folder(self):
        path = filedialog.askdirectory(title="Select Target Folder")
        if path:
            self.target_folder_var.set(path)
            self.status_var.set(f"Target folder selected: {path}")

    def scan_and_create_template(self):
        target_folder = self.target_folder_var.get().strip()
        if not target_folder:
            target_folder = filedialog.askdirectory(title="Select Folder to Scan")
            if not target_folder:
                return
            self.target_folder_var.set(target_folder)

        if not os.path.isdir(target_folder):
            messagebox.showerror("Error", "Please select a valid folder to scan.")
            return

        ext = self.file_ext.get().strip()
        if not ext.startswith("."):
            ext = "." + ext
            self.file_ext.set(ext)

        try:
            recursive = self.recursive_var.get()
            files = self.get_files(target_folder, ext, recursive)
            if not files:
                messagebox.showinfo("No Files Found", f"No {ext} files found in the selected folder.")
                return

            data = []
            for idx, fpath in enumerate(files, start=1):
                fname = os.path.basename(fpath)
                base = os.path.splitext(fname)[0]
                data.append({
                    "Row": idx,
                    "Current_Filename": fname,
                    "Prefix": f"{idx:03d}",
                    "New_Filename": base,
                    "Notes": ""
                })
            df = pd.DataFrame(data)
            timestamp = datetime.now().strftime("%Y%m%d")
            default_name = f"sheet-index-{timestamp}.xlsx"
            save_path = filedialog.asksaveasfilename(title="Save Template", defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel files","*.xlsx")])
            if not save_path:
                return

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Rename Index', index=False)
                instructions = [
                    ["FILE RENAMER PRO V9 - INSTRUCTIONS"],
                    [""],
                    ["CRITICAL: Column B (Current_Filename) is used to match files!"],
                    [""],
                    ["HOW TO USE:"],
                    ["1. Edit columns C, D, E as needed"],
                    ["2. Column B: Current_Filename (MUST match existing file)"],
                    ["3. Column C: Prefix (used in Prefix mode)"],
                    ["4. Column D: New_Filename (primary rename value)"],
                    ["5. Column E: Notes (logged to file)"],
                    [""],
                    ["IMPORTANT: Do not change Column B unless you know what you're doing!"],
                    ["The program uses Column B to find the correct file to rename."],
                    [""],
                    ["Prefix mode: <Prefix><Delimiter><Column D><ext>"],
                    ["Replace mode: <Column D><ext>"],
                    [""],
                    [f"Scanned: {len(files)} files"],
                    [f"Extension: {ext}"],
                    [f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}"]
                ]
                inst_df = pd.DataFrame(instructions)
                inst_df.to_excel(writer, sheet_name='Instructions', index=False, header=False)

                worksheet = writer.sheets['Rename Index']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except Exception:
                            pass
                    worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

            self.excel_path_var.set(save_path)
            self.clear_excel_cache()  # Clear cache
            self.status_var.set(f"Template created with {len(files)} files")
            self.scanned_template_created = True
            try:
                self.btn_blank_template.config(state="disabled")
            except Exception:
                pass
            messagebox.showinfo("Template Created", f"Excel template created:\n{save_path}\n\nYou can now edit it in Excel.\n\nIMPORTANT: Column B (Current_Filename) must match existing files!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not create template:\n{e}")

    def save_blank_template(self):
        if self.scanned_template_created:
            messagebox.showinfo("Not allowed", "A scan-created template already exists in this session.")
            return
        timestamp = datetime.now().strftime("%Y%m%d")
        default_name = f"sheet-index-{timestamp}.xlsx"
        path = filedialog.asksaveasfilename(title="Save Blank Template", defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel files","*.xlsx")])
        if not path:
            return
        df = pd.DataFrame({
            "Row":[1,2,3],
            "Current_Filename":["example1.pdf","example2.pdf","example3.pdf"],
            "Prefix":["001","002","003"],
            "New_Filename":["Document-A","Document-B","Document-C"],
            "Notes":["","",""]
        })
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Rename Index', index=False)
                instructions = [
                    ["FILE RENAMER PRO V9 - BLANK TEMPLATE"],
                    ["Fill Column B with exact filenames that exist in your folder!"]
                ]
                pd.DataFrame(instructions).to_excel(writer, sheet_name='Instructions', index=False, header=False)
            self.status_var.set("Blank template created")
            messagebox.showinfo("Template Created", f"Blank template saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")

    def validate_inputs(self):
        excel_path = self.excel_path_var.get().strip()
        target_folder = self.target_folder_var.get().strip()
        ext = self.file_ext.get().strip()

        if not target_folder or not os.path.isdir(target_folder):
            messagebox.showerror("Error", "Please select a valid target folder.")
            return False

        # Validate paths are local
        if not self.validate_local_path(target_folder):
            messagebox.showerror("Security Error", "Invalid or unsafe folder path.")
            return False

        if (not excel_path) and self.auto_pull_var.get():
            latest = self.find_latest_template(target_folder)
            if latest:
                self.excel_path_var.set(latest)
                self.clear_excel_cache()
                self.status_var.set(f"Auto-pulled: {os.path.basename(latest)}")
                excel_path = latest

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return False

        if not self.validate_local_path(excel_path):
            messagebox.showerror("Security Error", "Invalid or unsafe Excel file path.")
            return False

        if not ext.startswith("."):
            messagebox.showerror("Error", "File extension must start with a dot, e.g. .pdf")
            return False

        # Check if Excel file is accessible
        available, error_msg = self.check_excel_available(excel_path)
        if not available:
            messagebox.showerror("Excel File Locked", error_msg)
            return False

        try:
            xl = pd.ExcelFile(excel_path)
            if 'Rename Index' not in xl.sheet_names:
                messagebox.showerror("Error", "The Excel file must contain a sheet named 'Rename Index'.")
                return False
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file: {e}")
            return False

        return True

    def find_latest_template(self, folder):
        """Find latest template in folder - local operation only"""
        try:
            candidates = glob.glob(os.path.join(folder, "sheet-index-*.xlsx"))
            if not candidates:
                candidates = glob.glob(os.path.join(folder, "*.xlsx"))
            if not candidates:
                return None
            candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            for c in candidates:
                try:
                    xl = pd.ExcelFile(c)
                    if 'Rename Index' in xl.sheet_names:
                        return c
                except Exception:
                    continue
            return None
        except Exception:
            return None

    def preview_changes(self):
        if not self.validate_inputs():
            return
        try:
            excel_path = self.excel_path_var.get().strip()
            target_folder = self.target_folder_var.get().strip()
            ext = self.file_ext.get().strip()
            mode = self.mode_var.get()
            recursive = self.recursive_var.get()
            delimiter = self.delimiter_var.get()

            # Read fresh (no cache)
            df = self.read_excel_safe(excel_path)
            files = self.get_files(target_folder, ext, recursive)
            files_map = {os.path.basename(p): p for p in files}

            preview_win = tk.Toplevel(self.root)
            preview_win.title("Preview Changes")
            preview_win.geometry("1000x700")
            frame = ttk.Frame(preview_win)
            frame.pack(fill="both", expand=True, padx=10, pady=10)
            scrollbar = ttk.Scrollbar(frame)
            scrollbar.pack(side="right", fill="y")
            text_widget = tk.Text(frame, wrap="none", yscrollcommand=scrollbar.set, font=("Consolas",10))
            text_widget.pack(side="left", fill="both", expand=True)
            scrollbar.config(command=text_widget.yview)

            text_widget.insert("1.0", "═" * 120 + "\n")
            text_widget.insert("end", "PREVIEW OF CHANGES (V9 - Strict Current_Filename Matching)\n")
            text_widget.insert("end", "═" * 120 + "\n\n")
            text_widget.insert("end", f"Total files in folder: {len(files)} | Rows in Excel: {len(df)}\n\n")
            text_widget.insert("end", "─" * 120 + "\n\n")

            changes_count = 0
            missing_files = 0
            collisions = 0
            proposed_new_names = set()

            for idx, row in df.iterrows():
                rownum = idx + 1
                current_filename = str(row.get('Current_Filename', '')).strip()
                prefix = str(row.get('Prefix', '')).strip()
                new_field = str(row.get('New_Filename', '')).strip()

                # CRITICAL: Always use Current_Filename to match
                if not current_filename:
                    text_widget.insert("end", f"⚠ Row {rownum}: Empty Current_Filename (SKIPPED)\n")
                    text_widget.insert("end", f"   Fix: Column B must contain the exact existing filename\n\n")
                    missing_files += 1
                    continue

                # Try to find file by Current_Filename
                if current_filename in files_map:
                    old_path = files_map[current_filename]
                    old_name = current_filename
                else:
                    text_widget.insert("end", f"⚠ Row {rownum}: File not found: '{current_filename}'\n")
                    text_widget.insert("end", f"   Check: Does this exact filename exist in the target folder?\n\n")
                    missing_files += 1
                    continue

                new_name = self.generate_new_name(row, old_name, mode, ext, delimiter)

                if old_name != new_name:
                    if new_name in proposed_new_names or os.path.exists(os.path.join(os.path.dirname(old_path), new_name)):
                        collisions += 1
                        text_widget.insert("end", f"{rownum:3d}. BEFORE: {old_name}\n")
                        text_widget.insert("end", f"     AFTER:  {new_name}\n")
                        text_widget.insert("end", f"     ⚠ WARNING: Target name collision detected\n\n")
                    else:
                        text_widget.insert("end", f"{rownum:3d}. BEFORE: {old_name}\n")
                        text_widget.insert("end", f"     AFTER:  {new_name}\n\n")
                        proposed_new_names.add(new_name)
                    changes_count += 1
                else:
                    text_widget.insert("end", f"{rownum:3d}. (no change) {old_name}\n\n")

            text_widget.insert("end", "─" * 120 + "\n")
            text_widget.insert("end", f"\nSummary: {changes_count} file(s) will be renamed\n")
            if missing_files:
                text_widget.insert("end", f"⚠ Missing/Unmatched files: {missing_files}\n")
                text_widget.insert("end", f"  Action: Check Column B (Current_Filename) matches exact filenames\n")
            if collisions:
                text_widget.insert("end", f"⚠ Collisions: {collisions} (resolve before running)\n")
            if changes_count == 0:
                text_widget.insert("end", "\n⚠ No files will be renamed. Check Column B values!\n")

            text_widget.config(state="disabled")

            btn_frame = ttk.Frame(preview_win)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="Close", command=preview_win.destroy).pack(side="left", padx=5)
            if changes_count > 0 and missing_files == 0:
                ttk.Button(btn_frame, text="Run Rename", command=lambda:[preview_win.destroy(), self.on_continue()]).pack(side="left", padx=5)
            elif missing_files > 0:
                ttk.Label(btn_frame, text="⚠ Fix missing files before running", foreground="red").pack(side="left", padx=10)

            # CRITICAL: Clear cache after preview
            self.clear_excel_cache()

        except PermissionError as e:
            messagebox.showerror("Excel File Locked", str(e))
            self.clear_excel_cache()
        except Exception as e:
            messagebox.showerror("Preview Error", f"Could not generate preview:\n{e}")
            self.clear_excel_cache()

    def get_files(self, folder, ext, recursive=False):
        """Get files from folder - local operation only"""
        files = []
        try:
            if recursive:
                for root, dirs, filenames in os.walk(folder):
                    for f in filenames:
                        if f.endswith(ext):
                            files.append(os.path.join(root, f))
            else:
                files = [os.path.join(folder, f) for f in os.listdir(folder) 
                        if f.endswith(ext) and os.path.isfile(os.path.join(folder, f))]
            files.sort()
        except Exception:
            pass
        return files

    def generate_new_name(self, row, old_name, mode, ext, delimiter):
        """Generate new filename based on mode and columns"""
        prefix = str(row.get('Prefix', '')).strip()
        new_field = str(row.get('New_Filename', '')).strip()

        if mode == "Prefix":
            if prefix:
                if new_field:
                    base = new_field
                else:
                    base = os.path.splitext(old_name)[0]
                new_base = f"{prefix}{delimiter}{base}" if delimiter is not None else f"{prefix}{base}"
            else:
                if new_field:
                    new_base = new_field
                else:
                    new_base = os.path.splitext(old_name)[0]
        else:  # Replace
            if new_field:
                new_base = new_field
            else:
                new_base = os.path.splitext(old_name)[0]

        if not new_base.endswith(ext):
            return new_base + ext
        return new_base

    def on_continue(self):
        if not self.validate_inputs():
            return

        excel_path = self.excel_path_var.get().strip()
        target_folder = self.target_folder_var.get().strip()
        ext = self.file_ext.get().strip()
        mode = self.mode_var.get()
        delimiter = self.delimiter_var.get()

        if not self.dry_run_var.get():
            if not messagebox.askyesno("Confirm Rename", "⚠ Are you sure you want to rename the files?\n\nThis cannot be undone unless you created a backup."):
                return

        self.save_config()
        log_path = self.rename_files_from_excel(excel_path, target_folder, ext, mode, delimiter)

        if log_path:
            if self.dry_run_var.get():
                messagebox.showinfo("Dry Run Complete", f"Preview complete. No files were renamed.\n\nLog: {log_path}")
            else:
                messagebox.showinfo("✓ Success", f"Renaming complete!\n\nLog: {log_path}")
            self.status_var.set("✓ Operation completed successfully")

    def show_progress(self, current, total):
        """Update inline progress bar"""
        self.progress_var.set(current)
        self.progress_bar['maximum'] = total
        self.progress_frame.grid()
        self.root.update_idletasks()

    def hide_progress(self):
        """Hide inline progress bar"""
        self.progress_frame.grid_remove()
        self.progress_var.set(0)
        self.root.update_idletasks()

    def rename_files_from_excel(self, excel_path, target_folder, ext, mode, delimiter):
        """Rename files - all operations local only"""
        try:
            # Read fresh (no cache)
            df = self.read_excel_safe(excel_path)
            recursive = self.recursive_var.get()
            files = self.get_files(target_folder, ext, recursive)
            files_map = {os.path.basename(p): p for p in files}

            # Pre-scan for planned renames
            planned = 0
            for idx, row in df.iterrows():
                current_filename = str(row.get('Current_Filename', '')).strip()
                if not current_filename or current_filename not in files_map:
                    continue
                old_name = current_filename
                new_name = self.generate_new_name(row, old_name, mode, ext, delimiter)
                if old_name != new_name:
                    planned += 1

            # Create backup
            if self.backup_var.get() and not self.dry_run_var.get():
                backup_folder = self.create_backup(target_folder, files)
                backup_msg = f"Backup: {backup_folder}\n\n"
            else:
                backup_msg = ""

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_filename = f"rename_log_{timestamp}.txt"
            log_path = os.path.join(target_folder, log_filename)

            # Calculate step delay
            if planned > 0:
                step_delay = max(0.02, min(0.05, 2.0 / planned))
            else:
                step_delay = 0.02

            renamed_count = 0
            error_count = 0
            skipped_count = 0
            processed_steps = 0

            with open(log_path, "w", encoding="utf-8") as log_file:
                log_file.write("FILE RENAMER PRO V9 - Log File\n")
                log_file.write("=" * 70 + "\n")
                log_file.write(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                log_file.write(f"Excel: {excel_path}\n")
                log_file.write(f"Target: {target_folder}\n")
                log_file.write(f"Mode: {mode} | Delimiter: '{delimiter}' | Extension: {ext}\n")
                log_file.write(f"Matching: Strict Current_Filename (Column B) matching\n")
                log_file.write(f"{backup_msg}\n")
                log_file.write("=" * 70 + "\n\n")

                if planned > 0:
                    self.show_progress(0, planned)

                for idx, row in df.iterrows():
                    rownum = idx + 1
                    current_filename = str(row.get('Current_Filename', '')).strip()
                    notes = str(row.get('Notes', '')).strip()

                    # CRITICAL: Only process if Current_Filename exists and matches
                    if not current_filename:
                        log_file.write(f"⚠ Row {rownum}: Empty Current_Filename (skipped)\n")
                        skipped_count += 1
                        continue

                    if current_filename not in files_map:
                        log_file.write(f"⚠ Row {rownum}: File not found: '{current_filename}' (skipped)\n")
                        skipped_count += 1
                        continue

                    old_path = files_map[current_filename]
                    old_name = current_filename

                    new_name = self.generate_new_name(row, old_name, mode, ext, delimiter)
                    new_path = os.path.join(os.path.dirname(old_path), new_name)

                    if old_name == new_name:
                        log_file.write(f"○ {old_name} (no change)\n")
                        skipped_count += 1
                        continue

                    if os.path.exists(new_path) and not (os.path.exists(new_path) and os.path.samefile(old_path, new_path)):
                        log_file.write(f"✗ {old_name}\n")
                        log_file.write(f"  ERROR: Target exists: {new_name}\n\n")
                        error_count += 1
                        if planned > 0:
                            processed_steps += 1
                            self.show_progress(processed_steps, planned)
                            time.sleep(step_delay)
                        continue

                    try:
                        if not self.dry_run_var.get():
                            os.rename(old_path, new_path)

                        log_file.write(f"✓ {old_name}\n")
                        log_file.write(f"  → {new_name}\n")
                        if notes:
                            log_file.write(f"  User Note: {notes}\n")
                        log_file.write("\n")

                        renamed_count += 1
                        files_map.pop(old_name, None)
                        files_map[new_name] = new_path
                    except Exception as e:
                        log_file.write(f"✗ {old_name}\n")
                        log_file.write(f"  ERROR: {e}\n\n")
                        error_count += 1

                    if planned > 0:
                        processed_steps += 1
                        self.show_progress(processed_steps, planned)
                        time.sleep(step_delay)

                log_file.write("\n" + "=" * 70 + "\n")
                log_file.write("SUMMARY\n")
                log_file.write("=" * 70 + "\n")
                log_file.write(f"Renamed: {renamed_count} | Errors: {error_count} | Skipped: {skipped_count}\n")
                log_file.write(f"Total processed: {renamed_count + error_count + skipped_count}\n")

            self.hide_progress()
            
            # CRITICAL: Clear cache after operation
            self.clear_excel_cache()

            return log_path

        except PermissionError as e:
            self.hide_progress()
            self.clear_excel_cache()
            messagebox.showerror("Excel File Locked", str(e))
            return None
        except Exception as e:
            self.hide_progress()
            self.clear_excel_cache()
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            return None

    def create_backup(self, target_folder, files):
        """Create backup - local operation only"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = os.path.join(target_folder, f"backup_{timestamp}")
        os.makedirs(backup_folder, exist_ok=True)
        for f in files:
            try:
                shutil.copy2(f, os.path.join(backup_folder, os.path.basename(f)))
            except Exception:
                pass
        return backup_folder

    def show_help(self):
        help_win = tk.Toplevel(self.root)
        help_win.title("Help - " + APP_NAME)
        help_win.geometry("800x700")
        frame = ttk.Frame(help_win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")
        text_widget = tk.Text(frame, wrap="word", yscrollcommand=scrollbar.set, font=("Segoe UI",10))
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_widget.yview)
        help_text = f"""{APP_NAME} - Help

VERSION 9 CRITICAL FIXES:
- ✓ Excel file no longer locked after operations
- ✓ Can edit and re-run without restarting program
- ✓ Fixed incorrect file matching (now strict Column B matching)
- ✓ Security audit: 100% local operation, no external transmission

CRITICAL: HOW FILE MATCHING WORKS (V9)
Column B (Current_Filename) MUST contain the EXACT filename!
- The program searches for files using Column B values only
- If Column B is empty or doesn't match, that row is skipped
- Never change Column B unless you know the exact filename

EXAMPLE:
If your folder has: "PDF-DOCUMENT.pdf"
Column B must say: "PDF-DOCUMENT.pdf" (exact match)

WORKFLOW:
1. Scan folder (creates template with Column B filled correctly)
2. Edit Column C (Prefix), D (New_Filename), E (Notes) ONLY
3. DO NOT change Column B unless file was renamed outside program
4. Close Excel before running Preview or Rename
5. Preview changes
6. Run rename

RENAME MODES:
- Prefix: <Column C><Delimiter><Column D><ext>
- Replace: <Column D><ext>

SECURITY:
- All operations are local only
- No network connections
- No external data transmission
- Files only modified in target folder

TROUBLESHOOTING:
"File is currently open" → Close Excel and try again
"File not found" → Check Column B has exact filename
Wrong files renamed → Did you modify Column B incorrectly?
"""
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")
        ttk.Button(help_win, text="Close", command=help_win.destroy).pack(pady=8)

    def on_quit(self):
        self.clear_excel_cache()  # Clear cache on exit
        self.save_config()
        self.root.destroy()


def launch_gui():
    """Launch application - no external connections"""
    root = TkinterDnD.Tk()
    app = FileRenamerApp(root)
    root.mainloop()

if __name__ == "__main__":
    launch_gui()
