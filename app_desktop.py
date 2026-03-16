import re
import sqlite3
from datetime import datetime
from io import StringIO
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from data_cleaner import DataCleaner

DB_PATH = "data_manager.db"


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS dataset_snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                stage TEXT NOT NULL,
                created_at TEXT NOT NULL,
                rows_count INTEGER NOT NULL,
                cols_count INTEGER NOT NULL,
                data_json TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_name TEXT NOT NULL,
                industry TEXT,
                country TEXT,
                website TEXT,
                email TEXT,
                phone TEXT,
                notes TEXT,
                created_at TEXT NOT NULL
            )
            """
        )


def to_snake_case(name: str) -> str:
    value = re.sub(r"[^0-9a-zA-Z]+", "_", str(name).strip())
    value = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("_").lower() or "column"


def normalize_email(value: object) -> object:
    if pd.isna(value):
        return value
    email = str(value).strip().lower()
    if email == "":
        return None
    if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
        return email
    return email


def normalize_phone(value: object) -> object:
    if pd.isna(value):
        return value
    raw = str(value).strip()
    if raw == "":
        return None
    has_plus = raw.startswith("+")
    digits = re.sub(r"\D", "", raw)
    if digits == "":
        return None
    return f"+{digits}" if has_plus else digits


def clean_dataframe(df: pd.DataFrame, normalize_columns=True, trim_text=True, 
                   drop_duplicates=True, missing_strategy="Keep as is", 
                   normalize_contact_fields=True) -> pd.DataFrame:
    cleaned = df.copy()

    if normalize_columns:
        cleaned.columns = [to_snake_case(col) for col in cleaned.columns]

    if trim_text:
        for col in cleaned.select_dtypes(include=["object", "string"]).columns:
            cleaned[col] = cleaned[col].astype("string").str.strip()

    if drop_duplicates:
        cleaned = cleaned.drop_duplicates().reset_index(drop=True)

    if missing_strategy == "Drop rows with missing":
        cleaned = cleaned.dropna().reset_index(drop=True)
    elif missing_strategy == "Fill missing":
        for col in cleaned.columns:
            if pd.api.types.is_numeric_dtype(cleaned[col]):
                median = cleaned[col].median()
                cleaned[col] = cleaned[col].fillna(median if pd.notna(median) else 0)
            else:
                cleaned[col] = cleaned[col].fillna("Unknown")

    if normalize_contact_fields:
        for col in cleaned.columns:
            lower = col.lower()
            if "email" in lower:
                cleaned[col] = cleaned[col].apply(normalize_email)
            if "phone" in lower or "mobile" in lower or "tel" in lower:
                cleaned[col] = cleaned[col].apply(normalize_phone)

    return cleaned


def filter_dataframe(df: pd.DataFrame, search_text: str = "") -> pd.DataFrame:
    if not search_text.strip():
        return df
    
    mask = df.astype(str).apply(
        lambda x: x.str.contains(search_text, case=False, na=False)
    ).any(axis=1)
    return df[mask].reset_index(drop=True)


def export_to_excel(df: pd.DataFrame, filename: str) -> None:
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        worksheet = writer.sheets["Data"]
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        for column in worksheet.columns:
            column_letter = column[0].column_letter
            worksheet.column_dimensions[column_letter].width = 20


def save_snapshot(name: str, stage: str, df: pd.DataFrame) -> None:
    payload = df.to_json(orient="split", date_format="iso")
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO dataset_snapshots (
                name, stage, created_at, rows_count, cols_count, data_json
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                name,
                stage,
                datetime.now().isoformat(timespec="seconds"),
                int(df.shape[0]),
                int(df.shape[1]),
                payload,
            ),
        )


def list_snapshots() -> list:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, name, stage, created_at, rows_count, cols_count
            FROM dataset_snapshots
            ORDER BY id DESC
            """
        ).fetchall()
    return rows


def load_snapshot(snapshot_id: int) -> pd.DataFrame:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT data_json FROM dataset_snapshots WHERE id = ?",
            (snapshot_id,),
        ).fetchone()
    if row is None:
        return pd.DataFrame()
    return pd.read_json(StringIO(row["data_json"]), orient="split")


def add_company_record(company_name: str, industry: str, country: str, website: str, 
                      email: str, phone: str, notes: str) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO companies (
                company_name, industry, country, website, email, phone, notes, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (company_name, industry, country, website, normalize_email(email),
             normalize_phone(phone), notes, datetime.now().isoformat(timespec="seconds")),
        )


def list_companies() -> pd.DataFrame:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, company_name, industry, country, website, email, phone, notes, created_at
            FROM companies
            ORDER BY id DESC
            """
        ).fetchall()
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame([dict(row) for row in rows])


def delete_company(company_id: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM companies WHERE id = ?", (company_id,))


class DataManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Company Data Manager")
        self.root.geometry("1000x700")
        
        init_db()
        
        self.current_df = pd.DataFrame()
        self.cleaned_df = pd.DataFrame()
        
        self.setup_ui()
    
    def setup_ui(self):
        # Notebook (Tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Tab 1: Data Pipeline
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text="📊 Data Pipeline")
        self.setup_tab1(tab1)
        
        # Tab 2: Company Registry
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text="🏢 Company Registry")
        self.setup_tab2(tab2)
        
        # Tab 3: Snapshots
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text="💾 Snapshots")
        self.setup_tab3(tab3)
        
        # Tab 4: Data Quality & Advanced Cleaning
        tab4 = ttk.Frame(notebook)
        notebook.add(tab4, text="🔍 Data Quality")
        self.setup_tab4(tab4)
    
    def setup_tab1(self, tab):
        # File selection
        frame_file = ttk.LabelFrame(tab, text="Load Dataset", padding=10)
        frame_file.pack(fill="x", padx=5, pady=5)
        
        tk.Button(frame_file, text="Browse CSV/Excel", command=self.load_file).pack(side="left", padx=5)
        self.file_label = tk.Label(frame_file, text="No file selected", fg="gray")
        self.file_label.pack(side="left", padx=5)
        tk.Button(frame_file, text="Load Last Snapshot", command=self.load_last_snapshot).pack(side="left", padx=5)
        
        # Dataset name
        frame_name = ttk.LabelFrame(tab, text="Settings", padding=10)
        frame_name.pack(fill="x", padx=5, pady=5)
        
        tk.Label(frame_name, text="Dataset name:").pack(side="left", padx=5)
        self.dataset_name = tk.StringVar(value="company_dataset")
        tk.Entry(frame_name, textvariable=self.dataset_name, width=30).pack(side="left", padx=5)
        
        # Data info
        self.data_info = tk.Label(tab, text="No data loaded", fg="blue")
        self.data_info.pack(fill="x", padx=10, pady=5)
        
        # Table
        frame_table = ttk.LabelFrame(tab, text="Data Preview (first 50 rows)", padding=5)
        frame_table.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(frame_table, height=10)
        scrollbar = ttk.Scrollbar(frame_table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscroll=scrollbar.set)
        self.tree.pack(fill="both", expand=True)
        scrollbar.pack(fill="x")
        
        # Cleaning options
        frame_clean = ttk.LabelFrame(tab, text="Cleaning Options", padding=10)
        frame_clean.pack(fill="x", padx=5, pady=5)
        
        self.norm_cols = tk.BooleanVar(value=True)
        self.trim_text = tk.BooleanVar(value=True)
        self.drop_dup = tk.BooleanVar(value=True)
        self.norm_contact = tk.BooleanVar(value=True)
        
        tk.Checkbutton(frame_clean, text="Normalize column names", variable=self.norm_cols).pack(side="left")
        tk.Checkbutton(frame_clean, text="Trim text", variable=self.trim_text).pack(side="left")
        tk.Checkbutton(frame_clean, text="Remove duplicates", variable=self.drop_dup).pack(side="left")
        tk.Checkbutton(frame_clean, text="Normalize email/phone", variable=self.norm_contact).pack(side="left")
        
        tk.Label(frame_clean, text="Missing strategy:").pack(side="left", padx=10)
        self.missing_strategy = tk.StringVar(value="Keep as is")
        ttk.Combobox(frame_clean, textvariable=self.missing_strategy, 
                    values=["Keep as is", "Drop rows with missing", "Fill missing"],
                    state="readonly", width=20).pack(side="left")
        
        # Action buttons
        frame_buttons = ttk.Frame(tab)
        frame_buttons.pack(fill="x", padx=5, pady=5)
        
        tk.Button(frame_buttons, text="Run Cleaning", command=self.run_cleaning).pack(side="left", padx=5)
        tk.Button(frame_buttons, text="Search", command=self.search_data).pack(side="left", padx=5)
        tk.Button(frame_buttons, text="Export CSV", command=self.export_csv).pack(side="left", padx=5)
        tk.Button(frame_buttons, text="Export Excel", command=self.export_excel).pack(side="left", padx=5)
        tk.Button(frame_buttons, text="Save Snapshot", command=self.save_snap).pack(side="left", padx=5)
        
        self.status_label = tk.Label(tab, text="", fg="green")
        self.status_label.pack(fill="x", padx=10, pady=5)
    
    def setup_tab2(self, tab):
        # Input form
        frame_form = ttk.LabelFrame(tab, text="Add Company", padding=10)
        frame_form.pack(fill="x", padx=5, pady=5)
        
        form_items = [
            ("Company name*:", "comp_name"),
            ("Industry:", "comp_industry"),
            ("Country:", "comp_country"),
            ("Website:", "comp_website"),
            ("Email:", "comp_email"),
            ("Phone:", "comp_phone"),
        ]
        
        self.comp_vars = {}
        for label, var_name in form_items:
            tk.Label(frame_form, text=label).pack(side="left", padx=5)
            var = tk.StringVar()
            self.comp_vars[var_name] = var
            tk.Entry(frame_form, textvariable=var, width=15).pack(side="left", padx=2)
        
        tk.Button(frame_form, text="Add Company", command=self.add_company).pack(side="left", padx=10)
        tk.Button(frame_form, text="Refresh", command=self.refresh_companies).pack(side="left", padx=5)
        
        # Table
        frame_table = ttk.LabelFrame(tab, text="Companies", padding=5)
        frame_table.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.comp_tree = ttk.Treeview(frame_table, 
                                     columns=("id", "name", "industry", "country", "website", "email", "phone"),
                                     height=10)
        self.comp_tree.column("#0", width=0)
        self.comp_tree.column("id", width=30)
        self.comp_tree.column("name", width=120)
        self.comp_tree.column("industry", width=100)
        self.comp_tree.column("country", width=80)
        self.comp_tree.column("website", width=120)
        self.comp_tree.column("email", width=150)
        self.comp_tree.column("phone", width=100)
        
        self.comp_tree.heading("id", text="ID")
        self.comp_tree.heading("name", text="Company")
        self.comp_tree.heading("industry", text="Industry")
        self.comp_tree.heading("country", text="Country")
        self.comp_tree.heading("website", text="Website")
        self.comp_tree.heading("email", text="Email")
        self.comp_tree.heading("phone", text="Phone")
        
        scrollbar = ttk.Scrollbar(frame_table, orient="horizontal", command=self.comp_tree.xview)
        self.comp_tree.configure(xscroll=scrollbar.set)
        self.comp_tree.pack(fill="both", expand=True)
        scrollbar.pack(fill="x")
        
        # Delete
        frame_del = ttk.Frame(tab)
        frame_del.pack(fill="x", padx=5, pady=5)
        
        tk.Label(frame_del, text="Delete company ID:").pack(side="left", padx=5)
        self.del_id = tk.StringVar()
        tk.Entry(frame_del, textvariable=self.del_id, width=10).pack(side="left", padx=5)
        tk.Button(frame_del, text="Delete", command=self.delete_company).pack(side="left", padx=5)
        
        self.refresh_companies()
    
    def setup_tab3(self, tab):
        tk.Button(tab, text="Refresh Snapshots", command=self.refresh_snapshots).pack(padx=5, pady=5)
        
        frame_table = ttk.LabelFrame(tab, text="Saved Snapshots", padding=5)
        frame_table.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.snap_tree = ttk.Treeview(frame_table, 
                                     columns=("id", "name", "stage", "created", "rows", "cols"),
                                     height=15)
        self.snap_tree.column("#0", width=0)
        self.snap_tree.column("id", width=30)
        self.snap_tree.column("name", width=150)
        self.snap_tree.column("stage", width=80)
        self.snap_tree.column("created", width=100)
        self.snap_tree.column("rows", width=60)
        self.snap_tree.column("cols", width=60)
        
        self.snap_tree.heading("id", text="ID")
        self.snap_tree.heading("name", text="Name")
        self.snap_tree.heading("stage", text="Stage")
        self.snap_tree.heading("created", text="Created")
        self.snap_tree.heading("rows", text="Rows")
        self.snap_tree.heading("cols", text="Cols")
        
        scrollbar = ttk.Scrollbar(frame_table, orient="horizontal", command=self.snap_tree.xview)
        self.snap_tree.configure(xscroll=scrollbar.set)
        self.snap_tree.pack(fill="both", expand=True)
        scrollbar.pack(fill="x")
        
        frame_load = ttk.Frame(tab)
        frame_load.pack(fill="x", padx=5, pady=5)
        
        tk.Label(frame_load, text="Load snapshot ID:").pack(side="left", padx=5)
        self.load_snap_id = tk.StringVar()
        tk.Entry(frame_load, textvariable=self.load_snap_id, width=10).pack(side="left", padx=5)
        tk.Button(frame_load, text="Load", command=self.load_snap).pack(side="left", padx=5)
        
        self.refresh_snapshots()
    
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.xlsx *.xls"), ("All", "*.*")])
        if file_path:
            try:
                if file_path.lower().endswith(".csv"):
                    self.current_df = pd.read_csv(file_path)
                else:
                    self.current_df = pd.read_excel(file_path)
                
                self.file_label.config(text=f"✓ Loaded: {self.current_df.shape[0]} rows × {self.current_df.shape[1]} cols", fg="green")
                self.show_data_in_table(self.current_df)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")
    
    def load_last_snapshot(self):
        snaps = list_snapshots()
        if snaps:
            self.current_df = load_snapshot(snaps[0]["id"])
            self.file_label.config(text=f"✓ Loaded snapshot #{snaps[0]['id']}", fg="green")
            self.show_data_in_table(self.current_df)
        else:
            messagebox.showinfo("Info", "No snapshots found")
    
    def show_data_in_table(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.tree["columns"] = list(df.columns)
        self.tree.column("#0", width=0)
        self.tree.heading("#0", text="")
        
        for col in df.columns:
            self.tree.column(col, width=100)
            self.tree.heading(col, text=col)
        
        for idx, row in df.head(50).iterrows():
            self.tree.insert("", "end", values=list(row))
        
        self.data_info.config(text=f"Data: {df.shape[0]} rows × {df.shape[1]} cols")
    
    def run_cleaning(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        self.cleaned_df = clean_dataframe(
            self.current_df,
            normalize_columns=self.norm_cols.get(),
            trim_text=self.trim_text.get(),
            drop_duplicates=self.drop_dup.get(),
            missing_strategy=self.missing_strategy.get(),
            normalize_contact_fields=self.norm_contact.get()
        )
        
        self.show_data_in_table(self.cleaned_df)
        self.status_label.config(text=f"✓ Cleaned: {self.current_df.shape[0]} → {self.cleaned_df.shape[0]} rows")
    
    def search_data(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        search_text = tk.simpledialog.askstring("Search", "Enter search text:")
        if search_text:
            df = self.cleaned_df if not self.cleaned_df.empty else self.current_df
            filtered = filter_dataframe(df, search_text)
            self.show_data_in_table(filtered)
            self.status_label.config(text=f"✓ Found {filtered.shape[0]} of {df.shape[0]} rows")
    
    def export_csv(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        df = self.cleaned_df if not self.cleaned_df.empty else self.current_df
        filename = f"{self.dataset_name.get()}_export.csv"
        df.to_csv(filename, index=False)
        messagebox.showinfo("Success", f"Exported to {filename}")
        self.status_label.config(text=f"✓ Exported: {filename}")
    
    def export_excel(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        df = self.cleaned_df if not self.cleaned_df.empty else self.current_df
        filename = f"{self.dataset_name.get()}_export.xlsx"
        export_to_excel(df, filename)
        messagebox.showinfo("Success", f"Exported to {filename}")
        self.status_label.config(text=f"✓ Exported: {filename}")
    
    def save_snap(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        df = self.cleaned_df if not self.cleaned_df.empty else self.current_df
        stage = "cleaned" if not self.cleaned_df.empty else "raw"
        save_snapshot(self.dataset_name.get(), stage, df)
        messagebox.showinfo("Success", "Snapshot saved!")
        self.status_label.config(text=f"✓ Saved snapshot ({stage})")
    
    def add_company(self):
        if not self.comp_vars["comp_name"].get().strip():
            messagebox.showwarning("Warning", "Company name is required")
            return
        
        add_company_record(
            self.comp_vars["comp_name"].get().strip(),
            self.comp_vars["comp_industry"].get().strip(),
            self.comp_vars["comp_country"].get().strip(),
            self.comp_vars["comp_website"].get().strip(),
            self.comp_vars["comp_email"].get().strip(),
            self.comp_vars["comp_phone"].get().strip(),
            ""
        )
        
        for var in self.comp_vars.values():
            var.set("")
        
        self.refresh_companies()
        messagebox.showinfo("Success", "Company added!")
    
    def refresh_companies(self):
        for item in self.comp_tree.get_children():
            self.comp_tree.delete(item)
        
        df = list_companies()
        for idx, row in df.iterrows():
            self.comp_tree.insert("", "end", values=(
                row["id"], row["company_name"], row.get("industry", ""), 
                row.get("country", ""), row.get("website", ""), 
                row.get("email", ""), row.get("phone", "")
            ))
    
    def setup_tab4(self, tab):
        """Data Quality & Advanced Cleaning Tab"""
        # Quality report section
        frame_quality = ttk.LabelFrame(tab, text="Data Quality Analysis", padding=10)
        frame_quality.pack(fill="x", padx=5, pady=5)
        
        tk.Button(frame_quality, text="Generate Quality Report", command=self.generate_quality_report).pack(side="left", padx=5)
        tk.Button(frame_quality, text="Detect Outliers (IQR)", command=self.detect_outliers_iqr).pack(side="left", padx=5)
        tk.Button(frame_quality, text="Detect Outliers (Z-Score)", command=self.detect_outliers_zscore).pack(side="left", padx=5)
        
        # Advanced cleaning options
        frame_advanced = ttk.LabelFrame(tab, text="Advanced Cleaning Options", padding=10)
        frame_advanced.pack(fill="x", padx=5, pady=5)
        
        self.clean_remove_dup = tk.BooleanVar(value=True)
        self.clean_whitespace = tk.BooleanVar(value=True)
        self.clean_standardize = tk.BooleanVar(value=True)
        
        tk.Checkbutton(frame_advanced, text="Remove duplicates", variable=self.clean_remove_dup).pack(side="left")
        tk.Checkbutton(frame_advanced, text="Clean whitespace", variable=self.clean_whitespace).pack(side="left")
        tk.Checkbutton(frame_advanced, text="Standardize column names", variable=self.clean_standardize).pack(side="left")
        
        tk.Label(frame_advanced, text="Handle missing:").pack(side="left", padx=10)
        self.clean_missing_strategy = tk.StringVar(value="fill")
        ttk.Combobox(frame_advanced, textvariable=self.clean_missing_strategy,
                    values=["report", "fill", "drop", "drop_cols"],
                    state="readonly", width=15).pack(side="left")
        
        tk.Button(frame_advanced, text="Run Full Cleaning", command=self.run_full_cleaning).pack(side="left", padx=10)
        
        # Results text area
        frame_results = ttk.LabelFrame(tab, text="Cleaning Results", padding=5)
        frame_results.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.results_text = tk.Text(frame_results, height=15, width=80)
        scrollbar = ttk.Scrollbar(frame_results, orient="vertical", command=self.results_text.yview)
        
        self.results_text.configure(yscroll=scrollbar.set)
        self.results_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.quality_status = tk.Label(tab, text="Ready for analysis", fg="blue")
        self.quality_status.pack(fill="x", padx=10, pady=5)
    
    def generate_quality_report(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        try:
            cleaner = DataCleaner(self.current_df if self.cleaned_df.empty else self.cleaned_df)
            report = cleaner.generate_quality_report()
            
            self.results_text.delete("1.0", "end")
            self.results_text.insert("end", "=== DATA QUALITY REPORT ===\n\n")
            
            for col, info in report.items():
                self.results_text.insert("end", f"{col}:\n")
                self.results_text.insert("end", f"  Type: {info['dtype']}\n")
                self.results_text.insert("end", f"  Non-null: {info['non_null']}/{len(self.current_df)}\n")
                self.results_text.insert("end", f"  Missing: {info['null_count']} ({info['null_percent']}%)\n")
                self.results_text.insert("end", f"  Unique values: {info['unique_values']}\n")
                self.results_text.insert("end", f"  Duplicates: {info['duplicates']}\n\n")
            
            self.quality_status.config(text="✓ Quality report generated", fg="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {e}")
    
    def detect_outliers_iqr(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        try:
            cleaner = DataCleaner(self.current_df if self.cleaned_df.empty else self.cleaned_df)
            report = cleaner.detect_outliers(method="iqr")
            
            self.results_text.delete("1.0", "end")
            self.results_text.insert("end", "=== OUTLIER DETECTION (IQR METHOD) ===\n\n")
            
            if not report:
                self.results_text.insert("end", "No numeric columns found or no outliers detected.")
            else:
                for col, info in report.items():
                    self.results_text.insert("end", f"{col}:\n")
                    self.results_text.insert("end", f"  Method: {info['method']}\n")
                    self.results_text.insert("end", f"  Outliers found: {info['outlier_count']} ({info['outlier_percent']}%)\n")
                    self.results_text.insert("end", f"  Bounds: [{info['bounds']['lower']}, {info['bounds']['upper']}]\n")
                    self.results_text.insert("end", f"  Sample values: {info['values']}\n\n")
            
            self.quality_status.config(text="✓ Outliers detected (IQR)", fg="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to detect outliers: {e}")
    
    def detect_outliers_zscore(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        try:
            cleaner = DataCleaner(self.current_df if self.cleaned_df.empty else self.cleaned_df)
            report = cleaner.detect_outliers(method="zscore")
            
            self.results_text.delete("1.0", "end")
            self.results_text.insert("end", "=== OUTLIER DETECTION (Z-SCORE METHOD) ===\n\n")
            
            if not report:
                self.results_text.insert("end", "No numeric columns found or no outliers detected.")
            else:
                for col, info in report.items():
                    self.results_text.insert("end", f"{col}:\n")
                    self.results_text.insert("end", f"  Method: {info['method']}\n")
                    self.results_text.insert("end", f"  Outliers found: {info['outlier_count']} ({info['outlier_percent']}%)\n")
                    self.results_text.insert("end", f"  Threshold (Z > {info['threshold']})\n")
                    self.results_text.insert("end", f"  Sample values: {info['values']}\n\n")
            
            self.quality_status.config(text="✓ Outliers detected (Z-Score)", fg="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to detect outliers: {e}")
    
    def run_full_cleaning(self):
        if self.current_df.empty:
            messagebox.showwarning("Warning", "Please load data first")
            return
        
        try:
            cleaner = DataCleaner(self.current_df if self.cleaned_df.empty else self.cleaned_df)
            
            all_reports = cleaner.run_full_cleaning(
                remove_duplicates=self.clean_remove_dup.get(),
                handle_missing=self.clean_missing_strategy.get(),
                remove_whitespace=self.clean_whitespace.get(),
                standardize_names=self.clean_standardize.get()
            )
            
            self.cleaned_df = cleaner.get_cleaned_data()
            
            self.results_text.delete("1.0", "end")
            self.results_text.insert("end", "=== FULL CLEANING PIPELINE REPORT ===\n\n")
            
            # Standardize names
            if "standardize_names" in all_reports:
                self.results_text.insert("end", "Standardize Names:\n")
                report = all_reports["standardize_names"]
                self.results_text.insert("end", f"  Total renamed: {report['total_renamed']}\n")
                self.results_text.insert("end", f"  Mappings: {report['renamed_columns']}\n\n")
            
            # Whitespace
            if "whitespace" in all_reports:
                self.results_text.insert("end", "Remove Whitespace:\n")
                for col, info in all_reports["whitespace"].items():
                    self.results_text.insert("end", f"  {col}: {info['chars_removed']} chars removed\n")
                self.results_text.insert("end", "\n")
            
            # Duplicates
            if "duplicates" in all_reports:
                self.results_text.insert("end", "Remove Duplicates:\n")
                report = all_reports["duplicates"]
                self.results_text.insert("end", f"  Rows removed: {report['removed_rows']}\n")
                self.results_text.insert("end", f"  Rows remaining: {report['remaining_rows']}\n")
                self.results_text.insert("end", f"  Reduction: {report['percent_reduced']}%\n\n")
            
            # Missing values
            if "missing_values" in all_reports:
                self.results_text.insert("end", "Handle Missing Values:\n")
                report = all_reports["missing_values"]
                self.results_text.insert("end", f"  {report}\n\n")
            
            # Before/After
            if "before_after" in all_reports:
                self.results_text.insert("end", "Before/After Comparison:\n")
                report = all_reports["before_after"]
                self.results_text.insert("end", f"  Shape: {report['original_shape']} → {report['cleaned_shape']}\n")
                self.results_text.insert("end", f"  Rows removed: {report['rows_removed']}\n")
                self.results_text.insert("end", f"  Missing values reduced: {report['missing_reduced']}\n")
            
            self.show_data_in_table(self.cleaned_df)
            self.quality_status.config(text=f"✓ Full cleaning completed | {self.current_df.shape[0]} → {self.cleaned_df.shape[0]} rows", fg="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run cleaning: {e}")

    def delete_company(self):
        try:
            company_id = int(self.del_id.get())
            delete_company(company_id)
            self.del_id.set("")
            self.refresh_companies()
            messagebox.showinfo("Success", f"Company #{company_id} deleted!")
        except ValueError:
            messagebox.showwarning("Warning", "Invalid company ID")

    def refresh_snapshots(self):
        for item in self.snap_tree.get_children():
            self.snap_tree.delete(item)

        snaps = list_snapshots()
        for snap in snaps:
            self.snap_tree.insert("", "end", values=(
                snap["id"], snap["name"], snap["stage"],
                snap["created_at"][:10], snap["rows_count"], snap["cols_count"] 
            ))

    def load_snap(self):
        try:
            snap_id = int(self.load_snap_id.get())
            self.current_df = load_snapshot(snap_id)
            self.show_data_in_table(self.current_df)
            self.load_snap_id.set("")
            messagebox.showinfo("Success", f"Loaded snapshot #{snap_id}")       
        except ValueError:
            messagebox.showwarning("Warning", "Invalid snapshot ID")


if __name__ == "__main__":
    import tkinter.simpledialog

    root = tk.Tk()
    app = DataManagerApp(root)
    root.mainloop()
