import re
import sqlite3
from datetime import datetime
from io import StringIO, BytesIO
import os
import threading

import pandas as pd
import PySimpleGUI as sg
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

DB_PATH = "data_manager.db"
sg.theme("DarkBlue3")


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


def clean_dataframe(
    df: pd.DataFrame,
    normalize_columns: bool,
    trim_text: bool,
    drop_duplicates: bool,
    missing_strategy: str,
    normalize_contact_fields: bool,
) -> pd.DataFrame:
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


def filter_dataframe(df: pd.DataFrame, search_text: str = "", filters: dict = None) -> pd.DataFrame:
    filtered = df.copy()
    
    if search_text.strip():
        mask = filtered.astype(str).apply(
            lambda x: x.str.contains(search_text, case=False, na=False)
        ).any(axis=1)
        filtered = filtered[mask]
    
    if filters:
        for col, criteria in filters.items():
            if col not in filtered.columns:
                continue
            
            if isinstance(criteria, dict) and "min" in criteria and "max" in criteria:
                filtered = filtered[
                    (filtered[col] >= criteria["min"]) & 
                    (filtered[col] <= criteria["max"])
                ]
            elif isinstance(criteria, list) and criteria:
                filtered = filtered[filtered[col].isin(criteria)]
    
    return filtered.reset_index(drop=True)


def export_to_excel(df: pd.DataFrame, filename: str) -> None:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
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
    
    output.seek(0)
    with open(filename, "wb") as f:
        f.write(output.getvalue())


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


def list_snapshots() -> list[sqlite3.Row]:
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


def add_company_record(
    company_name: str,
    industry: str,
    country: str,
    website: str,
    email: str,
    phone: str,
    notes: str,
) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO companies (
                company_name, industry, country, website, email, phone, notes, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                company_name,
                industry,
                country,
                website,
                normalize_email(email),
                normalize_phone(phone),
                notes,
                datetime.now().isoformat(timespec="seconds"),
            ),
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
        return pd.DataFrame(
            columns=[
                "id",
                "company_name",
                "industry",
                "country",
                "website",
                "email",
                "phone",
                "notes",
                "created_at",
            ]
        )
    return pd.DataFrame([dict(row) for row in rows])


def delete_company(company_id: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM companies WHERE id = ?", (company_id,))


def dataframe_to_table_data(df: pd.DataFrame, max_rows: int = 50) -> tuple:
    """Convert dataframe to PySimpleGUI table format"""
    display_df = df.head(max_rows)
    data = [list(display_df.columns)] + [list(row) for row in display_df.values]
    return data


def create_data_pipeline_tab():
    return [
        [sg.Text("📊 DATA PIPELINE", font=("Arial", 14, "bold"))],
        [sg.Text("Load Dataset:")],
        [
            sg.FileBrowse("Choose CSV/Excel", file_types=(("Data Files", "*.csv *.xlsx *.xls"),)),
            sg.Input(key="-FILE_PATH-", size=(30, 1), disabled=True),
        ],
        [sg.Text("Dataset name:"), sg.Input("company_dataset", key="-DATASET_NAME-", size=(30, 1))],
        [sg.Button("Load File"), sg.Button("Load Last Snapshot")],
        [sg.Multiline(size=(100, 3), key="-DATA_INFO-", disabled=True)],
        [sg.Text("Current data (first 50 rows):")],
        [sg.Table(values=[[]], headings=[], max_col_width=20, auto_size_columns=False,
                  justification="left", num_rows=10, key="-DATA_TABLE-", size=(100, 10))],
        
        [sg.Text("Cleaning Options:", font=("Arial", 12, "bold"))],
        [
            sg.Checkbox("Normalize column names", default=True, key="-NORM_COLS-"),
            sg.Checkbox("Trim text", default=True, key="-TRIM-"),
            sg.Checkbox("Remove duplicates", default=True, key="-DUP-"),
        ],
        [
            sg.Checkbox("Normalize email/phone", default=True, key="-NORM_CONTACT-"),
            sg.Text("Missing strategy:"),
            sg.Combo(["Keep as is", "Drop rows with missing", "Fill missing"],
                    default_value="Keep as is", key="-MISSING_STRATEGY-"),
        ],
        [sg.Button("Run Cleaning")],
        
        [sg.Text("Search & Filter:", font=("Arial", 12, "bold"))],
        [sg.Text("Search text:"), sg.Input("", key="-SEARCH_TEXT-", size=(30, 1)), sg.Button("Search")],
        
        [sg.Text("Export Options:", font=("Arial", 12, "bold"))],
        [
            sg.Button("Export CSV"), 
            sg.Button("Export Excel"),
            sg.Button("Save Snapshot"),
        ],
        [sg.Text("", key="-STATUS_MSG-", text_color="lightblue")],
    ]


def create_company_registry_tab():
    return [
        [sg.Text("🏢 COMPANY REGISTRY", font=("Arial", 14, "bold"))],
        [sg.Text("Add New Company:")],
        [
            [sg.Text("Company name*:"), sg.Input(key="-COMP_NAME-", size=(30, 1))],
            [sg.Text("Industry:"), sg.Input(key="-COMP_INDUSTRY-", size=(30, 1))],
            [sg.Text("Country:"), sg.Input(key="-COMP_COUNTRY-", size=(30, 1))],
            [sg.Text("Website:"), sg.Input(key="-COMP_WEBSITE-", size=(30, 1))],
            [sg.Text("Email:"), sg.Input(key="-COMP_EMAIL-", size=(30, 1))],
            [sg.Text("Phone:"), sg.Input(key="-COMP_PHONE-", size=(30, 1))],
            [sg.Text("Notes:"), sg.Multiline(size=(30, 3), key="-COMP_NOTES-")],
            [sg.Button("Add Company"), sg.Button("Refresh List")],
        ],
        [sg.Text("Companies in database:")],
        [sg.Table(values=[[]], headings=[], max_col_width=20, auto_size_columns=False,
                  justification="left", num_rows=10, key="-COMPANIES_TABLE-", size=(100, 10))],
        [sg.Text("Delete company ID:"), sg.Input("", key="-DEL_COMPANY_ID-", size=(10, 1)),
         sg.Button("Delete")],
        [sg.Text("", key="-COMPANY_MSG-", text_color="lightgreen")],
    ]


def create_snapshots_tab():
    return [
        [sg.Text("💾 SAVED SNAPSHOTS", font=("Arial", 14, "bold"))],
        [sg.Button("Refresh Snapshots")],
        [sg.Table(values=[[]], headings=[], max_col_width=20, auto_size_columns=False,
                  justification="left", num_rows=10, key="-SNAPSHOTS_TABLE-", size=(100, 10))],
        [sg.Text("Load snapshot ID:"), sg.Input("", key="-LOAD_SNAPSHOT_ID-", size=(10, 1)),
         sg.Button("Load Snapshot")],
        [sg.Text("", key="-SNAPSHOT_MSG-", text_color="lightyellow")],
    ]


def main():
    init_db()
    
    layout = [
        [sg.Text("Company Data Manager", font=("Arial", 16, "bold"), justification="center")],
        [sg.TabGroup([
            [sg.Tab("Data Pipeline", create_data_pipeline_tab(), key="-TAB1-")],
            [sg.Tab("Company Registry", create_company_registry_tab(), key="-TAB2-")],
            [sg.Tab("Snapshots", create_snapshots_tab(), key="-TAB3-")],
        ])],
        [sg.Button("Exit", size=(10, 1))],
    ]
    
    window = sg.Window("Company Data Manager", layout, finalize=True, size=(1200, 800))
    
    current_df = pd.DataFrame()
    cleaned_df = pd.DataFrame()
    filtered_df = pd.DataFrame()
    
    while True:
        event, values = window.read()
        
        if event == sg.WINDOW_CLOSED or event == "Exit":
            break
        
        # Data Pipeline Tab
        if event == "Load File":
            file_path = values["-FILE_PATH-"]
            if file_path:
                try:
                    if file_path.lower().endswith(".csv"):
                        current_df = pd.read_csv(file_path)
                    else:
                        current_df = pd.read_excel(file_path)
                    
                    table_data = dataframe_to_table_data(current_df)
                    window["-DATA_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                    window["-DATA_INFO-"].update(f"✓ Loaded: {current_df.shape[0]} rows × {current_df.shape[1]} cols")
                except Exception as e:
                    sg.popup_error(f"Error loading file: {e}")
        
        elif event == "Load Last Snapshot":
            snapshots = list_snapshots()
            if snapshots:
                last = snapshots[0]
                current_df = load_snapshot(last["id"])
                table_data = dataframe_to_table_data(current_df)
                window["-DATA_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                window["-DATA_INFO-"].update(f"✓ Loaded snapshot #{last['id']}: {current_df.shape[0]} rows × {current_df.shape[1]} cols")
        
        elif event == "Run Cleaning":
            if current_df.empty:
                sg.popup_warning("Please load a dataset first!")
            else:
                try:
                    cleaned_df = clean_dataframe(
                        current_df,
                        normalize_columns=values["-NORM_COLS-"],
                        trim_text=values["-TRIM-"],
                        drop_duplicates=values["-DUP-"],
                        missing_strategy=values["-MISSING_STRATEGY-"],
                        normalize_contact_fields=values["-NORM_CONTACT-"],
                    )
                    
                    table_data = dataframe_to_table_data(cleaned_df)
                    window["-DATA_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                    window["-DATA_INFO-"].update(
                        f"Original: {current_df.shape[0]} rows | Cleaned: {cleaned_df.shape[0]} rows"
                    )
                    window["-STATUS_MSG-"].update("✓ Data cleaned successfully!")
                except Exception as e:
                    sg.popup_error(f"Error cleaning data: {e}")
        
        elif event == "Search":
            if cleaned_df.empty and current_df.empty:
                sg.popup_warning("Please load and clean data first!")
            else:
                df = cleaned_df if not cleaned_df.empty else current_df
                search_text = values["-SEARCH_TEXT-"]
                filtered_df = filter_dataframe(df, search_text)
                
                table_data = dataframe_to_table_data(filtered_df)
                window["-DATA_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                window["-DATA_INFO-"].update(f"✓ Filtered: {filtered_df.shape[0]} of {df.shape[0]} rows")
                window["-STATUS_MSG-"].update(f"Search found {filtered_df.shape[0]} rows")
        
        elif event == "Export CSV":
            if current_df.empty:
                sg.popup_warning("Please load data first!")
            else:
                df = cleaned_df if not cleaned_df.empty else current_df
                dataset_name = values["-DATASET_NAME-"]
                filepath = f"{dataset_name}_export.csv"
                df.to_csv(filepath, index=False)
                sg.popup_ok(f"✓ Exported to {filepath}")
                window["-STATUS_MSG-"].update(f"✓ Exported CSV: {filepath}")
        
        elif event == "Export Excel":
            if current_df.empty:
                sg.popup_warning("Please load data first!")
            else:
                df = cleaned_df if not cleaned_df.empty else current_df
                dataset_name = values["-DATASET_NAME-"]
                filepath = f"{dataset_name}_export.xlsx"
                export_to_excel(df, filepath)
                sg.popup_ok(f"✓ Exported to {filepath}")
                window["-STATUS_MSG-"].update(f"✓ Exported Excel: {filepath}")
        
        elif event == "Save Snapshot":
            if current_df.empty:
                sg.popup_warning("Please load data first!")
            else:
                df = cleaned_df if not cleaned_df.empty else current_df
                dataset_name = values["-DATASET_NAME-"]
                stage = "cleaned" if not cleaned_df.empty else "raw"
                save_snapshot(dataset_name, stage, df)
                sg.popup_ok(f"✓ Snapshot saved!")
                window["-STATUS_MSG-"].update(f"✓ Snapshot saved: {dataset_name} ({stage})")
        
        # Company Registry Tab
        elif event == "Add Company":
            if not values["-COMP_NAME-"].strip():
                sg.popup_warning("Company name is required!")
            else:
                try:
                    add_company_record(
                        values["-COMP_NAME-"].strip(),
                        values["-COMP_INDUSTRY-"].strip(),
                        values["-COMP_COUNTRY-"].strip(),
                        values["-COMP_WEBSITE-"].strip(),
                        values["-COMP_EMAIL-"].strip(),
                        values["-COMP_PHONE-"].strip(),
                        values["-COMP_NOTES-"].strip(),
                    )
                    window["-COMP_NAME-"].update("")
                    window["-COMP_INDUSTRY-"].update("")
                    window["-COMP_COUNTRY-"].update("")
                    window["-COMP_WEBSITE-"].update("")
                    window["-COMP_EMAIL-"].update("")
                    window["-COMP_PHONE-"].update("")
                    window["-COMP_NOTES-"].update("")
                    
                    companies_df = list_companies()
                    table_data = dataframe_to_table_data(companies_df)
                    window["-COMPANIES_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                    window["-COMPANY_MSG-"].update("✓ Company added successfully!")
                except Exception as e:
                    sg.popup_error(f"Error: {e}")
        
        elif event == "Refresh List":
            companies_df = list_companies()
            table_data = dataframe_to_table_data(companies_df)
            window["-COMPANIES_TABLE-"].update(values=table_data[1:], headings=table_data[0])
        
        elif event == "Delete":
            try:
                comp_id = int(values["-DEL_COMPANY_ID-"])
                delete_company(comp_id)
                
                companies_df = list_companies()
                table_data = dataframe_to_table_data(companies_df)
                window["-COMPANIES_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                window["-DEL_COMPANY_ID-"].update("")
                window["-COMPANY_MSG-"].update(f"✓ Company {comp_id} deleted!")
            except ValueError:
                sg.popup_warning("Please enter a valid company ID")
        
        # Snapshots Tab
        elif event == "Refresh Snapshots":
            snapshots = list_snapshots()
            table_data = [["ID", "Name", "Stage", "Created", "Rows", "Cols"]]
            for snap in snapshots:
                table_data.append([
                    snap["id"], snap["name"], snap["stage"],
                    snap["created_at"][:10], snap["rows_count"], snap["cols_count"]
                ])
            window["-SNAPSHOTS_TABLE-"].update(values=table_data[1:], headings=table_data[0])
        
        elif event == "Load Snapshot":
            try:
                snap_id = int(values["-LOAD_SNAPSHOT_ID-"])
                current_df = load_snapshot(snap_id)
                table_data = dataframe_to_table_data(current_df)
                window["-DATA_TABLE-"].update(values=table_data[1:], headings=table_data[0])
                window["-DATA_INFO-"].update(f"✓ Loaded snapshot #{snap_id}: {current_df.shape[0]} rows × {current_df.shape[1]} cols")
                window["-SNAPSHOT_MSG-"].update(f"✓ Loaded snapshot #{snap_id}")
            except ValueError:
                sg.popup_warning("Please enter a valid snapshot ID")
    
    window.close()


if __name__ == "__main__":
    main()
