import re
import sqlite3
from datetime import datetime
from io import StringIO, BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

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
    """Filter dataframe by text search and column-specific criteria"""
    filtered = df.copy()
    
    # Text search across all columns
    if search_text.strip():
        mask = filtered.astype(str).apply(
            lambda x: x.str.contains(search_text, case=False, na=False)
        ).any(axis=1)
        filtered = filtered[mask]
    
    # Column-specific filters
    if filters:
        for col, criteria in filters.items():
            if col not in filtered.columns:
                continue
            
            col_type = filtered[col].dtype
            
            if isinstance(criteria, dict) and "min" in criteria and "max" in criteria:
                # Numeric range filter
                filtered = filtered[
                    (filtered[col] >= criteria["min"]) & 
                    (filtered[col] <= criteria["max"])
                ]
            elif isinstance(criteria, list) and criteria:
                # Categorical filter (selected values)
                filtered = filtered[filtered[col].isin(criteria)]
    
    return filtered.reset_index(drop=True)


def export_to_excel(df: pd.DataFrame, filename: str) -> BytesIO:
    """Export dataframe to Excel with formatting"""
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
            max_length = 30
            column_letter = column[0].column_letter
            worksheet.column_dimensions[column_letter].width = min(max_length, 50)
    
    output.seek(0)
    return output



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


def render_visualizations(df: pd.DataFrame) -> None:
    st.subheader("Visualization")

    if df.empty:
        st.info("No data to visualize.")
        return

    missing_pct = (df.isna().mean() * 100).sort_values(ascending=False)
    if (missing_pct > 0).any():
        miss_df = missing_pct.reset_index()
        miss_df.columns = ["column", "missing_percent"]
        fig = px.bar(
            miss_df,
            x="column",
            y="missing_percent",
            title="Missing Data Percentage by Column",
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.success("No missing values detected.")

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        selected_num = st.selectbox("Numeric column", numeric_cols, key="num_col")
        num_fig = px.histogram(df, x=selected_num, nbins=30, title=f"Distribution: {selected_num}")
        st.plotly_chart(num_fig, use_container_width=True)

    categorical_cols = df.select_dtypes(include=["object", "string", "category"]).columns.tolist()
    if categorical_cols:
        selected_cat = st.selectbox("Category column", categorical_cols, key="cat_col")
        top_df = (
            df[selected_cat]
            .astype("string")
            .fillna("Unknown")
            .value_counts()
            .head(15)
            .reset_index()
        )
        top_df.columns = [selected_cat, "count"]
        cat_fig = px.bar(top_df, x=selected_cat, y="count", title=f"Top values: {selected_cat}")
        st.plotly_chart(cat_fig, use_container_width=True)


def main() -> None:
    st.set_page_config(page_title="Company Data Manager", layout="wide")
    init_db()

    st.title("Company Data Manager")
    st.caption("Manage, clean, and visualize company data with local SQLite storage")

    tab1, tab2, tab3 = st.tabs(["Data Pipeline", "Company Registry", "Saved Snapshots"])

    with tab1:
        st.subheader("1) Load Dataset")
        uploaded = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx", "xls"])

        current_df = pd.DataFrame()
        dataset_name = st.text_input("Dataset name", value="company_dataset")

        if uploaded is not None:
            if uploaded.name.lower().endswith(".csv"):
                current_df = pd.read_csv(uploaded)
            else:
                current_df = pd.read_excel(uploaded)

        col_load_a, col_load_b = st.columns([2, 1])
        with col_load_a:
            snapshots = list_snapshots()
            snapshot_options = {
                f"#{row['id']} | {row['name']} | {row['stage']} | {row['rows_count']}x{row['cols_count']}": row["id"]
                for row in snapshots
            }
            selected_snapshot = st.selectbox(
                "Or load from saved snapshot",
                options=["None"] + list(snapshot_options.keys()),
            )
        with col_load_b:
            if st.button("Load selected snapshot") and selected_snapshot != "None":
                current_df = load_snapshot(snapshot_options[selected_snapshot])

        if not current_df.empty:
            st.write(f"Rows: {current_df.shape[0]} | Columns: {current_df.shape[1]}")
            st.dataframe(current_df.head(30), use_container_width=True)

            if st.button("Save as raw snapshot"):
                save_snapshot(dataset_name, "raw", current_df)
                st.success("Raw snapshot saved.")

            st.subheader("2) Cleaning Settings")
            col1, col2, col3 = st.columns(3)
            with col1:
                normalize_columns = st.checkbox("Normalize column names", value=True)
                trim_text = st.checkbox("Trim spaces in text", value=True)
            with col2:
                drop_duplicates = st.checkbox("Remove duplicate rows", value=True)
                normalize_contacts = st.checkbox("Normalize email and phone", value=True)
            with col3:
                missing_strategy = st.selectbox(
                    "Missing value strategy",
                    ["Keep as is", "Drop rows with missing", "Fill missing"],
                )

            if st.button("Run cleaning"):
                cleaned = clean_dataframe(
                    current_df,
                    normalize_columns=normalize_columns,
                    trim_text=trim_text,
                    drop_duplicates=drop_duplicates,
                    missing_strategy=missing_strategy,
                    normalize_contact_fields=normalize_contacts,
                )
                st.session_state["cleaned_df"] = cleaned

            if "cleaned_df" in st.session_state:
                cleaned_df = st.session_state["cleaned_df"]
                st.subheader("3) Cleaned Result")
                metric_a, metric_b, metric_c = st.columns(3)
                metric_a.metric("Original rows", current_df.shape[0])
                metric_b.metric("Cleaned rows", cleaned_df.shape[0])
                metric_c.metric("Columns", cleaned_df.shape[1])

                st.subheader("4) Search & Filter")
                with st.expander("🔍 Search and Filter Options", expanded=False):
                    search_col, filter_col = st.columns(2)
                    
                    with search_col:
                        search_text = st.text_input("Search across all columns", "")
                    
                    with filter_col:
                        st.write("")  # Spacing
                    
                    # Column-specific filters
                    filters = {}
                    numeric_cols = cleaned_df.select_dtypes(include="number").columns.tolist()
                    categorical_cols = cleaned_df.select_dtypes(include=["object", "string"]).columns.tolist()
                    
                    if numeric_cols:
                        st.write("**Numeric Filters:**")
                        for col in numeric_cols[:3]:  # Limit to first 3
                            min_val, max_val = st.slider(
                                f"{col}",
                                min_value=float(cleaned_df[col].min()),
                                max_value=float(cleaned_df[col].max()),
                                value=(float(cleaned_df[col].min()), float(cleaned_df[col].max())),
                                key=f"slider_{col}"
                            )
                            filters[col] = {"min": min_val, "max": max_val}
                    
                    if categorical_cols:
                        st.write("**Category Filters:**")
                        for col in categorical_cols[:2]:  # Limit to first 2
                            unique_vals = cleaned_df[col].fillna("Unknown").unique().tolist()
                            selected_vals = st.multiselect(
                                f"{col}",
                                options=unique_vals,
                                default=unique_vals,
                                key=f"multiselect_{col}"
                            )
                            if selected_vals:
                                filters[col] = selected_vals
                    
                    apply_filter = st.button("Apply Filters", key="apply_filters")
                
                # Apply filters
                if apply_filter or search_text:
                    filtered_df = filter_dataframe(cleaned_df, search_text, filters)
                    st.session_state["filtered_df"] = filtered_df
                    st.info(f"Filtered: {filtered_df.shape[0]} of {cleaned_df.shape[0]} rows")
                else:
                    st.session_state["filtered_df"] = cleaned_df
                
                display_df = st.session_state["filtered_df"]
                st.dataframe(display_df.head(50), use_container_width=True)

                st.subheader("5) Export Data")
                export_col1, export_col2, export_col3 = st.columns(3)
                
                with export_col1:
                    csv_data = display_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "📥 Download CSV",
                        data=csv_data,
                        file_name=f"{dataset_name}_export.csv",
                        mime="text/csv",
                    )
                
                with export_col2:
                    excel_file = export_to_excel(display_df, f"{dataset_name}_export.xlsx")
                    st.download_button(
                        "📊 Download Excel",
                        data=excel_file,
                        file_name=f"{dataset_name}_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                
                with export_col3:
                    cleaned_csv = cleaned_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "💾 Save as Raw CSV",
                        data=cleaned_csv,
                        file_name=f"{dataset_name}_cleaned.csv",
                        mime="text/csv",
                    )

                col_save, _ = st.columns([1, 4])
                with col_save:
                    if st.button("Save cleaned snapshot"):
                        save_snapshot(dataset_name, "cleaned", cleaned_df)
                        st.success("Cleaned snapshot saved.")

                render_visualizations(cleaned_df)
        else:
            st.info("Upload a dataset or load a snapshot to start.")

    with tab2:
        st.subheader("Company Registry")
        st.caption("Quickly store core company details as structured records")

        with st.form("company_form", clear_on_submit=True):
            company_name = st.text_input("Company name*")
            industry = st.text_input("Industry")
            country = st.text_input("Country")
            website = st.text_input("Website")
            email = st.text_input("Email")
            phone = st.text_input("Phone")
            notes = st.text_area("Notes")
            submitted = st.form_submit_button("Add company")

            if submitted:
                if not company_name.strip():
                    st.error("Company name is required.")
                else:
                    add_company_record(company_name.strip(), industry.strip(), country.strip(), website.strip(), email.strip(), phone.strip(), notes.strip())
                    st.success("Company saved.")

        company_df = list_companies()
        st.dataframe(company_df, use_container_width=True)

        if not company_df.empty:
            del_id = st.selectbox("Select company ID to delete", company_df["id"].tolist())
            if st.button("Delete selected company"):
                delete_company(int(del_id))
                st.warning("Company deleted. Refresh by switching tabs or rerunning.")

    with tab3:
        st.subheader("Saved Snapshots")
        snapshot_rows = list_snapshots()
        if not snapshot_rows:
            st.info("No snapshots saved yet.")
        else:
            summary = pd.DataFrame([dict(row) for row in snapshot_rows])
            st.dataframe(summary, use_container_width=True)


if __name__ == "__main__":
    main()
