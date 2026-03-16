# Company Data Manager

A local data management app for company datasets with:
- dataset upload (CSV/Excel)
- data cleaning pipeline
- visual analytics
- local snapshot storage (SQLite)
- basic company registry

## Features

- Load data from CSV/XLSX
- Save snapshots (`raw` and `cleaned`) to local SQLite
- Cleaning options:
  - normalize column names
  - trim text fields
  - remove duplicates
  - handle missing values (keep/drop/fill)
  - normalize email and phone values
- Visualizations:
  - missing data percentage by column
  - numeric distribution histogram
  - top values for categorical columns
- Company registry CRUD (create/list/delete)

## Project Structure

- `app.py`: Streamlit application
- `requirements.txt`: Python dependencies
- `data_manager.db`: SQLite database (created at runtime)

## Run Locally

1. Create virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies:

```powershell
pip install -r requirements.txt
```

3. Start app:

```powershell
streamlit run app.py
```

The app will open in your browser at a local URL.

## Notes

- This app stores all data locally in SQLite.
- Snapshots are stored in `dataset_snapshots` table.
- Company records are stored in `companies` table.
