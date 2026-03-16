# Company Data Manager

A comprehensive data management application with multiple interfaces for managing, cleaning, and analyzing company datasets.

## 🎯 Versions

### 1. **Streamlit Web App** (`app.py`)
- Browser-based interface
- Interactive visualizations with Plotly
- Best for: Team collaboration, data exploration

**Run:**
```powershell
streamlit run app.py
```

### 2. **Native Desktop App (Tkinter)** (`app_desktop.py`)
- Built-in Python GUI (tkinter)
- No additional dependencies needed
- Lightweight, offline-ready
- **Best for: Standalone desktop use**

**Run:**
```powershell
python app_desktop.py
```

### 3. **Desktop App (PySimpleGUI)** (`app_pysimple.py`)
- Alternative GUI framework
- Can be compiled to `.exe` with PyInstaller

**Run:**
```powershell
python app_pysimple.py
```

## ✨ Features (All Versions)

- **📤 Data Import**: Load CSV/Excel files
- **🧹 Data Cleaning Pipeline**:
  - Normalize column names → snake_case
  - Trim whitespace in text fields
  - Remove duplicate rows
  - Handle missing values (keep/drop/fill)
  - Normalize email and phone formats
- **🔍 Search & Filter**: Text search across all columns
- **💾 Export**:
  - CSV format
  - Excel with styled headers
- **🏢 Company Registry**: CRUD operations
- **📸 Snapshots**: Save/load data versions
- **📊 Analytics**: Missing data statistics, distributions
- **💾 SQLite Storage**: Local database for persistence

## 📋 Project Structure

```
.
├── app.py                    # Streamlit web app
├── app_desktop.py            # Tkinter desktop app
├── app_pysimple.py           # PySimpleGUI desktop app
├── requirements.txt          # Dependencies
├── BUILD_EXE.md             # Instructions for building exe
├── data_manager.db          # SQLite database (created at runtime)
└── README.md                # This file
```

## 🚀 Quick Start

### Requirements
- Python 3.8+
- pandas, openpyxl, plotly, streamlit (for web version)

### Setup

1. Create virtual environment:
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies:
```powershell
pip install -r requirements.txt
```

3. Choose your version and run:

**Desktop (Recommended):**
```powershell
python app_desktop.py
```

**Web:**
```powershell
streamlit run app.py
```

## 🛠️ Building Executable

To create a standalone `.exe` file:

```powershell
pip install PyInstaller
pyinstaller --onefile --windowed --name "DataManager" app_desktop.py
# Output: dist\DataManager.exe
```

## 📖 Usage

### Data Pipeline Tab
1. Load CSV/Excel file or load from saved snapshot
2. Configure cleaning options
3. Run cleaning pipeline
4. Search/filter results
5. Export to CSV/Excel
6. Save as snapshot for later use

### Company Registry Tab
1. Add company details (name required)
2. View all companies in list
3. Delete by ID

### Snapshots Tab
1. View all saved data snapshots
2. Load previous versions
3. Track data cleaning history

## 🗄️ Database

SQLite database (`data_manager.db`) stores:
- **dataset_snapshots**: Raw and cleaned data versions
- **companies**: Company registry records

## 📦 Dependencies

```
streamlit>=1.28.0
pandas>=2.0.0
numpy>=1.23.0
plotly>=5.13.0
openpyxl>=3.0.0
python-dateutil>=2.8.2
PySimpleGUI>=4.60.0
PyInstaller>=6.0.0
```

## 📝 Notes

- Tkinter version (`app_desktop.py`) is recommended for standalone desktop use
- All versions use the same SQLite database and core logic
- Excel export includes formatted headers for better readability
- Phone numbers automatically formatted with/without country codes
- Emails normalized to lowercase

## 🔗 Repository

https://github.com/Hungthang1234/DataRnAPP

---

**Status**: Production Ready | Last Updated: Mar 2026
