# Build Information

## Executive Executable Build

**Application**: Data Manager Application
**Version**: 1.0.0
**Build Date**: March 16, 2026
**Build Tool**: PyInstaller 6.19.0

### Executable Details

```
File: DataManager.exe
Location: dist/DataManager.exe
Size: ~68.5 MB
Type: Windows Standalone Executable
Architecture: 64-bit
```

### What's Included

The executable bundles:
- **Python Runtime**: Python 3.x interpreter
- **GUI Framework**: tkinter (built-in)
- **Dependencies**: pandas, numpy, openpyxl, sqlite3
- **Data Cleaning Module**: data_cleaner.py
- **Database**: SQLite support (creates data_manager.db automatically)

### No Installation Needed

Simply run `DataManager.exe` - no Python installation or dependencies required!

### Features

📊 **Data Pipeline**
- Load CSV/Excel files
- Clean and normalize data
- Remove duplicates
- Handle missing values
- Export to CSV/Excel
- Save data snapshots

🏢 **Company Registry**
- Add/manage company records
- Auto-normalize contact info
- Store in SQLite database

💾 **Snapshots**
- Save data versions
- Load previous snapshots
- Track data history

🔍 **Data Quality Analysis**
- Generate quality reports
- Detect outliers (IQR & Z-Score)
- Advanced cleaning pipeline
- Before/after comparison

### How to Build Again

```powershell
# Install PyInstaller (if needed)
pip install pyinstaller

# Build executable
.\.venv\Scripts\pyinstaller.exe --onefile --windowed --name "DataManager" app_desktop.py

# Result: dist/DataManager.exe
```

### Build Options Used

- `--onefile`: Single executable file (not folder)
- `--windowed`: GUI app (no console window)
- `--name "DataManager"`: Custom executable name

### Technical Stack (Compiled Into Exe)

| Component | Version | Purpose |
|-----------|---------|---------|
| Python | 3.x | Runtime |
| pandas | Latest | Data manipulation |
| numpy | Latest | Numerical operations |
| openpyxl | Latest | Excel export |
| tkinter | Built-in | GUI framework |
| sqlite3 | Built-in | Database |

### Distribution

The executable is completely standalone and can be:
- Copied to any Windows machine (64-bit)
- Run without any installation
- Shared directly with users
- Placed on USB drives
- No dependency on Python installation on user's machine

### First Run

When first launched, the application will:
1. Create `data_manager.db` (SQLite database) in the same directory
2. Initialize database tables
3. Display the main window with 4 tabs

### Data Files

- `data_manager.db`: Created automatically in the same directory as exe
- Dataset files: Load CSV/Excel files via the interface
- Exports: Saved in the same directory by default

### Notes

- File size is large (~68.5 MB) due to bundled Python runtime and libraries
- First launch might take 1-2 seconds (Python runtime initialization)
- Subsequent launches are faster
- All data is stored locally (no cloud connectivity)
- Fully offline application

### Version Control

Build artifacts are located in:
- `dist/DataManager.exe` - Ready to distribute
- `build/` - Intermediate build files
- `.spec` file - PyInstaller configuration

---

**Status**: ✅ Production Ready
**Release Date**: March 16, 2026
