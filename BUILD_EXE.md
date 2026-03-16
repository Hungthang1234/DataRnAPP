# Build EXE from PySimpleGUI app

In command prompt/PowerShell, run:

```powershell
cd path\to\DataRnAPP
.\.venv\Scripts\python.exe -m pip install -q PySimpleGUI PyInstaller
.\.venv\Scripts\pyinstaller.exe --onefile --windowed --name "DataManager" app_pysimple.py
```

Executable will be in: `dist\DataManager.exe` (~65 MB)

## Alternative: Build with console (for debugging)

```powershell
.\.venv\Scripts\pyinstaller.exe --onefile --console --name "DataManager" app_pysimple.py
```

## Features

- ✅ Desktop GUI (PySimpleGUI) - no browser needed
- ✅ SQLite database for snapshots
- ✅ Load CSV/Excel files
- ✅ Data cleaning pipeline
- ✅ Search & filter
- ✅ Export CSV/Excel
- ✅ Company registry CRUD
- ✅ Snapshot management
