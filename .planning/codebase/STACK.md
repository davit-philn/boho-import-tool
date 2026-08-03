# Technology Stack

**Analysis Date:** 2026-08-03

## Languages

**Primary:**
- Python 3.10+ - Desktop application, data processing, database operations
- SQL (T-SQL) - SQL Server stored procedures and queries executed via pyodbc

## Runtime

**Environment:**
- Python 3.10+ (minimum required for f-strings and modern syntax patterns)
- Windows 10/11 (required for pyodbc ODBC Driver and desktop GUI)

**Package Manager:**
- pip (Python dependency management)
- Lockfile: Not present (uses `requirements.txt` with flexible versions)

## Frameworks

**Core Application:**
- CustomTkinter 5.x - Modern desktop UI framework with dark/light theme support (`views/main_window.py`, `views/widgets.py`)
- tkinter (stdlib) - Standard Python GUI foundation, wrapped by CustomTkinter

**Data Processing:**
- pandas - DataFrame operations for Excel parsing and data manipulation (`services/bom_parser.py`, `services/mapping_loader.py`)
- openpyxl - Excel file reading/writing, format handling (`services/bom_parser.py`, entire Tools/main.py workflow)

**Database:**
- pyodbc - SQL Server connection driver using ODBC Driver 17 for SQL Server (`views/main_window.py` line 7564-7582)

**Text Processing:**
- rapidfuzz - Fuzzy string matching for Excel column name resolution and mapping (`services/bom_parser.py`)
- unicodedata - Vietnamese diacritic normalization (`services/utils.py`)

**Build/Deployment:**
- PyInstaller 6.x - Converts Python scripts to Windows `.exe` executable (onedir mode) (`Tools/BOHO_IMPORT_BOM_THDM.spec`, `Tools/build.py`)

**Optional Enhancements:**
- tksheet - High-performance spreadsheet widget for large data tables (graceful fallback to ttk.Treeview if missing)
- msoffcrypto - Opens password-protected Excel files (graceful fallback if missing)
- PIL/Pillow - Image handling for application icons

## Key Dependencies

**Critical:**
- `pandas` - Reads Excel sheets, parses BOM/THDM data structures into DataFrames for validation and transformation
- `openpyxl` - Direct Excel workbook access for detecting sheet structure, reading config metadata
- `pyodbc` - **ONLY** database connectivity mechanism; no ORM, raw SQL execution
- `customtkinter` - Modern Python GUI; replaces legacy tkinter widgets with native Windows styling
- `rapidfuzz` - Fuzzy match Excel column headers to SQL mapping configuration for error-tolerant column detection

**Infrastructure:**
- `unicodedata` - Normalize Vietnamese characters (decompose diacritics: "đ" → "d", strip combining marks)
- `concurrent.futures` - Threading for async operations (file dialog, long-running imports)
- `json` - Parse `db_config.json` and `settings.json` configuration files

## Configuration

**Environment:**
- `db_config.json` - SQL Server connection string (server, database, username, password, driver, optional timeout)
  - Location: `Tools/config/db_config.json` (created by app if missing, template in code line 7018-7027)
  - Not committed to git (listed in `.gitignore`)
  - Template fields: `server`, `database`, `username`, `password`, `driver` (default: "ODBC Driver 17 for SQL Server"), optional `trusted_connection`, `timeout`
  
- `settings.json` - User preferences (theme, scaling percentage)
  - Location: `Tools/config/settings.json`
  - Fields: `theme` ("Dark"/"Light"), `scaling` ("100%", "125%", etc.)

- `CK_Mapping_v5.xlsx` - Column mapping configuration (read-only for end users)
  - Location: `Tools/config/CK_Mapping_v5.xlsx` or `Mapping/CK_Mapping_v5.xlsx`
  - 7 sheets: `_CONFIG`, `HEADER`, `DETAIL`, `THDM_*` sections
  - Defines SQL table/column names, required fields, data types, validation rules

**Build:**
- `Tools/BOHO_IMPORT_BOM_THDM.spec` - PyInstaller configuration (onedir mode, UPX compression, icon embedding)
- `Tools/build.py` - Build orchestration script with multi-step output:
  - `dist/BOHO_IMPORT_BOM_THDM/` - Runnable directory (executable + `_internal/` resources)
  - `installer_output/BOHO_ImportBOM_THDM_v2.1_Portable.zip` - Portable ZIP distribution
  - `installer_output/BOHO_ImportBOM_THDM_Setup_v2.1.exe` - Installer (requires Inno Setup external tool)

## Platform Requirements

**Development:**
- Windows 10/11 (only platform supported; pyodbc and CustomTkinter are Windows-centric)
- Python 3.10+ (f-string syntax, type hints)
- ODBC Driver 17 for SQL Server (installed separately; bundled in `.exe` binary reference only)
- VPN access (required to reach remote SQL Server instance in production)

**Production:**
- Windows 10/11 desktop application
- SQL Server instance (any version supported by ODBC Driver 17)
- Network connectivity to SQL Server (firewall, VPN tunnel)
- Bravo ERP database schema (application target system)

---

*Stack analysis: 2026-08-03*
