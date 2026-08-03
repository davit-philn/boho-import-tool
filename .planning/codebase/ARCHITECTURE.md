<!-- refreshed: 2026-08-03 -->
# Architecture

**Analysis Date:** 2026-08-03

## System Overview

```text
┌────────────────────────────────────────────────────────────────────────┐
│                        GUI / Views Layer                               │
│                  `Tools/views/main_window.py`                          │
│    (BOMToolApp + Tab UI: Import BOM, THDM, Catalog)                    │
└────────────────────┬─────────────────────────────────────────────────┘
                     │
        ┌────────────┼────────────┬─────────────┐
        │            │            │             │
        ▼            ▼            ▼             ▼
┌──────────────┐ ┌────────────┐ ┌──────────┐ ┌──────────────┐
│   Parser     │ │  Mapping   │ │Validators│ │   Utils      │
│  Layer       │ │  Loader    │ │  Layer   │ │  (Helpers)   │
│ bom_parser.py│ │mapping_    │ │validators│ │  utils.py    │
│  (~1100 lines)│ │loader.py   │ │  .py     │ │   +tokens    │
└──────────────┘ └────────────┘ └──────────┘ └──────────────┘
        │              │              │             │
        └──────────────┼──────────────┴─────────────┘
                       │
        ┌──────────────▼───────────────┐
        │    Config/Mapping Files      │
        │  (CK_Mapping_v5.xlsx)        │
        │  (db_config.json)            │
        │  (settings.json)             │
        └──────────────┬───────────────┘
                       │
        ┌──────────────▼───────────────┐
        │   SQL Server Database        │
        │  (Bravo ERP)                 │
        │  - B20BOM                    │
        │  - B20BOMDetail              │
        │  - B20Item                   │
        │  - Master tables (lookups)   │
        └──────────────────────────────┘
```

## Component Responsibilities

| Component | Responsibility | File |
|-----------|----------------|------|
| BOMToolApp | Main GUI window, tab management, event handling, session state | `Tools/views/main_window.py` |
| SheetTable | Custom data grid widget for preview/display | `Tools/views/widgets.py` |
| BOM Parser | Parse Excel sheets (BOM/THDM), extract metadata, resolve row mapping | `Tools/services/bom_parser.py` |
| Mapping Loader | Read config sheets from CK_Mapping_v5.xlsx, build reverse column maps | `Tools/services/mapping_loader.py` |
| Validators | Layer 1 validation (types, required fields, length) | `Tools/services/validators.py` |
| Utils | Shared constants, paths, Vietnamese text normalization, column alignment hints | `Tools/services/utils.py` |

## Pattern Overview

**Overall:** Layered Model-View-Control with configuration-driven parsing

**Key Characteristics:**
- **Config-driven:** All Excel parsing behavior (section detection, data start row, field mapping, validation rules) stored in `CK_Mapping_v5.xlsx` — no hardcoded parsing logic
- **Separation of concerns:** GUI layer completely independent from business logic (services)
- **Thread-safe operations:** Long-running DB operations and file parsing run in background threads, updates posted back to GUI via `self.after()`
- **Modular widgets:** Reusable UI components (`CLabel`, `CButton`, `_SearchCombo`, `SheetTable`, `Tooltip`)
- **Theme system:** Dark/Light switchable via `THEMES` dict in utils.py, persisted in settings.json

## Layers

**Presentation Layer (Views):**
- Purpose: Display data, handle user input, display results
- Location: `Tools/views/`
- Contains: GUI components (`BOMToolApp`, `SheetTable`, custom widgets)
- Depends on: Services layer for data operations
- Used by: User interactions, OS events

**Business Logic Layer (Services):**
- Purpose: Parse files, validate data, generate SQL, execute DB operations
- Location: `Tools/services/`
- Contains: Parser, mapping loader, validators, utilities
- Depends on: External libs (pandas, openpyxl, pyodbc), config files
- Used by: Views layer for data operations

**Data & Configuration Layer:**
- Purpose: Provide mapping rules, database connection, user settings
- Location: `Tools/config/` (runtime), `Mapping/` (reference), `Structure/` (schema docs)
- Contains: Mapping Excel file, DB config, application settings
- Depends on: None
- Used by: Services and Views layers for configuration

**Persistence Layer:**
- Purpose: Store data in SQL Server
- Location: SQL Server (Bravo ERP)
- Contains: B20BOM, B20BOMDetail tables + master data tables
- Depends on: pyodbc connection with valid credentials
- Used by: Business logic layer for INSERT/UPDATE/SELECT operations

## Data Flow

### Primary Request Path: BOM Import

1. **User selects Excel file** → `BOMToolApp._on_load_file()` (`main_window.py:~3100`)
2. **Load mapping config** → `load_mapping()` reads `CK_Mapping_v5.xlsx` sheets (_CONFIG, HEADER, DETAIL, etc.) (`mapping_loader.py:101`)
3. **Parse Excel file** → `parse_bom_file()` unpacks all sheets, detects section type using sheet_config, extracts metadata (`bom_parser.py:638`)
4. **Build reverse column map** → `build_reverse_map()` creates {normalized_excel_col → sql_col} lookup dict (`mapping_loader.py:175`)
5. **Parse each sheet/section**:
   - `_parse_sheet()` iterates rows, finds header row, extracts metadata (Dự án, Đơn hàng, etc.)
   - `_parse_section_excel_rows()` processes data rows: detects footers, section headers, applies Fill_Forward logic
   - `_resolve_row_mapping()` maps each Excel row dict to SQL columns using reverse_map
   - Filtered rows collected in section-specific table
6. **Validate data** → `validate_layer1()` checks required fields, data types, string lengths (`validators.py:35`)
7. **Display results** → Populate `self.tables` dict with parsed section tables, display in `tree` widget
8. **User clicks Import** → Generate SQL INSERT statement (`_generate_bom_header()`)
9. **Preview SQL** → Optional: show generated SQL in popup for offline verification
10. **Execute INSERT** → Connect to SQL Server via `pyodbc`, INSERT B20BOM header row, get identity, INSERT B20BOMDetail rows, COMMIT (`main_window.py:9525`)

### Secondary Flow: THDM (Tổng Hợp Định Mức) Processing

1. **User selects THDM file** in THDM tab
2. **Load THDM mapping** from mapping config
3. **Find TH VT (Tổng Hợp Vật Tư) sheet** → `_thdm_find_thvt_sheet()` scans sheet names
4. **Parse header row** → Detect columns via normalized matching (using aliases for common variations)
5. **Expand Mục rows** → For each CHI TIẾT GIAO row, cross-reference BOM quantities via `_thdm_load_bom_qty_dict()` (queries B20BOMDetail)
6. **Apply row filters** → Optional RowFilter expression (from _CONFIG)
7. **Transpose/restructure** → Convert detail data per mapping rules, apply formula fields
8. **Insert THDM data** → Direct INSERT to B20Item and detail tables

### State Management

**Session State (BOMToolApp instance variables):**
- `self.tables` — parsed data tables by section label (e.g., "[BOM] II_Vật Tư Chính")
- `self.global_meta` — extracted metadata (Dự án, Đơn hàng, Tên SP, etc.)
- `self.val_errors` — validation errors by table label
- `self.mapping` — loaded mapping config (dict)
- `self._current_file` — path to currently loaded Excel file
- `self._current_creator_user_id` — selected user ID for CreatedBy

**Persistent Settings (settings.json):**
- Theme ("Dark" or "Light")
- Scaling percentage ("100%", "125%", etc.)

## Key Abstractions

**Reverse Mapping:**
- Purpose: Map Excel column headers (after normalization) to SQL column names
- Examples: `{norm_ten_excel → sql_col}` — created by `build_reverse_map()`
- Pattern: Exact match first, then prefix-match fallback via `match_col_to_sql()`

**Sheet Config:**
- Purpose: Describe how to detect and parse BOM sections (A, B1, B2, C, D, E1, E2, etc.)
- Examples: `{"section": "BOM2", "contains": ["II"], "excludes": [], "data_start_row": "AUTO"}`
- Pattern: Loaded from _CONFIG sheet, used by `_detect_sheet_type()` and `_load_sheet_config()`

**Fill_Forward Logic:**
- Purpose: Carry material classification from section headers into data rows
- Examples: STT=A → Vật liệu column captured, propagated to rows A.1, A.2, etc.
- Pattern: Identified by mapping record `fill_forward='1'`, applied in `_parse_section_excel_rows()`

**Column Map:**
- Purpose: Map Excel column index to SQL column name for a specific section
- Examples: `{0: 'STT', 1: 'ItemId', 2: 'ItemName', ...}`
- Pattern: Built by `_build_excel_col_map()` per section

## Entry Points

**Application Launch:**
- Location: `Tools/main.py`
- Triggers: User runs `python main.py` or `.exe`
- Responsibilities: Load theme from settings, initialize CustomTkinter, create and run `BOMToolApp()`

**File Selection:**
- Location: `BOMToolApp._on_load_file()` in `main_window.py:~3100`
- Triggers: User clicks "Load File" button or drags Excel file
- Responsibilities: Open file dialog, validate file exists, trigger background parse thread

**Tab Switch Events:**
- Location: `BOMToolApp._on_tab_changed()` in `main_window.py:~2800`
- Triggers: User clicks different tab (Import BOM, THDM, Catalog)
- Responsibilities: Reload mapping if needed, update UI state

**Import Button:**
- Location: `BOMToolApp._on_import_clicked()` in `main_window.py:~9400`
- Triggers: User clicks "Import" after validating data
- Responsibilities: Generate SQL, connect DB, execute INSERT in background thread

## Architectural Constraints

- **Single Excel file:** Only one file loaded at a time; switching files clears previous tables
- **Threading model:** GUI thread runs UI loop; long operations (file parse, DB insert) run in daemon threads, communicate back via `self.after()`
- **No mutable global state in services:** All state passed as function arguments or stored in `BOMToolApp` instance
- **Encrypted Excel fallback:** If `msoffcrypto` not installed, password-protected files cannot be opened
- **tksheet fallback:** If `tksheet` not installed, UI falls back to `ttk.Treeview` (slower but functional)
- **SQL Server only:** pyodbc + ODBC Driver 17 required; no support for MySQL, PostgreSQL, SQLite
- **Python version:** 3.10+ required for type hints and f-strings used throughout

## Anti-Patterns

### Hardcoded Column Names

**What happens:** Early code used literal strings like `if 'Vật Tư' in header_name` scattered through parsing logic

**Why it's wrong:** Adding a new BOM section (e.g., "BOM5") requires finding and updating parsing logic in multiple places; configuration changes break code without obvious error messages

**Do this instead:** Add row to _CONFIG sheet in mapping file with `SheetNameContains`, `DataStartRow`, and field mapping; parsing logic `_detect_sheet_type()` and `_load_sheet_config()` handle it automatically (`bom_parser.py:33-68`)

### Silently Skipping Invalid Rows

**What happens:** Rows with no ItemId AND no ItemName are discarded without warning during `_parse_section_excel_rows()`

**Why it's wrong:** User may not realize data is being dropped; import succeeds but quantity is wrong

**Do this instead:** Track and report skipped rows in validation layer; generate warning message showing which rows were excluded and why (enhancement: add to validation report)

### Mixed Responsibilities in main_window.py

**What happens:** GUI class `BOMToolApp` contains DB connection logic, SQL generation, and UI rendering in single 10,000+ line file

**Why it's wrong:** Hard to test business logic independently; changes to SQL logic affect UI classes

**Do this instead:** Extract SQL generation to `db_layer.py` service module; DB operations limited to connection management and transaction control; see `_generate_bom_header()` and `_generate_bom_details()` — these could be moved to services with GUI calling them

## Error Handling

**Strategy:** Try-catch in UI event handlers; background threads catch exceptions, return error dict, GUI displays message box

**Patterns:**
- File parsing errors → catch, display error message with file path
- DB connection errors → check if server/database reachable before attempting INSERT; display connection string (with password masked) to help debug
- Validation errors → collect all errors, display in grid with severity colors (red=error, orange=warning), export to Excel report
- Excel encryption → prompt for password, allow 3 retries, fallback to cancel

**Example:** `_run_insert_bg()` at `main_window.py:9525` — wrapped in try-except, catches all exceptions, posts result via `self.after()` to GUI thread for message display

## Cross-Cutting Concerns

**Logging:** File-based log stored in `config/` directory; tab-delimited format with columns: FileName, Action, Count, Status, Message, DatetimeUTC. Method: `_log()` in `main_window.py`

**Validation:** Two-pass approach:
1. Layer 1 (structural) — type checking, required fields, string lengths → `validate_layer1()` in `validators.py`
2. Layer 2 (business rules) — master data lookups, state validation → in `_run_insert_bg()` before INSERT

**Authentication:** SQL Server authentication via `db_config.json` (username/password or integrated Windows auth via `Trusted_Connection=yes`)

**Theme Persistence:** User selects Dark/Light, applied to entire app (CTk + ttk.Treeview + custom tk.Listbox widgets), saved to `settings.json` for next startup

---

*Architecture analysis: 2026-08-03*
