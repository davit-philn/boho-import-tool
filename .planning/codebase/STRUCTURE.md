# Codebase Structure

**Analysis Date:** 2026-08-03

## Directory Layout

```
insertdatasqlserver/
├── Tools/                          # Main application directory
│   ├── main.py                     # Application entry point (launches BOMToolApp)
│   ├── build.py                    # PyInstaller build script (outputs .exe + Portable ZIP)
│   ├── BOHO_IMPORT_BOM_THDM.spec   # PyInstaller spec file for build configuration
│   ├── setup.iss                   # Inno Setup script for Windows installer
│   ├── install_deps.bat            # Batch file to install Python dependencies
│   │
│   ├── config/                     # Runtime configuration (user-editable, not committed)
│   │   ├── CK_Mapping_v5.xlsx      # MASTER: Maps Excel columns → SQL columns (7 sheets: _CONFIG, HEADER, DETAIL, etc.)
│   │   ├── db_config.json          # Database connection credentials (IGNORED by .gitignore)
│   │   └── settings.json           # UI theme and scaling preferences (auto-created)
│   │
│   ├── services/                   # Business logic layer (no UI code)
│   │   ├── __init__.py             # Package marker
│   │   ├── bom_parser.py           # Excel file parser: BOM sections + THDM sheets (~1100 lines)
│   │   ├── mapping_loader.py       # Load CK_Mapping_v5.xlsx, build reverse column maps
│   │   ├── validators.py           # Data validation: types, required fields, lengths
│   │   └── utils.py                # Shared: paths, constants, Vietnamese normalization, theme colors
│   │
│   ├── views/                      # Presentation layer (GUI only)
│   │   ├── __init__.py             # Package marker
│   │   ├── main_window.py          # BOMToolApp: main window, all 3 tabs (Import BOM, THDM, Catalog)
│   │   └── widgets.py              # Reusable widgets: CLabel, CButton, _SearchCombo, SheetTable, Tooltip
│   │
│   ├── scripts/                    # Ad-hoc utility scripts (not part of main app)
│   │   ├── dump_mapping.py         # Debug: print mapping config to console
│   │   ├── patch_mapping.py        # Utility: update mapping Excel programmatically
│   │   ├── read_docx.py            # Extract text from .docx (experimental)
│   │   ├── test_insert_bom.py      # Test: sample SQL INSERT
│   │   ├── test_sp_version.py      # Test: call SQL Server stored procedures
│   │   ├── testUser.py             # Test: user/employee lookup
│   │   └── update_mapping_cols.py  # Utility: bulk update mapping columns
│   │
│   ├── sql/                        # SQL scripts for schema setup (not executed by app)
│   │   ├── create_bomtool_log.sql  # Create logging table for import audit trail
│   │   ├── create_import_template_tables.sql  # Create temp staging tables
│   │   └── create_usp_bomtool_deletebom.sql   # Create stored proc for undo import
│   │
│   ├── archive/                    # Old code versions (kept for reference)
│   │   ├── main_v2.py - main_v6.py # Previous implementations
│   │   └── config_old/             # Old configuration backups
│   │
│   ├── dist/                       # PyInstaller output (built by build.py)
│   │   └── BOHO_IMPORT_BOM_THDM/   # Runnable app directory
│   │
│   ├── build/                      # PyInstaller intermediate build directory
│   ├── installer_output/           # Final deliverables (Portable ZIP, Installer EXE)
│   ├── redist/                     # Runtime redistributables (unused in onedir mode)
│   │
│   └── icon.ico                    # Application icon (embedded in .exe)
│
├── Mapping/                        # Reference mapping files (legacy/backups)
│   ├── CK_Mapping_v2.xlsx          # Old format (do not use)
│   ├── CK_Mapping_v2_bak.xlsx      # Backup of v2
│   ├── CK_MapsCotGiuaExcelVaSQL_BOM_THDM (1).xlsx
│   └── Mapping_Column.xlsx
│
├── Structure/                      # Database schema reference (documentation)
│   ├── B20BOM.xlsx                 # B20BOM table structure
│   ├── B20BOMDetail .xlsx          # B20BOMDetail table structure
│   └── B20Item.xlsx                # B20Item table structure
│
├── Template/                       # Sample/template Excel files for users
│   └── [BOM template files]
│
├── Doc/                            # User documentation
│   └── [.docx design specs, usage guides]
│
├── DemoData/                       # Sample Excel files for testing
│   └── [example BOM, THDM files]
│
├── BravoProgram/                   # (Legacy) Bravo ERP reference files
├── BravoDocuments/                 # (Legacy) Bravo documentation
├── Other/                          # Miscellaneous files
│
├── _backups/                       # Manual backups (not synced)
├── _releases/                      # Packaged releases
├── backup_phaseB_working/          # Phase B backup (full project snapshot)
│
├── requirements.txt                # Python dependencies (root level)
├── .gitignore                      # Git exclusions (ignores db_config.json, *.exe, __pycache__)
├── README.md                       # Project overview and setup guide
└── .planning/
    └── codebase/                   # This documentation (generated by gsd-map-codebase)
        ├── ARCHITECTURE.md         # System design, data flow, patterns
        ├── STRUCTURE.md            # This file
        ├── CONVENTIONS.md          # (if generated) Coding style guidelines
        ├── TESTING.md              # (if generated) Test structure and patterns
        ├── STACK.md                # (if generated) Technology stack
        ├── INTEGRATIONS.md         # (if generated) External services (SQL Server, etc.)
        └── CONCERNS.md             # (if generated) Technical debt and issues
```

## Directory Purposes

**Tools/**
- Purpose: Complete application — entry point, services, UI, config, build artifacts
- Contains: Python source code, configuration, compiled output
- Key files: `main.py` (entry), `services/` (logic), `views/` (UI), `config/` (settings)

**Tools/config/**
- Purpose: Runtime configuration user must provide/customize
- Contains: CK_Mapping_v5.xlsx (MASTER config), db_config.json (credentials), settings.json (UI preferences)
- Key dependency: App cannot run without valid `db_config.json` with server/database/credentials

**Tools/services/**
- Purpose: Business logic layer — parsing, validation, no UI dependencies
- Contains: Excel parser, mapping loader, validators, utilities
- Key modules:
  - `bom_parser.py` (~1100 lines) — Core: parse Excel sheets, extract metadata, map columns to SQL
  - `mapping_loader.py` — Load and organize mapping config from Excel
  - `validators.py` — Validate parsed data against schema
  - `utils.py` — Shared: paths, constants, color themes, text normalization

**Tools/views/**
- Purpose: Presentation layer — GUI rendering, event handling, user interaction
- Contains: CustomTkinter application window, reusable widgets
- Key modules:
  - `main_window.py` (~9600 lines) — BOMToolApp: main window, 3 tabs (BOM Import, THDM, Catalog), all event handlers
  - `widgets.py` — Reusable: CLabel, CButton, _SearchCombo (searchable dropdown), SheetTable (data grid), Tooltip

**Tools/scripts/**
- Purpose: Development utilities, not part of production app
- Contains: Debug scripts, mapping utilities, test data insertion
- Usage: Run manually for debugging, not called by main application

**Tools/sql/**
- Purpose: Database initialization scripts (customer deploys these separately)
- Contains: Create logging table, stored procedures for undo
- Usage: Run against SQL Server manually before first use

**Mapping/**
- Purpose: Archive/reference for mapping file evolution
- Contains: Old mapping Excel formats (v2, etc.)
- Usage: Do NOT use these — use Tools/config/CK_Mapping_v5.xlsx instead

**Structure/**
- Purpose: Reference documentation of SQL table schemas
- Contains: Excel files with B20BOM, B20BOMDetail, B20Item column definitions
- Usage: For users to understand data structure, not used by app code

## Key File Locations

**Entry Points:**
- `Tools/main.py` — Application launcher (initializes CustomTkinter theme, creates BOMToolApp, enters mainloop)
- `Tools/build.py` — Build script to create .exe and Portable ZIP (for developers)

**Configuration:**
- `Tools/config/CK_Mapping_v5.xlsx` — MASTER configuration: defines all BOM sections, field mappings, validation rules, stored procedures
- `Tools/config/db_config.json` — Database credentials (created with template if missing; user must fill in)
- `Tools/config/settings.json` — UI theme and scaling (auto-created with defaults)

**Core Logic:**
- `Tools/services/bom_parser.py` — Parse Excel files: `parse_bom_file()` (entry), `_parse_sheet()` (sheet processor), `_parse_section_excel_rows()` (row processor)
- `Tools/services/mapping_loader.py` — Load config: `load_mapping()` (reads 7 sheets from mapping file), `build_reverse_map()` (creates lookup dict)
- `Tools/services/validators.py` — Validate data: `validate_layer1()` (type, required, length checks)
- `Tools/services/utils.py` — Utilities: `BASE_DIR`, `MAPPING_FILE`, `_norm_vn()` (Vietnamese normalization), `THEMES` dict (Dark/Light colors)

**GUI:**
- `Tools/views/main_window.py` — Main app: `BOMToolApp.__init__()` (build UI), `_on_tab_changed()` (reload mapping when tab switches), `_on_load_file()` (parse Excel), `_on_import_clicked()` (insert to DB)
- `Tools/views/widgets.py` — Custom widgets: `SheetTable` (data grid), `_SearchCombo` (searchable dropdown), `CLabel`/`CButton` (theme-aware labels/buttons), `Tooltip`

**Testing:**
- `Tools/scripts/test_insert_bom.py` — Sample BOM insert (manual test)
- `Tools/scripts/test_sp_version.py` — Test stored procedure calls
- (No automated test suite currently; all validation via manual testing + validation layer)

## Naming Conventions

**Files:**
- Underscore prefix for private modules/functions: `_SearchCombo`, `_parse_section_excel_rows()`
- Service modules: snake_case (`bom_parser.py`, `mapping_loader.py`, `validators.py`, `utils.py`)
- View modules: snake_case (`main_window.py`, `widgets.py`)
- Config files: lowercase with version suffix (`CK_Mapping_v5.xlsx`, `db_config.json`, `settings.json`)

**Python Functions:**
- Snake_case: `parse_bom_file()`, `load_mapping()`, `build_reverse_map()`, `validate_layer1()`
- GUI event handlers: `_on_<event_name>()` (e.g., `_on_tab_changed()`, `_on_load_file()`, `_on_import_clicked()`)
- Internal helpers: leading underscore `_is_roman_numeral()`, `_norm_col()`, `_build_excel_col_map()`

**Classes:**
- PascalCase: `BOMToolApp`, `SheetTable`, `CLabel`, `CButton`, `Tooltip`, `_SearchCombo`
- Exception names: suffixed `Error` (not used in current code)

**Variables:**
- Global constants: UPPERCASE (`APP_VERSION`, `TAB_IMPORT`, `TAB_THDM`, `TAB_CATALOG`, `BASE_DIR`, `MAPPING_FILE`, `THEMES`)
- Instance variables: snake_case (`self.tables`, `self.mapping`, `self.global_meta`, `self._current_file`)
- Loop variables: short names (`i`, `r`, `row`, `col`, `h`)

**Columns/Fields:**
- Excel headers: Vietnamese names with spaces (e.g., "Mã SP", "Tên chi tiết", "Số lượng")
- SQL columns: PascalCase (e.g., `STT`, `ItemId`, `ItemName`, `ItemCode`, `Quantity`)
- Normalized (search): lowercase, no diacritics, no spaces (e.g., `masptenchitiet` from "Mã SP - Tên chi tiết")

## Where to Add New Code

**New Feature (e.g., add BOM9 section):**
1. **Define section in mapping:** Add row to `_CONFIG` sheet in `Tools/config/CK_Mapping_v5.xlsx`
   - Column A: Section name (e.g., "BOM9")
   - Column B: Label (display name)
   - Column C: View_Insert (SQL view name for import)
   - Column D: SheetNameContains (e.g., "IX" for sheet named "BOM_IX")
   - Column E: SheetNameExclude (exclude keywords if any)
   - Columns F+: DataStartRow, ParentSection, etc.
2. **Add field mapping:** Add rows to HEADER or DETAIL sheet mapping columns STT, ItemId, ItemName, etc. to SQL columns
3. **Testing:** Load file with BOM9 sheet, verify section detected and parsed correctly (no code changes needed in services layer)

**New Validation Rule:**
1. **Add validator to mapping:** Add row to VALIDATORS sheet in `CK_Mapping_v5.xlsx`
   - Column A: Section
   - Column B: ValidatorName
   - Column C: SQL (SQL query returning error rows or NULL if valid)
   - Columns D+: Params, WarningMessage, WarnOnly, IsActive
2. **Call validator in UI:** `validate_layer1()` automatically picks up new validators from mapping
3. **Testing:** Load file, click Validate, verify new rule displays errors if data violates it

**New UI Widget or Tab:**
1. **Add widget class to `Tools/views/widgets.py`** — inherit from `ctk.CTkFrame` or `ctk.CTk*`
2. **Add theme support** — accept `THEMES` color dict in `__init__`, apply to all sub-widgets
3. **Expose in `main_window.py`** — create instance in `_build_ui()`, add event handlers
4. **Test Dark/Light toggle** — call `apply_theme()` to verify colors switch correctly

**New Service Module (e.g., db_layer.py for SQL generation):**
1. **Create `Tools/services/db_layer.py`** — pure Python, no UI imports
2. **Export functions:** `generate_bom_insert_sql()`, `generate_bom_detail_rows()`, etc.
3. **Import in `main_window.py`:** Use functions in import flow instead of inline SQL
4. **Testing:** Can test with `-c "from services.db_layer import ...; ..."` without GUI

**New Configuration File:**
1. **Add to `Tools/config/`** — JSON preferred for portability
2. **Template creation:** In `_load_db_config()` pattern, create template if missing
3. **Path constant:** Add to `utils.py` (e.g., `CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")`)
4. **Loading logic:** `_load_config()` function in relevant service module

## Special Directories

**Tools/archive/:**
- Purpose: Version history of main.py evolution (v2–v6)
- Generated: No (manually archived)
- Committed: Yes (via Git)
- Usage: For reference only; old code should not be reused

**Tools/dist/ and Tools/installer_output/:**
- Purpose: PyInstaller outputs (built application)
- Generated: Yes (by `Tools/build.py`)
- Committed: No (ignored by .gitignore)
- Usage: Artifacts for distribution; delete before rebuild

**Tools/__pycache__/ and Tools/*/`__pycache__/:**
- Purpose: Python bytecode cache
- Generated: Yes (auto-generated by Python)
- Committed: No (ignored by .gitignore)
- Usage: Safe to delete; will be regenerated on next run

**Tools/config/:**
- Purpose: User-editable runtime configuration
- Generated: Partially (settings.json auto-created; db_config.json template created if missing)
- Committed: No — `.gitignore` excludes `db_config.json` and `settings.json` to protect credentials
- Usage: User fills in `db_config.json` with real credentials before first run

**_backups/ and _releases/:**
- Purpose: Manual backup snapshots (not integrated with version control)
- Generated: Manual backups by user
- Committed: No (ignored by .gitignore)
- Usage: For disaster recovery, not part of source control

---

*Structure analysis: 2026-08-03*
