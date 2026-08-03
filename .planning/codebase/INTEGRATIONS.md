# External Integrations

**Analysis Date:** 2026-08-03

## APIs & External Services

**None Detected**
- This application does not integrate with external REST APIs, webhooks, or cloud services
- All processing is local or direct database-connected

## Data Storage

**Databases:**
- **SQL Server** (via Bravo ERP system)
  - Connection: `pyodbc` using ODBC Driver 17 for SQL Server
  - Authentication: Windows integrated login (`trusted_connection=yes`) OR username/password
  - Config file: `Tools/config/db_config.json` (template: `{"server": "...", "database": "...", "username": "...", "password": "...", "driver": "ODBC Driver 17 for SQL Server", "timeout": 5}`)
  - Connection code: `views/main_window.py` lines 7561-7597 (`_get_db_connection()` method)
  - Timeout: Configurable, default 5 seconds for VPN detection
  - **No ORM** - Direct SQL execution via raw queries and stored procedures
  - Tables targeted (from `CK_Mapping_v5.xlsx`):
    - `B00UserList` - User accounts (CreatedBy field mapping)
    - `B20Employee` - Employee directory (for THDM tab creator selection)
    - `BOM_*` tables - Bill of Materials data storage
    - `THDM_*` tables - Material specifications storage
    - Custom schema determined by mapping configuration

**File Storage:**
- **Local filesystem only** - no cloud storage integration
  - Source Excel files: User selects via file dialog (`filedialog.askopenfilename`)
  - Export destinations: SQL scripts (`.sql`), validation reports (`.xlsx`), logs
  - Config directory: `Tools/config/` (contains `db_config.json`, `settings.json`, `CK_Mapping_v5.xlsx`)
  - Backup/archive directories: `.gitignore` excludes `_backups/`, `backup_phaseB_working/`, `Tools/archive/`

**Caching:**
- None - Application loads fresh data each session
- Mapping loaded once on startup: `load_mapping()` in `services/mapping_loader.py`
- In-memory DataFrames for current Excel file being processed

## Authentication & Identity

**Auth Provider:**
- **SQL Server Windows Authentication** (primary) OR username/password (secondary)
  - Windows mode: `Trusted_Connection=yes` (requires local Windows domain account)
  - User/password mode: UID/PWD in connection string
  - Config: `Tools/config/db_config.json` with optional `trusted_connection` field
  - Code: `views/main_window.py` lines 7574-7577 (conditional connection string builder)

**No external identity provider:**
- Application does not use OAuth2, SAML, LDAP, or cloud authentication
- User identity passed through:
  - BOM import: `CreatedBy` defaults to hardcoded user ID 823 (`DEFAULT_CREATOR_USER_ID` in `services/utils.py`) if employee not found
  - THDM import: User selected via dropdown in UI (`_current_creator_user_id`, `_thdm_creator_user_id`)
  - Logs: No audit trail of application user; only database-level SQL Server login

## Monitoring & Observability

**Error Tracking:**
- None - No external error/crash reporting
- Errors logged to console and status messages shown in UI messagebox
- Application logs: `*.log` files (excluded from git, not centralized)

**Logs:**
- Local file-based logging: Python logging to `*.log` in `Tools/logs/` (excluded from git by `.gitignore`)
- Database audit: Import history stored in SQL Server tables (determined by mapping configuration)
- UI displays: Recent 200 import records fetched via SQL query (code: `views/main_window.py` line 6802)
- No structured log aggregation (ELK, Splunk, etc.)

## CI/CD & Deployment

**Hosting:**
- Desktop application (Windows 10/11 local executable)
- Not cloud-hosted; end-users run `.exe` file locally
- No server component

**CI Pipeline:**
- None detected - No GitHub Actions, Jenkins, or automated build pipeline
- Manual build via `python Tools/build.py` creates:
  - `dist/BOHO_IMPORT_BOM_THDM/` - Development/test build
  - `installer_output/BOHO_ImportBOM_THDM_v2.1_Portable.zip` - Portable distribution
  - `installer_output/BOHO_ImportBOM_THDM_Setup_v2.1.exe` - Installer (requires Inno Setup 6 external tool)

**Deployment Steps:**
- User extracts `.zip` or runs `.exe` installer
- First run: App creates `Tools/config/` with templates for `db_config.json`, `settings.json`
- User manually edits `db_config.json` with SQL Server credentials
- App loads `CK_Mapping_v5.xlsx` from config directory (not bundled in `.exe` to allow user updates)

## Environment Configuration

**Required env vars:**
- None - Application uses `db_config.json` file instead of environment variables

**Secrets location:**
- `Tools/config/db_config.json` (git-ignored, created locally by user)
  - Contains: SQL Server server name/IP, database name, username, password
  - Never committed to repository
  - Template auto-created on first run if missing

## Webhooks & Callbacks

**Incoming:**
- None - Application does not expose any HTTP endpoints or receive webhooks

**Outgoing:**
- None - Application does not call external APIs or POST to webhooks
- Only interaction: Reading Excel files, connecting to SQL Server database, executing stored procedures

## Excel File Integration

**Input Format:**
- Source: User-provided Excel files (`.xlsx`)
- Structure: Multiple sheets (A, B1, B2, C, D, E1, E2, F, G, H, I, J, THDM sheets)
- Processing:
  - Optional password protection support (requires `msoffcrypto` library)
  - Sheet detection via regex patterns and `_CONFIG` mapping
  - Header row auto-detection or config-specified row number
  - Data parsing into pandas DataFrames

**Output Format:**
- SQL INSERT preview: User can view generated SQL before committing
- Export: SQL script file (`.sql`) for offline review
- Validation report: Excel file (`.xlsx`) with error/warning details

**Mapping Configuration:**
- Source: `CK_Mapping_v5.xlsx` (7 sheets defining column mappings)
- Purpose: Define Excel column → SQL column/table mapping without hardcoding
- Loaded: `services/mapping_loader.py` - `load_mapping()` on startup
- Usage: Column matching, validation rules, data type enforcement

## Database Stored Procedures

**Execution Pattern:**
- Generic caller: `_call_sp()` in `views/main_window.py` (line 7600+)
- Parameter binding: String-based `@Key=Value` notation (no parameterized queries visible)
- Return value: First column, first row or None
- Configuration: Stored procedure names and parameters from mapping file

**Sample Procedures Likely Called:**
- Insert BOM data → destination tables
- Insert THDM material specification data
- Validation queries (check for duplicate records, constraint violations)
- Audit/log procedures

---

*Integration audit: 2026-08-03*
