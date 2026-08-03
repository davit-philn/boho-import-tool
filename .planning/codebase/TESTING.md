# Testing Patterns

**Analysis Date:** 2026-08-03

## Test Framework

**Runner:**
- No formal testing framework (pytest/unittest not used)
- Manual test scripts using standalone Python executables
- Test files run directly with `python test_name.py`
- Tests execute via conditional `if __name__ == "__main__":` pattern

**Assertion Library:**
- No dedicated assertion library
- Manual assertions via try/except blocks
- Status printed to console with `print()`
- Return values checked and printed for verification

**Run Commands:**
```bash
cd Tools
python scripts/test_insert_bom.py       # Test BOM insertion
python scripts/test_sp_version.py       # Test stored procedures
python scripts/testUser.py              # Test user authentication
```

## Test File Organization

**Location:**
- `Tools/scripts/` directory for all manual test scripts
- Test files not co-located with source — separate directory
- Standalone scripts, not integrated into source modules

**File Structure:**
```
Tools/scripts/
├── test_insert_bom.py       (164 lines) — Integration: BOM insert
├── test_sp_version.py       (71 lines)  — Integration: SP execution
├── testUser.py              (31 lines)  — Integration: user auth
├── create_test_tables.sql   — DDL setup
├── dump_mapping.py          — Utility: export mapping to JSON
├── patch_mapping.py         — Utility: modify mapping config
├── read_docx.py             — Utility: DOCX file reading
└── update_mapping_cols.py   — Utility: schema sync
```

**Naming:**
- Test files: `test_*.py` prefix
- Utility scripts: descriptive name without test_ prefix
- All scripts have module docstring explaining purpose

## Test Structure

**Suite Organization:**
- No test suites (no class-based grouping)
- Each file = one test workflow
- Sequential execution of test steps

**Example Pattern from `Tools/scripts/test_insert_bom.py`:**
```python
"""
Test insert 1 dòng vào B20BOM với defaults đã chốt.
Chạy: python Tools\test_insert_bom.py

Mục đích: xác nhận cấu trúc cột + defaults đúng trước khi nối vào Excel.
"""

# 1. Configuration setup
DEFAULTS = { ... }  # Constants defined at top

# 2. Helper functions
def gen_row_id() -> str:
    ...

def make_conn_str(cfg: dict) -> str:
    ...

# 3. Main execution
def main():
    # Step 1: Load config
    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)
    
    # Step 2: Connect
    conn = pyodbc.connect(make_conn_str(cfg), timeout=10)
    
    # Step 3: Prepare test data
    from_excel = { ... }
    row = {**DEFAULTS, **from_excel}
    
    # Step 4: Execute
    cur.execute(sql, values)
    new_id = cur.fetchone()[0]
    conn.commit()
    
    # Step 5: Verify
    cur.execute("SELECT ... WHERE Id = ?", new_id)
    r = cur.fetchone()
    print(f"  Id={r[0]}  Code={r[1]}  ...")
    
    conn.close()

if __name__ == "__main__":
    main()
```

**Test Data Setup:**
- Constants defined at module level: `DEFAULTS = {...}`, `USERNAME = "USER001"`
- Test data hardcoded in `from_excel` dict or SQL parameters
- Configuration loaded from JSON file (`db_config.json`)
- No fixtures or factory functions

**Execution Pattern:**
- Sequential steps with print() status indicators
- Try/except blocks catch failures
- Print result for manual inspection
- Exit with status code on fatal error: `sys.exit(1)`

## Verification Output

**Patterns from `Tools/scripts/test_insert_bom.py`:**
```python
print(f"Kết nối {cfg['server']} / {cfg['database']}...", end=" ")
conn = pyodbc.connect(make_conn_str(cfg), timeout=10)
print("OK\n")

print("INSERT B20BOM...")
print(f"  Code    : {row['Code']}")
print(f"  DocDate : {row['DocDate']}")
print()

try:
    cur.execute(sql, values)
    new_id = cur.fetchone()[0]
    conn.commit()
    print(f"✅  INSERT thành công — B20BOM.Id = {new_id}")
except Exception as e:
    conn.rollback()
    print(f"❌  LỖI:\n{e}")
    sys.exit(1)

print("\nDữ liệu đã lưu:")
print(f"  Id={r[0]}  Code={r[1]}  DocDate={r[2]}  ...")
print("\nDone.")
```

**Patterns from `Tools/scripts/test_sp_version.py`:**
```python
print("=== Test 1: named params, None qua literal NULL ===")
try:
    sql = f"EXEC {SP} @_ItemId0=?, @_ParentBizDocId=NULL, @_TypeB=NULL"
    cur = conn.cursor()
    cur.execute(sql, [item_id])
    r = cur.fetchone()
    print("OK →", r[0] if r else None)
except Exception as e:
    print("FAIL:", e)
```

## Mocking

**Framework:** No mocking framework used

**Patterns:**
- No mock objects or patch decorators
- Tests interact with real database
- Real file I/O (Excel files, JSON config)
- Real network calls to database server

**Database Access:**
- Direct pyodbc connection strings
- Configuration loaded from JSON: `db_config.json`
- Connection parameters from file (not environment variables)
- Example from `Tools/scripts/test_sp_version.py`:
```python
cfg_path = os.path.join(os.path.dirname(__file__), 'config', 'db_config.json')
with open(cfg_path) as f:
    cfg = json.load(f)

conn_str = (
    f"DRIVER={{{cfg['driver']}}};"
    f"SERVER={cfg['server']};"
    f"DATABASE={cfg['database']};"
    f"UID={cfg['username']};"
    f"PWD={cfg['password']};"
    "TrustServerCertificate=yes;"
)
conn = pyodbc.connect(conn_str, autocommit=False)
```

## Fixtures and Factories

**Test Data:**
- Constants at module level: `DEFAULTS = {...}`, `NUMERIC_COL_PATTERNS = [...]`
- Dictionaries for test rows: `from_excel = {...}`
- Hardcoded SQL parameters passed directly to `execute()`
- No reusable data factories

**Location:**
- In-file constants and functions
- No separate fixtures directory
- Shared constants (e.g., defaults) defined at top of test file

**Example from `Tools/scripts/test_insert_bom.py`:**
```python
DEFAULTS = {
    "ParentId"          : -1,
    "IsGroup"           : 0,
    "BranchCode"        : "A01",
    "DocCode"           : "BOM",
    "DocStatus"         : 4,
    # ... 30 more fields
}

def main():
    from_excel = {
        "Code"          : "TEST-BOM-001",
        "DocDate"       : datetime.date(2025, 9, 5),
        "ItemId0"       : 19475,
        # ... more fields
    }
    row = {**DEFAULTS, **from_excel}
```

## Coverage

**Requirements:** No coverage targets or tracking

**What's Tested:**
- Database connectivity (connection string building)
- SQL INSERT execution (BOM row insertion)
- Stored procedure calls (usp_B20BOM_AutoVersion)
- User authentication (hash verification)
- Basic Excel file reading (docx parsing)

**What's Not Tested:**
- No unit tests for business logic (parsing, validation)
- No integration tests for full workflows
- No error scenario testing
- No regression test suite
- No automated regression detection

## Test Types

**Integration Tests (actual tests):**
- `test_insert_bom.py` — End-to-end BOM insertion
  - Setup: Load config, connect to DB, prepare test data
  - Action: INSERT SQL, commit transaction
  - Verify: Read back inserted row, check column values
  - Scope: Real database, real table (`B20BOM`)

- `test_sp_version.py` — Stored procedure testing
  - Setup: Load config, connect to DB, hardcode item ID
  - Action: Execute stored procedure `usp_B20BOM_AutoVersion` with 4 different parameter patterns
  - Verify: Check return value, print results
  - Scope: Real database, real stored procedure

- `testUser.py` — User authentication hash testing
  - Setup: Hardcode username, password, connection string
  - Action: Query user hash from DB, compute hashes with different encodings
  - Verify: Compare computed hashes to DB values
  - Scope: Real database, real `SYS_User` table

**Utility Scripts (not tests):**
- `create_test_tables.sql` — Database setup (DDL, not a test)
- `dump_mapping.py` — Export mapping config to JSON (data export, not a test)
- `patch_mapping.py` — Modify mapping in Excel (data modification, not a test)
- `read_docx.py` — Parse DOCX documents (data parsing, not a test)
- `update_mapping_cols.py` — Sync mapping to schema (data sync, not a test)

**E2E Tests:**
- No formal E2E test framework (Selenium, Cypress, etc.)
- Manual UI testing only — no automated browser/app testing
- Tests focus on data layer, not UI

## Common Testing Patterns

**Database Connection Pattern:**
```python
# Pattern 1: Direct pyodbc connection
import pyodbc, json
with open(config_path) as f:
    cfg = json.load(f)
conn_str = (
    f"DRIVER={{{cfg['driver']}}};"
    f"SERVER={cfg['server']};"
    f"DATABASE={cfg['database']};"
    f"UID={cfg['username']};"
    f"PWD={cfg['password']};"
    "TrustServerCertificate=yes;"
)
conn = pyodbc.connect(conn_str, timeout=10)
```

**SQL Execution Pattern:**
```python
# Pattern 1: INSERT with OUTPUT
sql = (
    f"INSERT INTO B20BOM ({', '.join(f'[{c}]' for c in cols)}) "
    f"OUTPUT INSERTED.Id "
    f"VALUES ({', '.join(['?']*len(cols))})"
)
cur = conn.cursor()
cur.execute(sql, values)
new_id = cur.fetchone()[0]
conn.commit()

# Pattern 2: Stored Procedure
sql = f"EXEC {SP} @_ItemId0=?, @_ParentBizDocId=NULL, @_TypeB=NULL"
cur.execute(sql, [item_id])
r = cur.fetchone()

# Pattern 3: SELECT with parameter
cur.execute("SELECT ... FROM table WHERE Id = ?", new_id)
r = cur.fetchone()
```

**Error Handling Pattern:**
```python
try:
    # Execute operation
    cur.execute(sql, values)
    result = cur.fetchone()[0]
    conn.commit()
    print(f"✅  Success — {result}")
except Exception as e:
    conn.rollback()
    print(f"❌  ERROR:\n{e}")
    sys.exit(1)
```

**Status Printing Pattern:**
```python
print("=== Test 1: Description ===")
print("Connecting...", end=" ")
conn = pyodbc.connect(conn_str)
print("OK")
print()

print(f"Inserted: Id={new_id}")
print(f"  Code: {row['Code']}")
print(f"  Date: {row['DocDate']}")
print()
print("Verifying...")
print(f"  Verified: {verified_data}")
```

## Manual Testing Notes

**Running Individual Tests:**
```bash
# From Tools directory
python scripts/test_insert_bom.py
python scripts/test_sp_version.py
python scripts/testUser.py
```

**Prerequisites:**
- `Tools/config/db_config.json` configured with valid server, database, credentials
- SQL Server accessible and online
- Required tables exist in database (`B20BOM`, `SYS_User`, etc.)
- Python 3.10+ with required packages: `pyodbc`, `pandas`, `openpyxl`

**Test Data Requirements:**
- For `test_insert_bom.py`: Valid `ItemId0` that exists in `B20Item` table
- For `test_sp_version.py`: Valid item ID for testing SP parameter patterns
- For `testUser.py`: Valid username and password for test user account

## Known Testing Gaps

**Not Covered:**
- Business logic layer functions (`_parse_sheet`, `_build_headers`, `_extract_meta`, etc.)
- Data validation functions (all validators.py functions)
- String utilities (`_norm_vn`, `guess_col_align`, etc.)
- UI components and event handlers
- Excel parsing edge cases (encrypted files, malformed sheets, etc.)
- Formula resolution and merged cells
- Fill-forward logic
- Section detection and roman numeral matching
- Mapping loading and column matching
- Error scenarios (network failures, corrupted data, etc.)

**Manual Testing Required For:**
- Full end-to-end import workflow (select file → parse → validate → insert)
- UI theme switching and persistence
- Tab navigation and state management
- Database configuration and connection validation
- Excel file password protection handling
- Large file performance (1000+ row sheets)
- Concurrent operations
- Rollback scenarios

---

*Testing analysis: 2026-08-03*
