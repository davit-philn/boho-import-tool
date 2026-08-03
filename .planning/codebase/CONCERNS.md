# Codebase Concerns

**Analysis Date:** 2026-08-03

## Tech Debt

**Monolithic GUI Layer:**
- Issue: `Tools/views/main_window.py` contains 9,667 lines of code in a single file. All UI, business logic, database operations, and data validation are mixed together with no separation of concerns.
- Files: `Tools/views/main_window.py`
- Impact: Extremely difficult to test individual features, maintain code, or reuse components. Any change risks breaking unrelated functionality. New developers cannot understand flow easily.
- Fix approach: Refactor into separate modules: (1) UI components layer (`views/`), (2) business logic layer (`controllers/`), (3) database layer (`repositories/`). Move data validation to `services/validators.py`, move database operations to dedicated `services/database.py`. Target: reduce main_window.py to <2000 lines containing only widget layout and event handlers.

**Large Complex Functions in bom_parser.py:**
- Issue: Multiple functions exceed 100+ lines with deep nesting and complex conditional logic:
  - `_thdm_parse_thvt_sheet()` — 151 lines
  - `parse_bom_file()` — 134 lines
  - `_parse_sheet()` — 128 lines
  - `_resolve_row_mapping()` — 115 lines
  - `_parse_section_excel_rows()` — 113 lines
- Files: `Tools/services/bom_parser.py`
- Impact: Hard to debug, test individual paths, and modify without introducing bugs. High cognitive load to understand logic flow.
- Fix approach: Extract sub-functions for each logical step (e.g., `_extract_sheet_metadata()`, `_validate_row_structure()`, `_map_row_to_sql_cols()`). Add type hints and docstrings documenting parameter contracts. Add unit tests for each sub-function before refactoring.

**Catch-All Exception Handlers:**
- Issue: 416 exception handlers in main_window.py, including bare `except: pass` statements that silently swallow errors (lines 1060, 1071, 8191, 8193, 8465, etc.). These mask bugs and make debugging impossible.
- Files: `Tools/views/main_window.py` (416 occurrences), `Tools/services/bom_parser.py` (41 occurrences)
- Impact: Silent failures, corrupted data imports, user confusion, no error visibility in logs. Bare except blocks even catch `KeyboardInterrupt` and `SystemExit`, preventing graceful shutdown.
- Fix approach: Replace all bare `except: pass` with specific exception types. Add logging before every catch. Use context managers for resource cleanup (`with` statements). Minimum example:
  ```python
  # Before (bad):
  except: pass
  
  # After (good):
  except (ValueError, KeyError) as e:
      logger.error(f"Failed to parse row: {e}")
      raise  # or handle with fallback
  ```

**No Automated Testing:**
- Issue: Only three manual test scripts exist (`Tools/scripts/test_insert_bom.py`, `test_sp_version.py`, `testUser.py`). No automated test suite, no CI/CD pipeline, no coverage reporting.
- Files: `Tools/scripts/` (manual only)
- Impact: Regressions go undetected. Each release is untested. Cannot safely refactor. Bugs discovered only in production.
- Fix approach: Set up pytest framework. Create test fixtures for common BOM formats. Write unit tests for parsers, validators, and database operations (target 70%+ coverage). Add pre-commit hooks to run tests before commit. Integrate with CI/CD.

## Known Bugs

**SQL Injection Vulnerability in Schema Detection:**
- Symptoms: Malformed table/column names in SQL error responses could break queries
- Files: `Tools/views/main_window.py` lines 4299, 4371; `Tools/views/main_window_backup_20260727c.py` lines 2437, 2509
- Trigger: When checking column existence or table permissions, column/table names are interpolated directly into SQL using f-strings:
  ```python
  cur.execute(f"SELECT TOP 0 [{col_name}] FROM {table}")  # UNSAFE
  ```
- Workaround: Currently mitigated by assumption that `col_name` and `table` come from schema metadata, but this is fragile if input validation is missing upstream.
- Fix approach: Use parameterized queries where possible. For schema queries that cannot be parameterized, validate column/table names against a whitelist before interpolation. Never allow user input directly into these queries.

**Database Connection Not Closed on Exception:**
- Symptoms: Connection leaks occur when exceptions happen between `cursor.execute()` and `conn.close()`. After many failed operations, connection pool exhausts.
- Files: `Tools/views/main_window.py` (throughout, e.g., lines 621, 648, 1762, 2195)
- Trigger: Long-running batch imports with network interruptions or constraint violations
- Workaround: Manually close connections in exception handlers
- Fix approach: Wrap all database operations in context managers:
  ```python
  with self._get_db_connection() as conn:
      cur = conn.cursor()
      # operations
  # conn.close() called automatically
  ```

## Security Considerations

**Credentials in Config Files:**
- Risk: `Tools/config/db_config.json` and `Other/acc.txt` contain database credentials and internal account info in plaintext
- Files: `Tools/config/db_config.json`, `Other/acc.txt`
- Current mitigation: Both files are in `.gitignore` and not committed to git
- Recommendations:
  1. Add `.env` file support using `python-dotenv` for runtime credential loading
  2. Encrypt stored credentials at rest (e.g., using `cryptography` library)
  3. Add access control: config files should have `600` permissions (owner-read-only)
  4. Add audit logging for all database operations (who, when, what)
  5. Rotate credentials monthly; enforce strong passwords
  6. Use SQL Server connection encryption (ENCRYPT=yes in connection string)

**No Input Validation for User-Supplied Queries:**
- Risk: THDM filter strings and other user inputs are not validated before SQL operations
- Files: `Tools/services/bom_parser.py` lines 1101-1134 (`_thdm_apply_row_filter()`)
- Current mitigation: Simple regex filtering on row numbers
- Recommendations: Implement strict input validation schema for all user-supplied parameters

**Unencrypted Excel File Handling:**
- Risk: Encrypted Excel files require password prompt in UI. If password is logged or cached unencrypted, it could leak.
- Files: `Tools/views/main_window.py` lines 8191-8193 (exception handling for password prompts)
- Recommendations: Use `msoffcrypto` library with careful password handling; clear password from memory after use

## Performance Bottlenecks

**No Database Connection Pooling:**
- Problem: Every operation creates a new database connection. For batch imports of 100+ rows, creates 100+ separate connections.
- Files: `Tools/views/main_window.py` throughout (each function creates its own `pyodbc.connect()`)
- Cause: Connection objects not reused or pooled. Each dialog and operation opens/closes independently.
- Improvement path: Implement connection pooling using `pyodbc.pooling` or a dedicated connection pool manager. Reuse single connection for entire import batch. Target: 50% reduction in connection overhead.

**Large Excel Files Loaded Into Memory Twice:**
- Problem: `_thdm_open_workbook()` (line 1086) loads workbook three times: once for data, once cached, once for hidden rows detection
- Files: `Tools/services/bom_parser.py` lines 1086-1100, `Tools/views/main_window.py` line 4863
- Cause: Each workbook load (`openpyxl.load_workbook()`) reads entire file into memory
- Improvement path: Load once with appropriate mode; cache result. Use streaming API for very large files (pandas `read_excel(chunksize=)`).

**DataFrame Operations in Memory for Validation:**
- Problem: Entire BOM sheet loaded into pandas DataFrame for validation before database import. With large BOMs (10K+ rows), memory usage spikes.
- Files: `Tools/services/validators.py` lines 35-198
- Cause: Full dataframe operations without chunking
- Improvement path: Process rows in chunks (1000 per batch). Stream validation results to UI instead of collecting all errors first.

**GUI Freezes During Long Imports:**
- Problem: Database imports run on main thread, blocking UI for seconds/minutes on large batches
- Files: `Tools/views/main_window.py` contains many database operations called directly from UI event handlers
- Cause: No threading/async separation between UI and I/O
- Improvement path: Move all database operations to worker threads. Use Qt Signal/Slot or callback patterns to update UI from worker threads. Add progress bar with cancellation support.

## Fragile Areas

**BOM Parser Row Mapping Logic:**
- Files: `Tools/services/bom_parser.py` lines 243-355, 1181-1275
- Why fragile: Complex nested conditionals for matching Excel columns to SQL columns. Row filtering with ROMAN_NUMERAL detection is brittle if BOM format changes. Multiple fallback paths make it hard to predict behavior.
- Safe modification: Add comprehensive test cases for each BOM format variant (BOM2, BOM3, etc.). Document expected input shapes. Add debug logging at each decision point. Refactor decision tree into explicit rules table.
- Test coverage: Only manual tests exist. Row mapping needs unit tests covering: (1) standard format, (2) missing columns, (3) extra columns, (4) non-ASCII characters in headers, (5) merged cells.

**Sheet Type Detection:**
- Files: `Tools/services/bom_parser.py` lines 70-96
- Why fragile: Regex matching on sheet names is case-sensitive and space-sensitive. Word boundary matching (`\b`) fails with underscores. If sheet naming convention changes (e.g., "BOM_V2" vs "BOM V2"), detection breaks.
- Safe modification: Centralize sheet naming conventions in configuration. Add validation against expected names. Log all attempted matches with debug info.
- Test coverage: No tests. Needs tests for: uppercase/lowercase variants, underscores/hyphens/spaces, non-ASCII characters, partial name matches.

**Excel Encrypted File Handling:**
- Files: `Tools/services/bom_parser.py` lines 612-635, `Tools/views/main_window.py` lines 8191-8193
- Why fragile: msoffcrypto library is optional (`try/except ImportError`). If missing, encrypted files silently fail with unclear error message. Password prompt in GUI may not work on all systems.
- Safe modification: Make msoffcrypto a required dependency or provide clear error message if not installed. Add tests with encrypted Excel files.
- Test coverage: No tests. Needs: encrypted file with standard password, encrypted file with special characters in password, detection of file encryption status.

**THDM Mục Expansion:**
- Files: `Tools/services/bom_parser.py` lines 855-934
- Why fragile: Complex logic merging BOM quantity dict with row expansion. Multiple hardcoded column references ("Mục", "Chi tiết", etc.). If source data structure changes, expansion silently produces wrong output.
- Safe modification: Add assertions checking expected columns exist before expansion. Log key expansion decisions. Add field-by-field comparison tests.
- Test coverage: Only manual tests. Needs: standard THDM format, THDM with missing detail rows, THDM with duplicate mục IDs.

## Scaling Limits

**Database Connection Timeout on Slow Networks:**
- Current capacity: 30-second default timeout in connection string
- Limit: Timeout exhausted on VPN connections or during network latency spikes. User must retry manually.
- Scaling path: Implement exponential backoff retry logic. Allow timeout to be configured per user. Add connection pooling so retries reuse existing connections. Target: support 100+ concurrent imports on high-latency networks.

**Excel Workbook Size Limit:**
- Current capacity: Tested up to 500 sheets, 10K rows per sheet
- Limit: Memory exhausted when loading very large consolidated BOMs (20K+ rows). No streaming.
- Scaling path: Implement streaming Excel read using `openpyxl.load_workbook(..., data_only=True)` with `iter_rows()`. Process rows one at a time instead of loading entire workbook.

**Batch Import Throughput:**
- Current capacity: ~100 rows/second on local network (depends on network latency and SQL Server load)
- Limit: User needs to wait 10+ minutes for large imports (1000+ rows). No parallelization.
- Scaling path: Implement parallel bulk insert using SQL Server `BULK INSERT` or `bcp` utility. Use thread pool for concurrent imports of multiple sections.

## Dependencies at Risk

**Optional Dependencies with Silent Fallback:**
- Risk: `tksheet` and `msoffcrypto` are optional but critical:
  - `tksheet`: UI degrades to slower `ttk.Treeview` if missing (hidden from user)
  - `msoffcrypto`: Encrypted Excel files cannot be opened if missing (error unclear)
- Impact: Features appear to work but silently downgrade functionality
- Migration plan: Make both required or provide clear user-facing error messages. Alternatively, add setup check on startup that warns if optional dependencies are missing.

**Outdated or Pinned Dependencies:**
- Risk: `requirements.txt` does not pin versions. New versions of `pandas`, `customtkinter`, `openpyxl`, `pyodbc` could introduce breaking changes.
- Files: `requirements.txt` lines 1-14
- Impact: Different users or CI environments get different versions. Bugs may be unreproducible.
- Migration plan: Use `pip freeze` to generate locked versions. Pin major+minor versions at minimum (e.g., `pandas>=1.5,<2.0`). Test upgrades in staging before deploying.

## Missing Critical Features

**No Undo/Rollback Capability:**
- Problem: Once data is imported to SQL Server, there is no undo. Large imports that fail midway leave database in partially-modified state.
- Blocks: Cannot safely recover from accidental bulk imports
- Fix approach: Implement transaction-based imports with automatic rollback on error. Add delete-by-BOM-ID function to reverse imports. Maintain audit trail of all imports.

**No Import Scheduling:**
- Problem: All imports are manual and require user to be at PC. No background/scheduled imports.
- Blocks: Cannot automate recurring imports from shared drives
- Fix approach: Add file watcher service or scheduled task integration. Support importing from folders on startup.

**No Email/Notification on Import Completion:**
- Problem: Long imports complete silently. User doesn't know if import succeeded or failed until manually checking database.
- Blocks: Cannot integrate with upstream workflows
- Fix approach: Send email notification on import success/failure with row count, duration, and any errors.

**No Data Lineage / Audit Trail:**
- Problem: No record of which user imported which file, when, with which version of tool
- Blocks: Cannot trace data provenance or debug production issues
- Fix approach: Log every import with user ID, timestamp, file hash, row count, and SQL Server session ID. Store in database for auditing.

## Test Coverage Gaps

**No Tests for Parser Edge Cases:**
- What's not tested: BOM files with merged cells, files with non-ASCII characters, files with formulas instead of values, files with hidden rows/columns, encrypted files, files with multiple header rows
- Files: `Tools/services/bom_parser.py` (all parsing functions)
- Risk: Regressions in real-world file handling go undetected
- Priority: High — parser is core to entire tool

**No Tests for Database Operations:**
- What's not tested: Connection failures, timeout handling, duplicate key errors, foreign key constraint violations, transaction rollback, concurrent imports
- Files: `Tools/views/main_window.py` (all database code)
- Risk: Silent data corruption or lost imports
- Priority: Critical — database is where data integrity matters most

**No Tests for Validation Rules:**
- What's not tested: Required field validation, numeric field validation, date format validation, string length validation against different SQL Server charsets
- Files: `Tools/services/validators.py`
- Risk: Invalid data slips through validation and causes SQL errors
- Priority: High — validation is the last line of defense

**No GUI/Integration Tests:**
- What's not tested: End-to-end workflows (select file → preview → import), theme switching, configuration changes, error dialog display and recovery
- Files: `Tools/views/main_window.py` (entire GUI layer)
- Risk: UI bugs discovered only by users
- Priority: Medium — can be addressed after unit test coverage improves

**No Performance/Load Tests:**
- What's not tested: Import throughput, memory usage with large files, connection pool saturation, timeout behavior under network stress
- Files: All modules
- Risk: Tool becomes unusable with real-world file sizes
- Priority: Medium — important for production deployment

---

*Concerns audit: 2026-08-03*
