# Coding Conventions

**Analysis Date:** 2026-08-03

## Naming Patterns

**Files:**
- Lowercase with underscores: `bom_parser.py`, `mapping_loader.py`, `main_window.py`
- Test files in `scripts/` directory: `test_insert_bom.py`, `test_sp_version.py`
- Convention: `module_name.py` for utility modules, `*_window.py` for UI modules

**Functions:**
- Snake case for all functions: `parse_bom_file()`, `_load_sheet_config()`, `_build_headers()`
- Private/internal functions prefixed with single underscore: `_is_roman_numeral()`, `_resolve_formula()`, `_extract_meta()`
- Examples from `Tools/services/bom_parser.py`:
  ```python
  def _is_roman_numeral(s: str) -> bool:
  def _merge_meta_rows(live_rows, cached_rows):
  def _build_excel_col_map(headers, mapping_recs, col_aliases=None):
  ```

**Variables:**
- Snake case for all variables: `current_ff`, `row_text`, `stt_str`, `hidden_rows`, `stt_idx`
- Local variables lowercase: `df`, `sql`, `conn`, `rows_out`

**Classes:**
- PascalCase for all classes: `BOMToolApp`, `CLabel`, `CButton`, `_SearchCombo`, `SheetTable`
- UI components inherit from CustomTkinter base classes: `class CLabel(ctk.CTkLabel):`
- Example from `Tools/views/main_window.py`:
  ```python
  class BOMToolApp(ctk.CTk):
      def __init__(self):
  ```

**Constants:**
- UPPER_SNAKE_CASE for module-level constants: `APP_VERSION`, `BASE_DIR`, `CONFIG_DIR`, `MAPPING_FILE`, `DB_CONFIG_FILE`, `FOOTER_KEYWORDS`, `DEFAULT_CREATOR_USER_ID`
- Example from `Tools/services/utils.py`:
  ```python
  APP_VERSION = "2.1"
  TAB_IMPORT  = "📥  Import BOM"
  TAB_THDM    = "📊  Tổng hợp định mức"
  TAB_CATALOG = "🌳  Danh mục vật tư"
  DEFAULT_CREATOR_USER_ID = 823
  ```

## Code Style

**Formatting:**
- No automated formatter detected
- Consistent spacing around operators
- 4-space indentation
- Multiline imports shown with parentheses:
  ```python
  from services.mapping_loader import (
      load_mapping, build_reverse_map, match_col_to_sql,
      build_meta_keys_from_mapping, _load_section_rows,
  )
  ```

**Linting:**
- No linting configuration (.flake8, .pylintrc, pyproject.toml)
- Manual style adherence to conventions observed

**Documentation Style:**
- Module-level docstring using triple quotes:
  ```python
  """
  services/bom_parser.py — Parse BOM Excel, THDM sheets, shared row resolver.
  """
  ```
- Function docstrings with Vietnamese explanation:
  ```python
  def _is_roman_numeral(s: str) -> bool:
      """True nếu s là số La Mã thực dùng trong BOM (I–XXXIX, chỉ dùng I/V/X).
      Loại trừ chữ cái đơn A–Z làm section header (A,B,C,D không khớp pattern này)."""
  ```

## Import Organization

**Order:**
1. Standard library imports: `re`, `os`, `sys`, `datetime`, `unicodedata`, `io`, `math`
2. Third-party imports: `pandas`, `openpyxl`, `customtkinter`, `pyodbc`, `tkinter`, `rapidfuzz`
3. Conditional third-party (optional with try/except): `msoffcrypto`, `tksheet`
4. Local imports with relative module paths: `from services.utils import`, `from services.mapping_loader import`, `from views.widgets import`

**Example from `Tools/services/bom_parser.py`:**
```python
import re, os, sys, datetime, unicodedata, io, math
import pandas as pd
from openpyxl import load_workbook

try:
    import msoffcrypto
    _HAS_MSOFFCRYPTO = True
except ImportError:
    _HAS_MSOFFCRYPTO = False

from services.utils import _norm_vn, _nan_str, guess_col_align
from services.mapping_loader import (
    load_mapping, build_reverse_map, match_col_to_sql,
)
```

**Path Aliases:**
- No path aliases configured
- Relative imports from local modules: `from services.X import`, `from views.X import`
- All imports assume `Tools/` is in sys.path (see `Tools/main.py` line 10)

## Type Hints

**Function Signatures:**
- Type hints used for function parameters and return types
- Examples from codebase:
  ```python
  def _is_roman_numeral(s: str) -> bool:
  def guess_col_align(col_name: str) -> str:
  def _extract_meta(rows, meta_keys=None):
  def _build_excel_col_map(headers, mapping_recs, col_aliases=None):
  ```
- Python 3.10+ union types with `|` operator used in some cases:
  ```python
  self._popup: tk.Toplevel | None = None
  self._listbox: tk.Listbox | None = None
  ```

## Error Handling

**Patterns:**
- Try/except for optional dependencies with feature flags:
  ```python
  try:
      import msoffcrypto
      _HAS_MSOFFCRYPTO = True
  except ImportError:
      _HAS_MSOFFCRYPTO = False
  ```

- Generic exception catching during initialization:
  ```python
  except Exception:
      return []
  ```

- Specific exception handling in business logic:
  ```python
  except (ValueError, TypeError):
      return False
  ```

- Return None or empty defaults on error (no raise):
  ```python
  def _load_sheet_config(mapping_path=None):
      try:
          # ... logic
          return cfg
      except Exception:
          return []
  ```

- Database transaction rollback pattern:
  ```python
  try:
      cur.execute(sql, values)
      conn.commit()
  except Exception as e:
      conn.rollback()
      print(f"❌  LỖI:\n{e}")
      sys.exit(1)
  ```

## Logging

**Framework:** Console output via `print()` — no dedicated logging framework

**Patterns:**
- Status messages printed during operations: `print(f"Kết nối {cfg['server']} / {cfg['database']}...", end=" ")`
- Error messages prefixed with emoji indicators: `❌  LỖI:`, `✅  INSERT thành công`
- Progress indicators: `print("OK\n")`
- Debug output with structured formatting: `print(f"  Code    : {row['Code']}")`

**When to Log:**
- When entering major processing steps
- Database connection status
- Error conditions
- Operation completion status
- File I/O operations

Example from `Tools/scripts/test_insert_bom.py`:
```python
print(f"Kết nối {cfg['server']} / {cfg['database']}...", end=" ")
conn = pyodbc.connect(make_conn_str(cfg), timeout=10)
print("OK\n")
```

## Comments

**When to Comment:**
- Complex logic requiring explanation: `# Forward-fill r1 để xử lý merged cell`
- Vietnamese comments for domain-specific concepts
- Section dividers for major logical blocks: `# ─── Shared Excel parse pipeline ───`
- Explain "why" not "what": `# Label cell phải ngắn (<=40 ký tự) — loại instruction text dài`
- Flag edge cases and workarounds: `# Công thức phức tạp không resolve được`

**Comment Style:**
- Inline comments on same line or line before code
- Vietnamese language (matches codebase)
- Section headers use Unicode box-drawing: `# ─────────────────────────────────────`
- Numbered guard explanations:
  ```python
  # Hai guard chống false-positive:
  # 1. Label cell phải ngắn (<=40 ký tự)
  # 2. Quét giá trị tối đa 8 cột sang phải
  ```

## Function Design

**Size:** 
- Small utility functions: 3-15 lines (e.g., `_is_roman_numeral`, `_resolve_formula`)
- Medium processing functions: 20-80 lines (e.g., `_extract_meta`, `_build_headers`)
- Large orchestration functions: 100-300 lines (e.g., `_parse_section_excel_rows`, `parse_bom_file`)
- Keep complexity low — break into helper functions when exceeding 100 lines

**Parameters:**
- Named parameters for complex functions
- Optional parameters with default None: `meta_keys=None`, `sheet_config=None`
- Avoid long parameter lists (max 5-6 parameters for public functions)
- Example from `Tools/services/bom_parser.py`:
  ```python
  def _parse_section_excel_rows(all_rows, headers, col_map, mapping_recs, stt_idx,
                                 skip_no_item=True, item_cols=('ItemId', 'ItemName'),
                                 ff_always_override=False, extra_col_extractor=None,
                                 roman_as_hdr=False):
  ```

**Return Values:**
- Return data structures (dict, list, DataFrame) for results
- Return None on error in some paths (fallback patterns)
- Return empty collections ([], {}, DataFrame()) not None for no-data cases
- Example patterns:
  ```python
  # Pattern 1: Return dict for complex results
  return {"section": section, "label": label, "contains": contains}
  
  # Pattern 2: Return empty list on error
  except Exception:
      return []
  
  # Pattern 3: Return None for optional values
  return None
  
  # Pattern 4: Return DataFrame
  return pd.DataFrame(rows_out, columns=headers)
  ```

## Module Design

**Exports:**
- No `__all__` defined
- Public functions: no leading underscore
- Internal functions: leading underscore (`_is_roman_numeral`, `_load_sheet_config`)
- Callers explicitly import what they need

**Module Organization:**
```
Tools/
├── main.py                          # Entry point, theme initialization
├── services/                        # Business logic layer
│   ├── __init__.py                  # Empty marker
│   ├── bom_parser.py                # Excel parsing, row filtering, formula resolution
│   ├── mapping_loader.py            # Configuration loading from CK_Mapping_v5.xlsx
│   ├── utils.py                     # String normalization, path utilities, constants
│   └── validators.py                # Data validation layer 1 (types, required fields)
├── views/                           # UI layer (CustomTkinter)
│   ├── main_window.py               # Main application window, all tabs
│   └── widgets.py                   # Reusable UI components (CLabel, CButton, SearchCombo)
└── scripts/                         # Standalone test/utility scripts
    ├── test_insert_bom.py           # Integration test: BOM insertion
    ├── test_sp_version.py           # Stored procedure testing
    ├── testUser.py                  # User authentication testing
    └── create_test_tables.sql       # Database schema setup
```

**No Barrel Files:**
- No `index.py` or `__init__.py` with re-exports
- Each module imports directly what it needs from specific modules

**Module Responsibilities:**
- `bom_parser.py` (1700+ lines) — Excel reading, formula resolution, row filtering, section detection
- `mapping_loader.py` (400+ lines) — Config loading, reverse mapping, SQL column matching
- `validators.py` (300+ lines) — Data type validation, required field checking, date format validation
- `utils.py` (350+ lines) — String normalization, alignment guessing, path resolution, constants
- `main_window.py` (2000+ lines) — UI orchestration, event handling, tab management, DB operations
- `widgets.py` (500+ lines) — CustomTkinter wrapper classes, SearchCombo popup, SheetTable wrapper

## Special Patterns

**Fill-Forward Logic:**
- Section headers (A, B, C) capture values that propagate to data rows
- Implementation: `current_ff` dict updated on header detection, applied to subsequent rows
- Used in `Tools/services/bom_parser.py` for dimension and grouping fields

**Normalized Comparison:**
- `_norm_vn(s)` removes diacritics, converts to lowercase, removes non-alphanumeric
- Used for flexible column matching: Vietnamese `Dày` matches "Width", `Mã sản phẩm` matches "ItemCode"
- Regex patterns use `re.IGNORECASE` with `_norm_vn()` preprocessing

**Optional Feature Flags:**
- `_HAS_MSOFFCRYPTO`, `_HAS_TKSHEET` checked before importing
- Fallback UI component when optional library missing (e.g., Treeview instead of tksheet)
- Graceful degradation: feature works differently, not blocked

**Configuration Driven Parsing:**
- Section definitions loaded from `CK_Mapping_v5.xlsx` (not hardcoded)
- Sheet detection via regex word-boundary matching: `\b` patterns
- Mapping file is single source of truth for column definitions

---

*Convention analysis: 2026-08-03*
