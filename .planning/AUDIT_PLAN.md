# Audit Plan — BOHO IMPORT v2.1

**Ngày audit:** 2026-08-03  
**Công cụ:** ruff 0.16.1 · pytest 9.1.1 · manual review  
**Kết quả pytest hiện tại:** 22/22 passed ✅  
**Ruff issues tổng:** 170 issues

---

## Tóm tắt

| Nhóm | Loại | Số lượng | Mức độ |
|------|------|----------|--------|
| A | Lỗi thực (real bugs) | 28 issues | 🔴 Cao |
| B | Style tự fix được (auto-fix) | 129 issues | 🟡 Thấp |
| C | Test còn thiếu | 4 module/function groups | 🔴 Cao |
| D | Code chết / biến không dùng | 5 issues | 🟢 Thấp |

---

## Nhóm A — Lỗi thực (cần fix thủ công)

### A1 · `tk` chưa import trong `bom_parser.py` — F821 · 11 occurrences 🔴

**Vấn đề:** Hàm `_ask_excel_password()` dùng `tk.Toplevel()`, `tk.Label()`, `tk.Entry()`... nhưng `tkinter` không được import trong `bom_parser.py`. Hiện tại chỉ hoạt động vì `main_window.py` import `tkinter as tk` *trước* khi import `bom_parser`, khiến `tk` rò rỉ vào module scope — fragile và undefined behavior.

**Vị trí:** [`bom_parser.py:541,555,558,560,563,569,584,586,589,541`](Tools/services/bom_parser.py:541)  
**Fix đề xuất:** Thêm `import tkinter as tk` có guard ở đầu file, hoặc move `_ask_excel_password()` sang `main_window.py` nơi nó thuộc về.

---

### A2 · Bare `except:` — E722 · 17 occurrences 🔴

**Vấn đề:** Bare `except:` nuốt mọi exception kể cả `KeyboardInterrupt`, `SystemExit`. Bugs biến mất im lặng, không có log.

**Vị trí:**
- [`bom_parser.py:510`](Tools/services/bom_parser.py:510) — parsing workbook
- [`bom_parser.py:1239,1241,1253,1255`](Tools/services/bom_parser.py:1239) — row mapping (4 chỗ)
- [`main_window.py:1060,1071`](Tools/views/main_window.py:1060) — import handlers
- [`main_window.py:8191,8193`](Tools/views/main_window.py:8191) — Excel password dialog

**Fix đề xuất:** Thay bằng `except Exception as e: logger.warning(...)` — giữ hành vi catch-all nhưng có log. Không refactor rộng, chỉ thêm log và đổi keyword.

---

### A3 · f-string không có placeholder — F541 · 7 occurrences 🟠

**Vấn đề:** f-string nhưng không có `{variable}` nào bên trong — lãng phí, và thường là dấu hiệu đã xóa biến nhưng quên bỏ chữ `f`.

**Vị trí:** [`main_window.py:6298,6304,6305,6310,6312,6331,6334`](Tools/views/main_window.py:6298) — cụm liên tiếp trong 1 function  
**Fix đề xuất:** Xóa chữ `f` hoặc thêm lại placeholder đúng nếu thiếu biến.

---

## Nhóm B — Style (auto-fixable bằng `ruff check --fix`)

### B1 · Multiple statements on one line — E701/E702 · 111 occurrences 🟡

Ruff có thể tự sửa phần lớn. Chủ yếu là pattern `if x: y = 1` và `a = 1; b = 2` trên 1 dòng.

**Vị trí:** Rải rác khắp `bom_parser.py`, `main_window.py`, `scripts/`  
**Fix:** `ruff check Tools/ --fix --select E701,E702`

### B2 · Multiple imports on one line — E401 · 17 occurrences 🟡

`import re, os, sys, datetime, unicodedata, io, math` → tách thành dòng riêng.

**Vị trí:** Dòng đầu của `bom_parser.py`, `mapping_loader.py`, `utils.py`, `validators.py`, `main_window.py`, scripts  
**Fix:** `ruff check Tools/ --fix --select E401`

### B3 · No newline at end of file — W292 · 1 occurrence 🟢

**Vị trí:** [`scripts/testUser.py:31`](Tools/scripts/testUser.py:31)  
**Fix:** `ruff check Tools/ --fix --select W292`

---

## Nhóm C — Tests còn thiếu (viết mới)

Hiện tại chỉ có 22 tests, tất cả cho `bom_parser._parse_section_excel_rows` và regex patterns. Các module quan trọng sau **chưa có test nào:**

### C1 · `validators.py` — 🔴 ưu tiên cao

Module chịu trách nhiệm validate toàn bộ data trước khi INSERT vào SQL Server. Nếu validate sai → data lỗi vào DB không phát hiện được.

**Cần test:**
- `validate_bom_row()` — required fields, numeric fields, date format
- `validate_thdm_row()` — tương tự
- Edge cases: `None`, `NaN`, chuỗi rỗng, số âm, ngày sai format

### C2 · `utils.py` — hàm `_norm_vn()` và `_nan_str()` 🟠

Dùng ở khắp nơi để normalize tiếng Việt và xử lý NaN — lỗi ở đây ảnh hưởng toàn bộ pipeline.

**Cần test:**
- `_norm_vn()` với các ký tự có dấu, không dấu, uppercase/lowercase, ký tự đặc biệt
- `_nan_str()` với `None`, `float('nan')`, `pd.NA`, chuỗi thực

### C3 · `mapping_loader._detect_sheet_type()` 🟠

Hàm quan trọng nhưng fragile — xác định sheet nào là BOM I, BOM II, TH VT... Nếu sai → toàn bộ file bị parse sai section.

**Cần test:**
- Match đúng: `"BOM I"`, `"BOM_I"`, `"BOM IV"`, `"TH VT"`
- Không nhầm: `"BOM III"` không match `"BOM II"`
- Edge: uppercase, underscore, dấu cách

### C4 · `bom_parser._detect_sheet_type()` thêm edge cases 🟢

Bổ sung thêm cho test file hiện có:
- STT có trailing whitespace: `" 1 "`, `" H "`
- STT là `float`: `1.0` (openpyxl đôi khi trả số thực thay int)
- ItemName có mixed case: `"Kính Thủy Theo Mẫu"` (không all-caps → không bị skip)

---

## Nhóm D — Code chết / biến không dùng

| File | Dòng | Biến | Ghi chú |
|------|------|------|---------|
| `main_window.py` | 108 | `is_dark` | Gán giá trị nhưng không dùng |
| `main_window.py` | 2737 | `ALIGN_L` | Constant không dùng |
| `main_window.py` | 7232 | `err_msgs` | List build nhưng bỏ qua |
| `read_docx.py` | 10 | `texts` | Gán list nhưng không dùng |

**Fix:** Xóa hoặc dùng biến đó. Không cần refactor lớn.

---

## Thứ tự thực hiện đề xuất

```
Bước 1 (fix bugs trước)
  └─ A1: Thêm import tkinter guard vào bom_parser.py
  └─ A2: Thay bare except → except Exception + log (17 chỗ)
  └─ A3: Fix f-string không placeholder (7 chỗ)

Bước 2 (auto-fix style)
  └─ B1+B2+B3: ruff check --fix (111+17+1 = 129 issues tự sửa)

Bước 3 (thêm tests)
  └─ C1: tests/test_validators.py
  └─ C2: tests/test_utils.py
  └─ C3: tests/test_mapping_loader.py
  └─ C4: bổ sung edge cases vào test_bom_parser.py

Bước 4 (dọn code chết)
  └─ D: Xóa 4 biến unused
```

**Tổng ước tính:** ~3-4 giờ nếu làm tuần tự. Bước 1 & 3 cần đọc code, Bước 2 & 4 gần như tự động.

---

*Chờ duyệt trước khi thực hiện bất kỳ bước nào.*
