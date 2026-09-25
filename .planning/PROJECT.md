# BOHO Import Tool — Refactor BOM Logic

## What This Is

Desktop app Python/CustomTkinter dùng để import dữ liệu BOM và THDM từ file Excel (.xlsm) vào hệ thống ERP Bravo (SQL Server), phục vụ nội bộ phòng SX/kế hoạch. Đang chạy production ở v2.2.25 (94 file BOM mẫu thật đã được dùng để audit).

## Core Value

Dữ liệu BOM import vào Bravo phải **đúng và nhất quán** — sai lệch (sai mã vật tư, sai số lượng, sai người tạo, sai liên kết dòng) gây hậu quả trực tiếp lên kế hoạch sản xuất và dữ liệu ERP thật.

## Requirements

### Validated

- ✓ Parse file Excel BOM (.xlsm) nhiều section (BOM2-BOM5), detect header 2 dòng, fill-forward, BTP detection — hoạt động qua 94 file mẫu thật (audit trong phiên trước), không crash.
- ✓ EmployeeId lấy động theo người dùng chọn trên UI (không còn hard-code) — xác minh trong code hiện tại (`_current_creator_employee_id`).
- ✓ ParentDetailRowId_SO tính đúng theo Mục số (không còn gán nhầm mã đơn hàng) — đã fix ở v2.2.23.

### Active

- [ ] Refactor logic xử lý **BOM** (không đụng THDM/phần khác) sang kiến trúc Staging First: tách `core/` (pure function, test được bằng pytest) khỏi `ui/` (CustomTkinter).
- [ ] Rà soát từng logic nghiệp vụ BOM hiện có (BA-style) → quyết định GIỮ / SỬA / BỎ trước khi viết lại, dựa trên bằng chứng từ 94 file mẫu thật.
- [ ] Sửa `_mkt_cache` dict-collision (MKT fallback ItemId) — query thiếu `ORDER BY`, dict ghi đè không xác định khi 1 `ItemTypeSX_Parent` có nhiều dòng trong `vB20Item_MKT`.
- [ ] Sửa `_build_headers()` tạo cột rác khi có vùng tính toán ẩn nằm cạnh bảng thật trong Excel (forward-fill label không giới hạn).
- [ ] Rà soát toàn bộ pattern `(x or '').strip()` trong luồng đọc dữ liệu BOM (Excel lẫn DB) — gói vào `core/sanitizer.py` để chặn tận gốc lớp lỗi crash `.strip()`.
- [ ] Thiết lập `core/bravo_staging.py`: Insert BOM luôn qua bảng Staging trên SQL Server → Validate → Transfer qua Transaction, không insert thẳng vào bảng nghiệp vụ Bravo.
- [ ] Logging chuẩn (`logging` module ghi `logs/app.log`) cho toàn bộ luồng BOM mới — không còn `except` nuốt lỗi không log.

### Out of Scope

- THDM logic (`_thdm_*` trong `bom_parser.py`, phần import THDM trong `main_window.py`) — không đụng trong lần refactor này, giữ nguyên hành vi hiện tại.
- Rewrite toàn bộ UI CustomTkinter — chỉ đổi cách `ui/` gọi vào `core/` mới, không thiết kế lại giao diện.
- Cutover thay thế bản production ngay — làm song song trong `v3/`, không tắt/đổi bản đang chạy cho tới khi đối chiếu output đầy đủ.

## Context

- Toàn bộ logic hiện dồn trong 1 file `Tools/views/main_window.py` (11,133 dòng, 270 hàm, 229 khối `except`, **0 lần gọi `logger`**) — nguyên nhân gốc của chuỗi bug vá vụn liên tiếp (v2.2.21 → v2.2.25).
- `Tools/services/bom_parser.py` (1382 dòng) đã tách được phần parse Excel thuần (pure-ish), nhưng phần resolve/insert row (BTP detection, MKT fallback, SP_HOOK) vẫn nằm trong `main_window.py`.
- Đã audit 94 file BOM mẫu thật (`v3/sample_data/bom_import/`, kết quả `v3/docs/audit/bom_sample_scan.json`) bằng cách chạy trực tiếp `parse_bom_file()` — không crash, nhưng phát hiện: `_mkt_cache` dict-collision (chưa fix), cột header rác ở ~11% file (10/94), STT trộn kiểu int/float/str (nguồn gốc lớp lỗi `.strip()` crash).
- Rebuild thực hiện trong `v3/` (copy source từ `Tools/`, loại bỏ build artifact/log/backup cũ) — không đụng bản gốc `Tools/` đang chạy production.

## Constraints

- **DB Safety**: Không DROP/TRUNCATE; mọi UPDATE/DELETE phải preview SQL trước khi chạy qua MCP.
- **Tech stack**: Python + CustomTkinter (giữ nguyên, không đổi framework UI) + pyodbc/SQL Server (Bravo ERP).
- **Không downtime**: bản `Tools/` gốc phải tiếp tục chạy được cho người dùng cuối trong suốt quá trình refactor.
- **Phạm vi**: chỉ BOM — THDM và các phần khác giữ nguyên, không refactor kèm.

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Rebuild song song trong `v3/`, không sửa trực tiếp `Tools/` | Tool đang chạy production, cutover 1 lần rủi ro cao với dữ liệu ERP thật | ✓ Good |
| Chỉ refactor BOM, không đụng THDM | Người dùng yêu cầu thu hẹp phạm vi, giảm rủi ro regression | ✓ Good |
| Đóng vai BA rà soát logic (giữ/sửa/bỏ) dựa trên 94 file mẫu thật trước khi viết code core/ mới | Tránh rewrite mù — nhiều logic hiện tại (BTP detection nhiều tầng, Fill-Forward, MKT fallback) là các fix tích lũy từ bug thật, cần hiểu rõ trước khi giữ hay bỏ | — Pending |

## Evolution

This document evolves at phase transitions and milestone boundaries.

**After each phase transition** (via `/gsd-transition`):
1. Requirements invalidated? → Move to Out of Scope with reason
2. Requirements validated? → Move to Validated with phase reference
3. New requirements emerged? → Add to Active
4. Decisions to log? → Add to Key Decisions
5. "What This Is" still accurate? → Update if drifted

**After each milestone** (via `/gsd-complete-milestone`):
1. Full review of all sections
2. Core Value check — still the right priority?
3. Audit Out of Scope — reasons still valid?
4. Update Context with current state

---
*Last updated: 2026-09-25 after initialization*
