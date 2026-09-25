# Phase 1: BA Analysis & Decision Lock - Context

**Gathered:** 2026-09-25
**Status:** Ready for planning

<domain>
## Phase Boundary

Kiểm kê toàn bộ logic xử lý dữ liệu **BOM** hiện có (không THDM) — section/header detection, Fill-Forward, BTP detection, MKT fallback, SP_HOOK, resolve row theo NguonDL, EmployeeId/ParentDetailRowId_SO, sanitize `.strip()` — xuống tới từng nhánh điều kiện/rule con. Mỗi logic nhận một verdict **GIỮ / SỬA / BỎ** có bằng chứng thật kèm theo. Không viết code `core/` trong phase này — chỉ tạo tài liệu quyết định làm input cho Phase 2.

</domain>

<decisions>
## Implementation Decisions

### Độ sâu kiểm kê
- **D-01:** Kiểm kê xuống mức từng nhánh điều kiện/rule con (vd Rule 1 "tên rỗng + có con/Tên chi tiết" và Rule 2 "chứa dấu +" trong BTP detection đánh giá riêng biệt; từng nguồn `NguonDL` — Excel/CoDinh/HeThong/UILookup/MucLookup — trong `_resolve_row_mapping`/`_resolve_detail_row` đánh giá riêng), không gộp ở mức hàm.

### Tiêu chí chốt GIỮ/SỬA/BỎ
- **D-02:** Mọi verdict SỬA/BỎ bắt buộc kèm bằng chứng cụ thể — trích dẫn được từ 94 file mẫu (`v3/sample_data/bom_import/`) hoặc dữ liệu SQL thật.
- **D-03:** Với rủi ro cấu trúc không thể chứng minh bằng file Excel mẫu (vd query thiếu `ORDER BY`, pattern `except` không log) — chấp nhận bằng chứng dạng đọc code + đối chiếu schema/dữ liệu thật qua `mssql-boho` MCP (vd query thật `vB20Item_MKT` để xác nhận có nhiều dòng trùng `ItemTypeSX_Parent` hay không). Không chấp nhận suy luận thuần từ code không đối chiếu gì.

### Nguồn bổ sung bằng chứng
- **D-04:** Đọc lại lịch sử commit/changelog từ v2.2.1 → v2.2.25 (`git log`, `version.json` release notes) để hiểu context từng fix đã vá trước khi quyết định — tránh đề xuất SỬA một chỗ đã từng bị fix rồi revert/đổi hướng.

### Hình thức tài liệu quyết định
- **D-05:** Output là bảng chi tiết 5 cột: `Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro` — mỗi dòng là 1 logic hoặc 1 rule con, đủ chi tiết để Phase 2 lập plan trực tiếp không cần hỏi lại.

### Claude's Discretion
- Thứ tự trình bày các logic trong bảng (theo file nguồn hay theo mức rủi ro) — để Claude tự sắp xếp sao cho dễ đọc nhất khi viết tài liệu.

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Codebase hiện tại (logic BOM cần kiểm kê)
- `v3/Tools/services/bom_parser.py` — parser Excel thuần (section/header detection, Fill-Forward, `_parse_section_excel_rows`, `_resolve_row_mapping`)
- `v3/Tools/views/main_window.py` lines 9127-9250+ (`_resolve_detail_row`) — resolve từng field BOM row theo NguonDL
- `v3/Tools/views/main_window.py` lines 9490-9890 (`_build_bom_detail_caches`/BTP detection/`_mkt_cache`/SP_HOOK insert loop) — vùng trọng tâm cần rà soát
- `v3/Tools/views/main_window.py` lines 4780-4820, 8940-8945 (`_current_creator_employee_id`) — EmployeeId resolve
- `v3/Tools/services/mapping_loader.py` — cấu trúc mapping (`_CONFIG`, `SP_HOOK`, `SP_CONFIG`) mà logic BOM phụ thuộc vào

### Bằng chứng đã thu thập (phiên trước)
- `v3/docs/audit/bom_sample_scan.json` — kết quả quét 94 file BOM mẫu thật bằng `parse_bom_file()` (không crash, có warning/mixed-dtype/header rác)
- `.planning/codebase/CONCERNS.md` — audit codebase độc lập 2026-08-03, xác nhận trùng 2 điểm fragile ("BOM Parser Row Mapping Logic", "Sheet Type Detection") + rủi ro No Rollback/Transaction, 416 exception handler không log
- `.planning/codebase/ARCHITECTURE.md`, `STRUCTURE.md` — cấu trúc tổng thể codebase (đọc nếu cần thêm bối cảnh module hóa)

### Nguồn cần đọc thêm trong Phase 1 (theo quyết định D-04)
- `git log --oneline` (v2.2.1 → v2.2.25) và nội dung `version.json` qua các lần release — lịch sử fix
- Query trực tiếp `vB20Item_MKT` qua `mssql-boho` MCP để xác minh giả thuyết dict-collision bằng dữ liệu thật (theo D-03)

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `bom_parser.py` đã tách phần lớn parser Excel ra khỏi UI (không phụ thuộc CustomTkinter trực tiếp, trừ `_ask_excel_password`) — có thể tái dùng gần như nguyên vẹn khi chuyển vào `core/excel_parser.py` ở Phase 2, sau khi đã chốt GIỮ/SỬA/BỎ từng phần.
- 94 file mẫu trong `v3/sample_data/bom_import/` và script audit trong scratchpad đã có sẵn — dùng lại để verify từng verdict thay vì viết script mới từ đầu.

### Established Patterns
- Mapping-driven: hầu hết logic BOM đọc cấu hình từ `CK_Mapping_v5.xlsx` (`_CONFIG`, `SP_HOOK`, `SP_CONFIG`) thay vì hardcode — cần giữ pattern này khi viết lại `core/`.
- `NguonDL` (Excel/CoDinh/HeThong/UILookup/MucLookup/SP/TinhToan) là trục phân loại chính cho mọi field resolve — cả BOM và THDM dùng chung qua `_resolve_row_mapping()`.

### Integration Points
- `_resolve_detail_row` (BOM) và `_resolve_row_mapping` (dùng chung BOM+THDM) là ranh giới rõ nhất giữa "logic thuần" và "logic phụ thuộc UI/conn" — điểm tách tự nhiên cho `core/` vs phần cần `conn`/`self`.

</code_context>

<specifics>
## Specific Ideas

Không có yêu cầu hình thức cụ thể khác ngoài bảng 5 cột đã chốt ở D-05.

</specifics>

<deferred>
## Deferred Ideas

- THDM logic (`_thdm_*`) — ngoài phạm vi phase này, đã ghi trong PROJECT.md Out of Scope.
- Cutover thay bản `Tools/` gốc — ngoài phạm vi, chỉ làm sau khi đối chiếu output đầy đủ (PROJECT.md Out of Scope).

### None khác — thảo luận không đi lệch phạm vi phase.

</deferred>

---
*Phase: 1-BA Analysis & Decision Lock*
*Context gathered: 2026-09-25*
