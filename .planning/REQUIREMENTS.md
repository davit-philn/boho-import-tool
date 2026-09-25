# Requirements: BOHO Import Tool — Refactor BOM Logic

**Defined:** 2026-09-25
**Core Value:** Dữ liệu BOM import vào Bravo phải đúng và nhất quán — sai lệch gây hậu quả trực tiếp lên kế hoạch sản xuất và dữ liệu ERP thật.

## v1 Requirements

Requirements cho đợt refactor BOM lần này. Mỗi mục map vào roadmap phase.

### BA-Analysis

- [x] **BA-01**: Toàn bộ logic xử lý dữ liệu BOM hiện có trong `bom_parser.py` + phần liên quan trong `main_window.py` (BTP detection, Fill-Forward, MKT fallback, SP_HOOK, section/header detection) được liệt kê đầy đủ kèm bằng chứng hành vi thực tế (từ 94 file mẫu đã audit).
- [ ] **BA-02**: Mỗi logic có nhận định rõ ràng: GIỮ NGUYÊN / SỬA (kèm lý do + cách sửa) / BỎ (kèm lý do) — được người dùng (vai trò business owner) xác nhận trước khi viết code `core/` mới.
- [x] **BA-03**: Các quyết định BA được ghi lại thành tài liệu tham chiếu cho giai đoạn viết code (không để trôi mất qua hội thoại).

### Core-Architecture

- [ ] **ARCH-01**: Logic nghiệp vụ BOM được tách thành pure function trong `core/`, test được độc lập bằng pytest, không phụ thuộc CustomTkinter.
- [ ] **ARCH-02**: `ui/` (hoặc `views/` hiện tại) chỉ gọi vào `core/` mới, không chứa logic nghiệp vụ.
- [ ] **ARCH-03**: Insert BOM vào Bravo luôn đi qua Staging Table trên SQL Server → Validate → Transfer qua Transaction — không insert thẳng vào bảng nghiệp vụ Bravo.
- [ ] **ARCH-04**: `core/sanitizer.py` cung cấp `safe_str`, `safe_int`, `safe_decimal`, `safe_date` — dùng thống nhất ở mọi nơi đọc dữ liệu thô (Excel lẫn kết quả query DB) trong luồng BOM.
- [ ] **ARCH-05**: Bản `Tools/` gốc (production) không bị chỉnh sửa trong suốt quá trình — làm song song trong `v3/`, đối chiếu output trước khi cân nhắc cutover.

### Bug-Fixes

- [ ] **FIX-01**: `_mkt_cache` không còn ghi đè không xác định khi 1 `ItemTypeSX_Parent` có nhiều dòng trong `vB20Item_MKT` (thêm `ORDER BY` + tiêu chí ưu tiên rõ ràng, hoặc dedupe ở tầng SQL).
- [ ] **FIX-02**: `_build_headers()` không còn forward-fill label vô hạn sang các ô tính toán/ẩn nằm ngoài vùng bảng thật — không sinh cột tên rác trong DataFrame preview.
- [ ] **FIX-03**: Toàn bộ pattern `(x or '').strip()` trong luồng đọc dữ liệu BOM (Excel + DB) được thay bằng `sanitizer.safe_str()` — chặn tận gốc lớp lỗi crash `.strip()` trên giá trị numeric.
- [ ] **FIX-04**: Warning "Không tìm thấy sheet BOMx" chỉ hiển thị cho section bắt buộc — section tùy chọn (vd BOM5) không được cảnh báo như lỗi.

### Observability

- [ ] **OBS-01**: Toàn bộ exception trong luồng BOM mới được `logger.exception()` ghi vào `logs/app.log` — không còn `except` nuốt lỗi không log.
- [ ] **OBS-02**: Log đủ thông tin để truy vết: file Excel nguồn, section, dòng dữ liệu, lỗi cụ thể.

## v2 Requirements

Ghi nhận nhưng chưa đưa vào roadmap lần này.

### THDM-Refactor

- **THDM-01**: Áp dụng kiến trúc Staging First tương tự cho luồng THDM.
- **THDM-02**: Rà soát BA cho logic THDM (expand Mục, BOM qty lookup) như đã làm với BOM.

### Cutover

- **CUT-01**: Thay thế bản `Tools/` gốc bằng `v3/` sau khi đối chiếu output đầy đủ trên dữ liệu thật.

## Out of Scope

| Feature | Reason |
|---------|--------|
| Refactor logic THDM | Người dùng yêu cầu thu hẹp phạm vi lần này, chỉ BOM |
| Rewrite UI CustomTkinter | Chỉ đổi cách `ui/` gọi vào `core/`, không thiết kế lại giao diện |
| Cutover thay bản production ngay | Rủi ro cao với dữ liệu ERP thật, cần đối chiếu output trước |
| Đổi tech stack (Python/CustomTkinter/pyodbc) | Không có lý do kỹ thuật để đổi, giữ nguyên |

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| BA-01, BA-02, BA-03 | Phase 1 | Pending |
| ARCH-01 .. ARCH-05 | Phase 2 | Pending |
| FIX-01 .. FIX-04 | Phase 2 | Pending |
| OBS-01, OBS-02 | Phase 2 | Pending |

**Coverage:**

- v1 requirements: 14 total
- Mapped to phases: 14
- Unmapped: 0 ✓

---
*Requirements defined: 2026-09-25*
*Last updated: 2026-09-25 after initial definition*
