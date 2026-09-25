# Phase 1: BA Analysis & Decision Lock - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-09-25
**Phase:** 1-BA Analysis & Decision Lock
**Areas discussed:** Độ sâu kiểm kê logic, Tiêu chí chốt GIữ/SỮA/BỌ, Nguồn bổ sung bằng chứng, Hình thức tài liệu quyết định

---

## Độ sâu kiểm kê logic

| Option | Description | Selected |
|--------|-------------|----------|
| Mức hàm chính | ~6-8 mục lớn, mỗi mục 1 verdict, không đi sâu rule con | |
| Sâu từng nhánh điều kiện | Từng rule con (vd Rule 1/Rule 2 trong BTP detection, từng NguonDL) đánh giá riêng | ✓ |

**User's choice:** Sâu từng nhánh điều kiện.
**Notes:** Chi tiết hơn để dev dùng trực tiếp khi viết lại ở Phase 2.

---

## Tiêu chí chốt GIữ/SỮA/BỌ

| Option | Description | Selected |
|--------|-------------|----------|
| Bắt buộc bằng chứng từ dữ liệu thật | Mọi SỬA/BỎ phải trích dẫn được từ 94 file hoặc query SQL thật | ✓ |
| Cho phép code review thuần | Ghi nhận SỬA ngay khi thấy rủi ro logic, không cần minh họa | |

**User's choice:** Bắt buộc bằng chứng từ dữ liệu thật.

**Câu hỏi bổ sung:** Với rủi ro cấu trúc không thể chứng minh bằng Excel mẫu (vd thiếu ORDER BY, except không log) — loại bằng chứng nào chấp nhận?

| Option | Description | Selected |
|--------|-------------|----------|
| Đọc code + đối chiếu schema SQL thật | Dùng mssql-boho MCP query thật để xác nhận (vd vB20Item_MKT có trùng ItemTypeSX_Parent không) | ✓ |
| Chỉ cần đọc code | Không cần query DB thật | |

**User's choice:** Đọc code + đối chiếu schema/dữ liệu SQL thật qua mssql-boho MCP.

---

## Nguồn bổ sung bằng chứng

| Option | Description | Selected |
|--------|-------------|----------|
| Không cần thêm | 94 file + đọc code + query mssql-boho là đủ | |
| Log lỗi thật từ production | Đọc import_crash.log (nhắc đến trong commit v2.2.24) | |
| Lịch sử commit/changelog v2.2.x | Đọc lại toàn bộ commit message v2.2.1 → v2.2.25 | ✓ |

**User's choice:** Lịch sử commit/changelog v2.2.x.
**Notes:** Giúp hiểu context từng fix đã vá trước khi quyết định giữ/sửa.

---

## Hình thức tài liệu quyết định

| Option | Description | Selected |
|--------|-------------|----------|
| Bảng chi tiết 5 cột | logic \| verdict \| bằng chứng \| cách sửa \| risk | ✓ |
| Bảng đơn giản 3 cột | logic \| verdict \| lý do | |

**User's choice:** Bảng chi tiết 5 cột.

---

## Claude's Discretion

- Thứ tự trình bày các logic trong bảng quyết định (theo file nguồn hay theo mức rủi ro).

## Deferred Ideas

- THDM logic — ngoài phạm vi, đã có trong PROJECT.md Out of Scope.
- Cutover thay bản `Tools/` gốc — ngoài phạm vi, đã có trong PROJECT.md Out of Scope.
