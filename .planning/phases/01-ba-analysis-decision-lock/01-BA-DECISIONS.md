# Kiểm kê & Quyết định logic BOM (GIỮ / SỬA / BỎ)

**Trạng thái:** DRAFT

Tài liệu quyết định nghiệp vụ (BA) cho toàn bộ logic xử lý dữ liệu **BOM** hiện có trong `Tools/views/main_window.py` và `Tools/services/bom_parser.py` (production, không đụng THDM). Mỗi logic/rule con nhận 1 verdict **GIỮ / SỬA / BỎ** có bằng chứng thật kèm theo (94 file mẫu, SQL Server thật, lịch sử fix v2.2.1 → v2.2.25). Dùng làm input trực tiếp cho Phase 2 (viết `core/`), không cần hỏi lại BA.

## 0. Phạm vi, nguồn & quy ước

### 0.1 Baseline mã nguồn

Đã đối chiếu `v3/Tools/` với `Tools/` (production, HEAD tại commit `8a07639`, release v2.2.25) bằng:

```bash
git -C D:/project/insertdatasqlserver diff --no-index --ignore-cr-at-eol --stat -- Tools/services/bom_parser.py v3/Tools/services/bom_parser.py
git -C D:/project/insertdatasqlserver diff --no-index --ignore-cr-at-eol --stat -- Tools/views/main_window.py v3/Tools/views/main_window.py
git -C D:/project/insertdatasqlserver diff --no-index --ignore-cr-at-eol --stat -- Tools/services/mapping_loader.py v3/Tools/services/mapping_loader.py
```

Kết quả: không có diff nội dung nào cho cả 3 file (chỉ cảnh báo CRLF/LF, không có khối thay đổi). Xác nhận `v3/Tools/services/bom_parser.py`, `v3/Tools/views/main_window.py`, `v3/Tools/services/mapping_loader.py` khớp 100% nội dung với `Tools/` production tại commit `8a07639` (release v2.2.25).

**Mọi số dòng trích dẫn trong tài liệu này trỏ vào `v3/Tools/...`** (không phải `Tools/...`), vì `v3/` là bản đang dùng để đọc/refactor, nhưng nội dung là một.

### 0.2 Quy ước dùng xuyên suốt tài liệu (Plan 02 và Phase 2 phải tuân theo)

a. Mọi bảng ở §4 dùng đúng header sau, không có cột thứ 6:

   `| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |`
   `|---|---|---|---|---|`

b. Row ID được viết **in đậm** ở ngay đầu ô Logic, ví dụ `| **G1** ...`. ID theo mẫu chữ cái A-K + số + hậu tố chữ thường (tuỳ chọn), ví dụ `G1`, `F3a`.

c. Ô Verdict chỉ chứa đúng 1 token:
   - **GIỮ** = port nguyên trạng vào `core/` ở Phase 2.
   - **SỬA** = port kèm thay đổi ghi cụ thể trong ô "Cách sửa".
   - **BỎ** = không port. Ô "Cách sửa" khi đó ghi `Lý do bỏ: … / Thay thế: …`.

d. Ô Mức rủi ro chỉ chứa đúng 1 trong: **Cao**, **Trung bình**, **Thấp** — đánh giá rủi ro hành vi HIỆN TẠI gây ra cho dữ liệu Bravo:
   - Cao = sai dữ liệu ERP thật hoặc crash import.
   - Trung bình = sai cảnh báo hoặc sai trường phụ.
   - Thấp = chỉ ảnh hưởng khả năng bảo trì.

e. Ô Bằng chứng luôn có ít nhất 1 trích dẫn code dạng `main_window.py:9529-9538` và ít nhất 1 mã bằng chứng. Mã bằng chứng gồm:
   - `EV-Snn` = số liệu từ 94 file mẫu (§2).
   - `EV-Mnn` = census mapping production (§2).
   - `SQL-nn` = kết quả SQL thật (§3).
   - Có thể thêm hash commit từ §1.

f. Dấu `|` xuất hiện bên trong 1 ô được viết thành `\|`.

g. Mỗi mục bằng chứng là 1 heading cấp 3 dạng `### EV-S01 — …` / `### SQL-01 — …` để ID luôn tìm được bằng grep.

h. Mỗi query ở §3 xuất hiện verbatim trong 1 khối ```sql, theo sau là 1 dòng `**Kết quả:**`.

### 0.3 Ghi chú về nguồn SQL thật (§3)

Công cụ MCP `mssql-boho` không được expose cho executor subagent chạy plan này (giới hạn `tools:` của agent định nghĩa — cùng loại lỗi upstream anthropics/claude-code#13898 đã ghi trong hướng dẫn thực thi). Điều phối viên (coordinator, phiên orchestrator có `mssql-boho` hoạt động) đã tự chạy toàn bộ các query SQL-00 → SQL-08 và chuyển verbatim câu query + kết quả thật cho executor chép nguyên văn vào §3. Không có số liệu nào trong §3 được suy luận hay ước lượng — toàn bộ là kết quả SELECT/WITH/catalog thật, đọc-only, xác nhận qua `check_permissions` (DDL/DROP/DELETE bị khoá ở cấp kết nối).

**Thời điểm chạy:** 2026-09-25. **DB kết nối:** `B10_Boho_Data`. `vB20Item_MKT` nằm ở `BOMTool.dbo`. `usp_B20BOM_Create_ItemCode` là SYNONYM trong `B10_Boho_Data` trỏ tới `B10_Boho.dbo.usp_B20BOM_Create_ItemCode` (xem SQL-03).

## 1. Lịch sử fix 2.2.1 → 2.2.25 (D-04)

_Bảng đầy đủ toàn bộ commit trong khoảng `73081dc^..8a07639` sẽ được Task 2 điền. Dưới đây là hàng khởi tạo (seed) từ Task 1, tìm bằng:_

```bash
git -C D:/project/insertdatasqlserver log --format="%h %ad %s" --date=short -S "_mkt_cache" -- Tools/views/main_window.py
```

| Commit | Ngày | Release | Nội dung | Loại | Row ID liên quan |
|---|---|---|---|---|---|
| `e92c060` | 2026-08-11 | v2.2.4 | fix: lookup trùng tên + MKT fallback cho ItemId NULL — build `_mkt_cache` lần đầu: query `vB20Item_MKT` 1 lần/import, không `ORDER BY`, dict `{ItemTypeSX_Parent: Id}` ghi đè dòng trước bằng dòng sau; áp dụng sau Fill-Forward khi `ItemId=NULL`, tránh gọi `usp_B20BOM_Create_ItemCode` tạo mã rác | FEAT (commit gắn nhãn `fix:` nhưng đây là lần đầu tiên cơ chế MKT fallback tồn tại — không có hành vi cũ nào để "fix", nên phân loại FEAT theo nội dung thật, không theo tiền tố message) | G1 |

`git log -S "_mkt_cache" -- Tools/views/main_window.py` chỉ trả về **đúng 1 commit** (`e92c060`). Không có commit nào sau đó chạm lại vào `_mkt_cache` — tức là dict-collision (nếu có) chưa từng được vá kể từ khi tạo ra, xuyên suốt v2.2.4 → v2.2.25.

### Chuỗi đổi hướng

_Sẽ điền đầy đủ ở Task 2 (§1 Part A, action mục "Chuỗi đổi hướng"). Task 1 chỉ xác nhận mắt xích đầu của 1 chuỗi liên quan trực tiếp đến G1: `_mkt_cache` được tạo tại `e92c060` (v2.2.4) như một phần của thay đổi lớn hơn (lookup trùng tên + MKT fallback) và không bị sửa lại cho tới hết v2.2.25._

## 2. Bằng chứng từ 94 file mẫu & mapping (D-02)

### EV-S00 — Baseline quét 94 file

Nguồn: `v3/docs/audit/bom_sample_scan.json`, đọc bằng script Python một lần (in aggregate, không mở toàn bộ 1.3 MB).

- **Số file:** 94/94 entries.
- **`error` = null:** 94/94 (không file nào bị lỗi parse ở lượt quét này).
- **Cảnh báo (`warnings`) — phân bố:**
  - `Không tìm thấy sheet BOM5`: 14/94 file. Ví dụ: `2573_LAGOME_CT07\2.Kitchen_( TH 07 & TH09 & TH11)_new.xlsm`, `2606_DRC_CT02\12. TABLE_MO41.xlsm`, `2606_DRC_CT02\14. ANIKA CHAIR_AS49.xlsm`.
- **`skipped_sheets` — phân bố (107 lượt bỏ qua trên 94 file):**
  - `BOM_Phan I`: 94/94 file (sheet này bị bỏ qua ở MỌI file — lý do cụ thể EV-S02 sẽ điền ở Task 2).
  - `LOẠI 2`: 12 lượt.
  - `LOẠI 3`: 1 lượt.
- **Loại section (`type`) có mặt, đếm trên tổng 470 section (94 file × 5 section/file):** `HEADER`=94, `BOM2`=94, `BOM3`=94, `BOM4`=94, `BOM5`=94.
- **`mixed_dtype_columns` (khớp fact đã xác minh ở planning-time):** 352/470 section có ít nhất 1 cột mixed-dtype.

Số liệu này là baseline cho toàn bộ EV-S01..EV-S14 mà Task 2 sẽ điền chi tiết.

## 3. Bằng chứng SQL thật qua mssql-boho (D-03)

Xem quy ước nguồn tại §0.3. Mọi query dưới đây là SELECT/WITH hoặc catalog-metadata (`INFORMATION_SCHEMA`, `sys.*`, `OBJECT_DEFINITION`), không có statement ghi dữ liệu, DDL, hay gọi stored procedure. Không ghi lại tên server, login hay connection string — chỉ ghi tên DB, kết quả (đếm/Id/Code/Name).

### SQL-00 — Ngữ cảnh DB & tồn tại đối tượng

```sql
SELECT DB_NAME() AS current_db
```
**Kết quả:** `current_db = B10_Boho_Data`.

```sql
SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE FROM [BOMTool].INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'vB20Item_MKT'
```
**Kết quả:** 1 dòng — `TABLE_CATALOG=BOMTool, TABLE_SCHEMA=dbo, TABLE_NAME=vB20Item_MKT, TABLE_TYPE=VIEW`. Xác nhận `vB20Item_MKT` là 1 VIEW có thật trong `BOMTool.dbo`.

```sql
SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE FROM [BOMTool].INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'vB20Item_MKT' ORDER BY ORDINAL_POSITION
```
**Kết quả:** 60 cột được xác nhận, khớp với cách code dùng. Các cột chính: `ItemTypeSX_Parent` (varchar, nullable), `Id` (int, not null), `Code` (varchar), `Name` (nvarchar), `IsActive` (bit), `CreatedAt` (smalldatetime).

### SQL-01 — Trùng khóa ItemTypeSX_Parent trong vB20Item_MKT (G1, G4)

```sql
SELECT COUNT(*) AS total_rows, COUNT(DISTINCT ItemTypeSX_Parent) AS distinct_types,
       SUM(CASE WHEN ItemTypeSX_Parent IS NULL THEN 1 ELSE 0 END) AS null_type_rows
FROM [BOMTool].[dbo].[vB20Item_MKT]
```
**Kết quả:** `total_rows=15, distinct_types=9 (không tính NULL), null_type_rows=2`.

```sql
SELECT ItemTypeSX_Parent, COUNT(*) AS n, MIN(Id) AS min_id, MAX(Id) AS max_id
FROM [BOMTool].[dbo].[vB20Item_MKT]
GROUP BY ItemTypeSX_Parent
HAVING COUNT(*) > 1
ORDER BY n DESC
```
**Kết quả:** 4 nhóm trùng khóa (tổng 9/15 dòng của view nằm trong 1 nhóm trùng):

| ItemTypeSX_Parent | n | min_id | max_id |
|---|---|---|---|
| C | 3 | 86047 | 424633 |
| F | 2 | 86049 | 424819 |
| I | 2 | 85531 | 424509 |
| NULL | 2 | 421889 | 424817 |

**=> Giả thuyết FIX-01 (dict-collision khi build `_mkt_cache`) được XÁC NHẬN bằng dữ liệu thật.** Với `ItemTypeSX_Parent` có 2-3 dòng, `_mkt_cache[key]` chỉ giữ được 1 `Id` — dòng nào "thắng" phụ thuộc thứ tự trả về của `SELECT ... FROM vB20Item_MKT` (không có `ORDER BY`), tức là **không xác định (nondeterministic)** giữa các lần chạy khác nhau.

Chi tiết từng dòng trong 4 nhóm trùng (khớp đúng với báo cáo bug gốc của người dùng — SIMILI/VẢI/DA và CƠ KHÍ/KÍNH):

```sql
SELECT ItemTypeSX_Parent, Id, Code, Name, IsActive, CreatedAt
FROM [BOMTool].[dbo].[vB20Item_MKT]
WHERE ItemTypeSX_Parent IN ('C','F','I') OR ItemTypeSX_Parent IS NULL
ORDER BY ItemTypeSX_Parent, Id
```
**Kết quả:** 9 dòng —

| ItemTypeSX_Parent | Id | Code | Name | CreatedAt |
|---|---|---|---|---|
| NULL | 421889 | MKT_COKHI | Mã kiến trúc - Cơ khí, kim loại | 2026-08-11 |
| NULL | 424817 | MKT_KINH | Mã kiến trúc - Kính | 2026-09-03 |
| C | 86047 | MKT_SIMILI | Mã kiến trúc - Simili | 2026-07-28 |
| C | 424631 | MKT_VAI | Mã kiến trúc - Vải | 2026-09-03 |
| C | 424633 | MKT_DA | Mã kiến trúc - Da | 2026-09-03 |
| F | 86049 | MKT_VTP | Mã kiến trúc - vật tư phụ | 2026-07-28 |
| F | 424819 | MKT_KEO | Mã kiến trúc - nhóm Vật tư Sofa + Đinh + Keo | 2026-09-03 |
| I | 85531 | MKT_BB | Mã kiến trúc - Bao bì | 2026-07-22 |
| I | 424509 | MKT_BAOBI | Mã kiến trúc - nhóm bao bì | 2026-09-03 |

```sql
SELECT OBJECT_DEFINITION(OBJECT_ID('BOMTool.dbo.vB20Item_MKT')) AS view_def
```
**Kết quả:** `NULL` — không có quyền `VIEW DEFINITION` trên login hiện tại cho đối tượng này, hoặc view được bảo vệ. Không xác nhận được view có tự dedupe/sắp xếp hay không từ định nghĩa — phải dựa vào kết quả truy vấn thực tế ở trên (2 query trước) để kết luận.

**SQL-01(d) — bằng chứng dùng thực tế của nhóm trùng (nối tiếp SQL-01, không phải truy vấn riêng trong đặc tả gốc, nhưng do coordinator chạy thêm vì làm rõ trực tiếp mức rủi ro G1):**

```sql
SELECT bd.ItemId, m.Code, m.Name, COUNT(*) AS n
FROM B20BOMDetail bd
JOIN [BOMTool].[dbo].[vB20Item_MKT] m ON m.Id = bd.ItemId
WHERE m.ItemTypeSX_Parent = 'C' OR m.ItemTypeSX_Parent IS NULL
GROUP BY bd.ItemId, m.Code, m.Name
ORDER BY m.Code
```
**Kết quả:**

| ItemId | Code | Name | n (số dòng BOM chi tiết thật dùng Id này) |
|---|---|---|---|
| 421889 | MKT_COKHI | Mã kiến trúc - Cơ khí, kim loại | 9 |
| 424633 | MKT_DA | Mã kiến trúc - Da | 227 |
| 424817 | MKT_KINH | Mã kiến trúc - Kính | 63 |
| 424631 | MKT_VAI | Mã kiến trúc - Vải | 27 |

Chứng minh rủi ro KHÔNG chỉ lý thuyết: `MKT_SIMILI` (Id cũ nhất trong nhóm C, `86047`) chưa bao giờ "thắng" — 0 dòng BOM nào dùng Id này, tức nó không bao giờ được `_mkt_cache` chọn dù luôn có mặt trong view. `MKT_VAI` (27 dòng) và `MKT_DA` (227 dòng) đều xuất hiện là "người thắng" ở các thời điểm khác nhau → thứ tự trả về của SQL Server **không ổn định giữa các lần chạy ứng dụng khác nhau**, không phải "luôn chọn sai 1 mã cố định". Tương tự nhóm NULL: `MKT_COKHI` thắng 9 lần, `MKT_KINH` thắng 63 lần — đây chính là hiện tượng "SQL sinh BTP rác hoặc nhận nhầm MKT_KINH" mà người dùng mô tả ban đầu.

_(SQL-02 → SQL-08 sẽ được Task 3 điền.)_

## 4. Bảng kiểm kê & verdict (D-05)

### 4.A Nhận diện sheet/section

_(Điền ở Plan 02.)_

### 4.B Nhận diện header

_(Điền ở Plan 02.)_

### 4.C Parse dòng dữ liệu

_(Điền ở Plan 02.)_

### 4.D Resolve theo NguonDL (_resolve_row_mapping)

_(Điền ở Plan 02.)_

### 4.E Resolve dòng chi tiết BOM (_resolve_detail_row)

_(Điền ở Plan 02.)_

### 4.F BTP detection (BOM2)

_(Điền ở Plan 02.)_

### 4.G MKT fallback

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **G1** Build `_mkt_cache` từ `vB20Item_MKT`: `SELECT ItemTypeSX_Parent, Id` không `ORDER BY` → dict, dòng sau ghi đè dòng trước | SỬA | `main_window.py:9529-9538` (build cache), `main_window.py:9835-9839` (áp dụng lookup `_mkt_cache.get(_item_type) or _mkt_cache.get(None)`); `SQL-01` — 4 nhóm trùng `ItemTypeSX_Parent` (C=3, F=2, I=2, NULL=2) trên 15 dòng của view (9 khóa distinct); `SQL-01(d)` cho thấy "người thắng" thực tế đổi giữa các lần chạy (`MKT_VAI`=27 dòng vs `MKT_DA`=227 dòng cùng nhóm C; `MKT_COKHI`=9 vs `MKT_KINH`=63 cùng nhóm NULL) và `MKT_SIMILI` không bao giờ thắng dù luôn có trong view; commit `e92c060` (tạo `_mkt_cache` lần đầu, 2026-08-11, chưa từng được sửa lại tới v2.2.25) | Thêm `ORDER BY` xác định (ví dụ theo `CreatedAt DESC` hoặc `Id DESC`, ưu tiên bản ghi mới nhất) khi build `_mkt_cache`, đồng thời áp dụng tường minh 1 quy tắc ưu tiên khi 1 `ItemTypeSX_Parent` có nhiều dòng — quy tắc ưu tiên cụ thể (dòng nào thắng) là quyết định nghiệp vụ, xem `### Q-01` ở §5, không tự chọn | Cao |

### 4.H SP_HOOK

_(Điền ở Plan 02.)_

### 4.I Fill-Forward lúc insert

_(Điền ở Plan 02.)_

### 4.J Trường header: EmployeeId / ParentDetailRowId_SO / đơn hàng

_(Điền ở Plan 02.)_

### 4.K Xuyên suốt: sanitize, exception, insert

_(Điền ở Plan 02.)_

## 5. Câu hỏi mở cho business owner

### Q-01

**Bối cảnh (G1 — MKT fallback dict-collision):** `vB20Item_MKT` có 4 `ItemTypeSX_Parent` trùng khóa với nhiều `Id`/`Code` khác nghĩa nhau thật sự (không phải trùng dữ liệu rác):

- `C` → `MKT_SIMILI` (Id 86047), `MKT_VAI` (Id 424631), `MKT_DA` (Id 424633)
- `F` → `MKT_VTP` (Id 86049), `MKT_KEO` (Id 424819)
- `I` → `MKT_BB` (Id 85531), `MKT_BAOBI` (Id 424509)
- `NULL` (không map `ItemTypeSX_Parent`) → `MKT_COKHI` (Id 421889), `MKT_KINH` (Id 424817)

Hiện tại code không có quy tắc ưu tiên tường minh — `Id` nào "thắng" phụ thuộc thứ tự SQL Server trả về (không `ORDER BY`), đã chứng minh là không ổn định giữa các lần import (SQL-01(d)). **Câu hỏi:** khi 1 `ItemTypeSX_Parent` (hoặc trường hợp không map `ItemTypeSX_Parent`) trùng nhiều mã kiến trúc như trên, business muốn:

(a) luôn ưu tiên bản ghi mới nhất (`CreatedAt`/`Id` lớn nhất), hay
(b) có 1 bảng ánh xạ tường minh `ItemTypeSX_Parent → Id` cố định (business tự chọn từng cặp), hay
(c) phương án khác?

Trả lời câu này quyết định nội dung cụ thể của ô "Cách sửa" cho G1 ở Phase 2.

_(Các câu hỏi khác — nếu SQL-02..SQL-08 ở Task 3 phát hiện thêm gap bằng chứng — sẽ được thêm ở đây.)_

## 6. Tổng hợp & ánh xạ sang Phase 2

_(Điền ở Plan 02.)_

## 7. Xác nhận của business owner

_(Điền ở Plan 03 — ghi lại xác nhận bằng văn bản, kèm ngày, từ người dùng trong chat. Im lặng, auto-advance hay yolo mode không tính là xác nhận.)_

## 8. Phụ lục

_(Điền ở Plan 02, nếu cần.)_
