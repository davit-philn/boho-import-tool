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

Nguồn: `git log --format="%h|%ad|%s" --date=short 73081dc^..8a07639 -- Tools/services/bom_parser.py Tools/views/main_window.py Tools/services/mapping_loader.py "*CK_Mapping_v5.xlsx"`. `73081dc` = release v2.2.1, `8a07639` = release v2.2.25. Range trả về **41 commit**. Release của mỗi commit xác định bằng `git tag --contains` (cho các release có tag, v2.2.1→v2.2.14) hoặc trực tiếp từ tiền tố `vX.X.X` trong message (cho v2.2.15→v2.2.25, không có git tag) — riêng v2.2.15→18 không có tiền tố version trong message, xác định qua đối chiếu `release_date` trong `version.json` của từng bản portable với ngày commit thật (khớp 1-1, xem cột Ngày).

**8 commit chỉ đụng UI/theme/updater đã loại khỏi bảng dưới đây** (không ảnh hưởng logic xử lý dữ liệu BOM): `013f2be` (bắt buộc cập nhật, popup modal), `89a8110` (ui-review theme tokens/diacritics), `c9ebfcc` (bổ sung cập nhật tự động), `c383e47` (độ rộng ô Dự án), `1cbd29f` (ẩn nút toolbar), `461dba0` (thông báo/popup thân thiện, gỡ thuật ngữ kỹ thuật), `398c7bd` (hiện AppVersion trong Log DB), `9854b27` (không nuốt exception khi tải update lỗi — lỗi của bộ tự-cập-nhật, không phải import BOM).

| Commit | Ngày | Release | Nội dung | Loại | Row ID liên quan |
|---|---|---|---|---|---|
| `73081dc` | 2026-08-08 | v2.2.1 | feat: bump v2.2.1 + UPPER/LOWER/TRIM transform cho BOM và THDM | FEAT | K (Biến_Đổi — xem `1954780` bên dưới, hành vi này bị hỏng và phải vá lại) |
| `4cddcdb` | 2026-08-08 | v2.2.2 | Update CK_Mapping_v5.xlsx (thay đổi binary, message không mô tả nội dung cụ thể) | chore | — |
| `e92c060` | 2026-08-11 | v2.2.4 | fix: lookup trùng tên + MKT fallback cho ItemId NULL — build `_mkt_cache` lần đầu | FEAT (xem giải thích phân loại ở §0/hàng gốc Task 1) | G1, E (lookup trùng tên) |
| `ab66103` | 2026-08-11 | v2.2.5 | perf: revert duplicate-name popup in `_lookup_generic` | REVERT | E (Chuỗi #2, bước 2) |
| `449972a` | 2026-08-12 | v2.2.6 | fix: skip fuzzy_name lookup for placeholder item names (`_`, `--`, etc.) | FIX | E10, F4 (placeholder) |
| `1694b70` | 2026-08-24 | v2.2.6 | fix: tie-break tự động khi 1 tên trùng nhiều mã B20Item (ưu tiên NVL), index hóa Tier 1/2 | FIX | E8, E12 (liên quan trực tiếp `SQL-08`) |
| `8354cef` | 2026-08-24 | v2.2.6 | feat: ưu tiên lookup ItemId qua Mã Vật Tư (Code) trước khi fallback Name (BOM3) | FEAT | D, E |
| `65e322d` | 2026-08-24 | v2.2.6 | feat: phân biệt BTP vs NVL khi ItemId trống (chỉ BOM2) — BTP dùng USP thay MKT | FEAT | F, G (Chuỗi #1, bước 2) |
| `639d938` | 2026-08-25 | v2.2.6 | feat: tái dùng mã BTP đã tạo trước đó thay vì tạo rác mỗi lần import lại | FEAT | E9 (liên quan trực tiếp `SQL-02`) |
| `ebc9afe` | 2026-08-25 | v2.2.6 | feat: thêm rule BTP cho dòng STT nguyên có dấu "+" trong Tên vật tư | FEAT | F (Chuỗi #3, bước 1) |
| `7166b36` | 2026-08-25 | v2.2.6 | fix: `_build_bom_detail_caches` không build cache riêng cho `code_then_name` | FIX | D, E |
| `8926a2b` | 2026-08-25 | v2.2.6 | fix: BOM4.ItemId tắt fuzzy Tier3, chỉ exact/normalize — tránh mở nhầm dự án khác | FIX | D, E |
| `63f6e6a` | 2026-08-25 | v2.2.6 | fix: dòng BTP dạng "+"-ghép bị SP âm thầm bỏ qua, không tạo mã | FIX | F, H (Chuỗi #3, bước 2) |
| `a9a891a` | 2026-08-25 | v2.2.6 | feat: BOM Import — suy ra Dự án từ Đơn hàng đã chọn thay vì fuzzy_name Excel | FEAT | J (trường header) |
| `432fa70` | 2026-08-26 | v2.2.8 | feat: bổ sung ParentDetailRowId_SO vào HEADER mapping (B20BOM) | FEAT | J4 (Chuỗi #7, bước 1) |
| `f977b79` | 2026-08-27 | v2.2.9 | fix: BizDocId_SO/DetailRowId_SO ra NULL dù đã chọn đúng Đơn hàng trên dropdown | FIX | J |
| `3e679fc` | 2026-08-27 | v2.2.9 | fix: compound-key parsing thiếu Pass2 suffix-match, rơi hết vào MKT | FIX | D, E, G (tăng lưu lượng rơi vào MKT fallback) |
| `113c443` | 2026-08-27 | v2.2.10 | fix: coi "_"/"-" như rỗng khi nhận diện BTP + chặn placeholder ở `nvl_or_btp` | FIX | F, E10 |
| `e414085` | 2026-08-27 | v2.2.10 | fix: XML blank-ItemName cho SP_HOOK thiếu case placeholder ("_"/"-"), chỉ có "+" | FIX | H |
| `fb137cc` | 2026-08-27 | v2.2.10 | fix: BTP Rule 1 áp dụng mọi tầng STT (không chỉ STT nguyên) + bổ sung IsTool mapping | FIX | F (Chuỗi #4) |
| `e8a5007` | 2026-08-28 | v2.2.13 | feat: thêm Nguồn_DL=ExcelCell cho field Quantity, đọc ưu tiên ở Y3 | FEAT | A, J (EV-S14) |
| `1954780` | 2026-09-03 | v2.2.14 | fix: BOM không áp dụng Biến_Đổi (UPPER/LOWER/TRIM) dù đã khai báo mapping | FIX | D, K (vá lại tính năng do `73081dc` khai báo nhưng không chạy) |
| `817bbbe` | 2026-09-19 | v2.2.15 | fix: từ khóa footer "duyệt" quá chung chung, làm mất dữ liệu hàng loạt | FIX | C2 (nghiêm trọng — xoá nhầm toàn bộ dòng phía dưới) |
| `458fdfd` | 2026-09-19 | v2.2.16 | fix: BOM2 dòng con rỗng Tên vật tư nhầm thành BTP khi dòng cha đã có vật tư thật | FIX | F (Chuỗi #1, bước 3) |
| `006fc61` | 2026-09-19 | v2.2.17 | fix: nguy cơ treo popup từ background thread khi Import gặp mã trùng tên | FIX | E (Chuỗi #2, bước 3) |
| `cb2c1d7` | 2026-09-19 | v2.2.18 | fix: STT trùng số giữa các section khác nhau làm lây nhiễm sai nhận diện BTP | FIX | F (Chuỗi #5) |
| `3fd7985` | 2026-09-21 | v2.2.19 | feat: BOM2 mảnh cắt vật liệu (vải/da) ưu tiên tìm mã tái dùng, fallback MKT thay vì tạo mã mới; SSL certifi cho auto-update; mở rộng logic tìm vật liệu cha cho STT dạng phẳng (AS27-style) | FEAT | F, G (Chuỗi #1, bước 4) |
| `6beb486` | 2026-09-21 | v2.2.20 | chore: sửa label mapping DETAIL cho FinishCode/FinishSide trong CK_Mapping_v5.xlsx | chore | E4 (áp dụng cho BOM2 — xem ghi chú dưới bảng) |
| `feda97e` | 2026-09-21 | v2.2.21 | chore: sửa label mapping DETAIL cho trường Diện tích Sơn hoàn thiện trong CK_Mapping_v5.xlsx | chore | E4 (áp dụng cho BOM2, **KHÔNG** áp dụng cho BOM5 — xem ghi chú dưới bảng và `EV-S04`) |
| `9dc3c92` | 2026-09-21 | v2.2.22 | fix: EmployeeId của B20BOM (tab BOM) bị hard-code = 1, không lấy theo nhân viên chọn trên UI | FIX | J1, J2 (Chuỗi #6) |
| `78b36fc` | 2026-09-22 | v2.2.23 | fix: ParentDetailRowId_SO bị gán nhầm = mã đơn hàng thay vì công thức Mục số + Mã đơn hàng | FIX | J4 (Chuỗi #7, bước 2) |
| `45031a8` | 2026-09-25 | v2.2.24 | fix: ghi đầy đủ traceback vào `import_crash.log` khi Import thất bại | FIX | K (Chuỗi #8, bước 1 — logging để chẩn đoán) |
| `8a07639` | 2026-09-25 | v2.2.25 | fix: v2.2.25 — Import thất bại với lỗi 'int object has no attribute strip' | FIX | K1, G1-liền kề (Chuỗi #8, bước 2 — xem ghi chú dưới bảng) |

**Ghi chú — `6beb486`/`feda97e` chỉ vá 1 phần (liên quan `EV-S04`):** đối chiếu trực tiếp mapping hiện tại (`load_mapping()`), `BOM2.FinishArea` có `Ten_Excel = "Diện tích (m2) - Sơn H. Thiện"` (đã rút gọn, đúng như `feda97e` mô tả), nhưng `BOM5.FinishArea` vẫn còn `Ten_Excel = "Diện tích (m2) - Sơn hoàn thiện"` (dạng đầy đủ CŨ, chưa được vá). Đây chính là 1 trong các mục `EV-S04` "mapped Ten_Excel entries với 0 match" ở §2 — hàng BOM5 không khớp bất kỳ header Excel thật nào trong 94 file mẫu. Khớp với `docs/audit/AUDIT-2026-09-21.md` finding #3 (2 test snapshot thất bại trên `bom_valid.xlsm`, đúng thời điểm các commit sửa label DETAIL v2.2.20/v2.2.21 — audit nghi ngờ chính các commit này là nguyên nhân, và bằng chứng ở đây xác nhận nghi ngờ đó cho riêng trường `FinishArea`/BOM5).

**Ghi chú — `8a07639` không sửa `_mkt_cache` (liên quan `G1`):** commit vá dòng `main_window.py:9847` — `str(row_vals.get('Unit') or '').strip()` — nằm NGAY TRONG khối MKT fallback (Unit-fallback cho dòng nhận `ItemId` từ `_mkt_cache`, main_window.py:9847-9856, đọc ở Task 1), nhưng đây là 1 lỗi khác: crash khi `Unit` đọc từ Excel bị nhầm kiểu số (`int`/`float`) thay vì chuỗi, không liên quan tới việc `_mkt_cache` build thiếu `ORDER BY` (G1). `45031a8` (v2.2.24) thêm traceback đầy đủ vào `import_crash.log` để chẩn đoán chính xác lỗi này trước khi `8a07639` vá nó — 1 chuỗi 2 bước độc lập với G1 nhưng nằm trong cùng vùng code.

### Chuỗi đổi hướng

1. **MKT fallback → tách BTP khỏi MKT → mảnh cắt quay lại MKT.** `e92c060` (v2.2.4) tạo `_mkt_cache`, mọi dòng `ItemId=NULL` đều rơi vào MKT fallback. `65e322d` (v2.2.6) đổi hướng: phân biệt BTP khỏi NVL, để BTP dùng USP tạo mã thật thay vì nhận mã MKT tạm. `458fdfd` (v2.2.16) và `3fd7985` (v2.2.19) đổi hướng lần 2: dòng con (mảnh cắt, X.1/X.2...) mà dòng cha đã có vật tư thật KHÔNG còn bị coi là BTP nữa — quay lại nhận MKT fallback (hoặc tìm mã tái dùng trước). Ảnh hưởng: G1 (MKT fallback), F (BTP detection Rule 1b/cut-piece).
2. **Popup trùng tên → revert vì chậm → chạy nền không popup.** `e92c060` (v2.2.4) thêm popup cho user chọn khi 1 tên khớp nhiều bản ghi. `ab66103` (v2.2.5) revert popup này (lý do hiệu năng, message "perf:"). `006fc61` (v2.2.17) xử lý lại theo hướng thứ 3: chạy trên background thread, không mở popup treo UI khi gặp trùng tên lúc Import thật (khác với luồng lookup tương tác). Ảnh hưởng: E (lookup theo Name).
3. **BTP Rule 2 "+" → SP bỏ sót → vá SP-side.** `ebc9afe` (v2.2.6) thêm rule: STT nguyên có dấu "+" trong Tên vật tư → BTP. `63f6e6a` (v2.2.6, cùng ngày) phát hiện các dòng này bị `usp_B20BOM_Create_ItemCode`/SP_HOOK âm thầm bỏ qua (không tạo mã) — vá bổ sung XML case cho placeholder "+". Ảnh hưởng: F (Rule 2), H (SP_HOOK).
4. **BTP Rule 1 mở rộng phạm vi.** `fb137cc` (v2.2.10) mở rộng Rule 1 (Tên vật tư rỗng + có Tên chi tiết/con) từ chỉ-STT-nguyên sang MỌI tầng STT (1, 1.1, 1.1.1...). Không phải revert — mở rộng phạm vi áp dụng của rule đã có. Ảnh hưởng: F (Rule 1).
5. **STT trùng số giữa section → lây nhiễm sai BTP → scope theo section.** `cb2c1d7` (v2.2.18) vá lỗi STT "1" ở section A bị lẫn với STT "1" ở section B khi nhận diện BTP — thêm `_sec_id_seq` để scope theo section-header gần nhất. Ảnh hưởng: F (toàn bộ BTP detection, đã đọc ở Task 1 dòng 9662-9674).
6. **EmployeeId hard-code = 1.** `9dc3c92` (v2.2.22) sửa EmployeeId của B20BOM bị hard-code thành 1 thay vì lấy theo nhân viên chọn trên UI. Đây là sửa hành vi sai từ trước, không phải đổi hướng thiết kế. Ảnh hưởng: J1/J2. **Lưu ý còn hở (xem `SQL-04` ở §3):** dữ liệu thật SAU commit này vẫn cho `EmployeeId=1` là giá trị phổ biến nhất — chưa xác nhận được đây là hành vi đúng (nhân viên Id=1 có thật) hay vẫn còn 1 nhánh fallback cũ.
7. **ParentDetailRowId_SO: thêm field → gán sai → vá theo công thức.** `432fa70` (v2.2.8) thêm `ParentDetailRowId_SO` vào HEADER mapping (B20BOM). `78b36fc` (v2.2.23) phát hiện field này bị gán nhầm = mã đơn hàng (`ParentBizDocId`) thay vì công thức đúng "Mục số|@ParentBizDocId" — vá lại. Ảnh hưởng: J4. **Lưu ý còn hở (xem `SQL-05` ở §3):** dữ liệu thật SAU `78b36fc` vẫn có dòng sai theo mẫu cũ, tỷ lệ sai không giảm về 0 — nghi ngờ còn 1 nhánh code khác (có thể luồng THDM hoặc 1 loại sản phẩm cụ thể) chưa được `78b36fc` bao phủ.
8. **Crash `.strip()` trên `int`: thêm log để chẩn đoán → xác định root cause → vá.** `45031a8` (v2.2.24) thêm traceback đầy đủ vào `import_crash.log` vì thông báo lỗi cũ quá ngắn để chẩn đoán. `8a07639` (v2.2.25) dùng traceback đó xác định chính xác dòng `main_window.py:9847` (Unit-fallback trong khối MKT, không phải chính `_mkt_cache`) và vá bằng `str()` ép kiểu trước `.strip()`. Đây là bằng chứng sống rằng vùng code MKT fallback (khối G1) đã từng gây crash toàn bộ Import trên dữ liệu thật — củng cố mức rủi ro Cao đã gán cho G1 ở Task 1, dù sub-issue này (kiểu dữ liệu Unit) khác với sub-issue dict-collision.
9. **(Phát hiện thêm, không nằm trong danh sách gốc của Task 2) Label mapping DETAIL sửa nửa vời.** `6beb486`/`feda97e` (v2.2.20/21) đổi `Ten_Excel` của `FinishCode`/`FinishSide`/`FinishArea` sang dạng rút gọn — nhưng chỉ áp dụng cho **BOM2**, không áp dụng cho **BOM5** (field `FinishArea` trùng `sql_col` nhưng khai báo Ten_Excel riêng theo section). Kết quả: BOM5.FinishArea hiện không khớp header Excel thật nào trong 94 file mẫu (`EV-S04`). Trùng thời điểm với `docs/audit/AUDIT-2026-09-21.md` finding #3 (2 test snapshot thất bại trên fixture `bom_valid.xlsm`, audit nghi ngờ chính 2 commit này). Ảnh hưởng: E4 (header→sql_col matching, mục §4.B).

Ghi chú chung: `version.json` trong `Tools/installer_output/BOHO_IMPORT_BOM_THDM_v2.2.N_Portable/` cho các bản v2.2.20 → v2.2.25 đều lặp lại NGUYÊN VĂN `release_notes` của v2.2.19 — xác nhận trực tiếp (`cat version.json` từng bản). Coi các note này là STALE cho các bản 2.2.20-2.2.25; tiêu đề commit git là nguồn đáng tin cậy duy nhất cho các bản này.

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

Số liệu này là baseline cho toàn bộ EV-S01..EV-S14 dưới đây.

**Nguồn chung cho EV-S01 → EV-S14, EV-M01:** script một-lần `tests/tmp/ba_sample_evidence.py` — import trực tiếp `parse_bom_file`, `_is_encrypted_excel`, `_load_sheet_config`, `_detect_sheet_type`, `_find_header_row`, `_build_headers`, `_parse_sheet` từ `services.bom_parser` và `load_mapping`, `SECTION_STT_PATTERN`, `NUMERIC_STT_PATTERN` từ `services.mapping_loader`, chạy trên toàn bộ 94 file dưới `v3/sample_data/bom_import/`. Với các predicate chỉ tồn tại inline trong `main_window.py` (gắn với `self`, ví dụ BTP detection, footer/section-header/identity-drop trong `_parse_sheet`), script **tái hiện lại y hệt logic** (trích dẫn số dòng gốc ở từng mục) thay vì import trực tiếp. Kết quả ghi vào `tests/tmp/ba_sample_evidence.json`, đọc 1 lần để viết các mục dưới đây (không encrypt file nào trong 94 file mẫu — `encrypted_files: []`, không lỗi — `errors: []`).

### EV-S01 — Cảnh báo thiếu section (A5)

`Không tìm thấy sheet BOM5`: **14/94 file**. Ví dụ: `2573_LAGOME_CT07/2.Kitchen_( TH 07 & TH09 & TH11)_new.xlsm`, `2606_DRC_CT02/12. TABLE_MO41.xlsm`. Không file nào thiếu BOM2/BOM3/BOM4 (0/94 mỗi loại — các sản phẩm mẫu luôn có đủ 3 section này, chỉ BOM5 — "Chi tiết hoàn thiện" — là tuỳ chọn).

### EV-S02 — skipped_sheets: hidden vs không khớp _CONFIG (A2, A3)

Đối chiếu trực tiếp `wb[name].sheet_state` và `_detect_sheet_type(name, sheet_config)` cho từng sheet, từng file (không dựa vào field `skipped_sheets` của `parse_bom_file()` một mình, vì field đó chỉ liệt kê sheet VISIBLE mà `_parse_sheet` trả `None` — sheet hidden bị `continue` ở `parse_bom_file` dòng 767-769 TRƯỚC KHI tới bước phân loại, nên không bao giờ xuất hiện trong `skipped_sheets`).

- `BOM_Phan I`: **94/94 file** — LUÔN VISIBLE, LUÔN `_detect_sheet_type()` = UNKNOWN (không có mục `_CONFIG` nào khai `contains` khớp "I" — `_CONFIG` chỉ đăng ký BOM2="II", BOM3="III", BOM4="IV", BOM5="V"). Sheet "Phần I" (thường là trang bìa/hướng dẫn) không bao giờ được map vào section nào, ở mọi file.
- `LOẠI 2`: visible+UNKNOWN ở **12/94 file** (12 lượt, các file còn lại sheet này ở trạng thái hidden nên không xuất hiện trong danh sách bỏ qua).
- `LOẠI 3`: visible+UNKNOWN ở **1/94 file** (còn lại hidden).
- `TC_CUA`, `TC_KH`, `LOẠI 1`, `LOẠI 4`, `LOẠI 5 AIR`/`LOẠI 5-AIR`, `LOẠI 5 PALEET`/`LOẠI 5-PALEET`, `LOẠI 6`: xuất hiện trong workbook nhưng LUÔN hidden ở cả 94 file (0/94 visible+UNKNOWN) — không bao giờ vào `skipped_sheets` của `parse_bom_file()`, hoàn toàn vô hình với người đọc report cũ.

Kết luận: `skipped_sheets` của `bom_sample_scan.json` chỉ phản ánh 1 phần sự thật (sheet visible nhưng không khớp _CONFIG) — sheet phụ trợ hidden (tham chiếu tra cứu LOẠI/TC) bị bỏ qua hoàn toàn không để lại dấu vết trong report gốc.

### EV-S03 — Header artifacts từ `_build_headers` (B2, B3, B4)

Đo trên header GỐC (trước khi `_parse_sheet` xoá cột `COL_*`/toàn-rỗng ở dòng 585-586), bằng cách gọi trực tiếp `_find_header_row` + `_build_headers` trên từng sheet BOM2-5 của cả 94 file (376 sheet BOM2-5 xét được, = 94×4, không tính các file thiếu BOM5 vì mất sheet đó).

- **Tên bị thêm hậu tố dedupe (`_N`, ví dụ `Số_lượng_1`, `SLSP_LỌC_IN_1`):** 157 lượt cột trên **94/94 file** (mọi file có ít nhất 1 cột bị dedupe ở 1 section nào đó).
- **Tên `COL_N` (không xác định được nhãn từ 2 dòng header):** 362 lượt cột (trung bình gần 1 cột `COL_N`/section) trên **94/94 file**.
- **Ký tự lỗi công thức Excel trong tên cột (`#NAME?`):** 6 lượt trên **2/94 file** — `2632_ISHCMC_CT03_REV01/13_MBT401.xlsm` (BOM2, BOM3, BOM5) và `2632_ISHCMC_CT03_REV01/2.BOOK CASE_LC-402A.xlsm` (BOM2, BOM3, BOM5). Tên cột thực tế ghi nhận: `SLSP_#NAME?`, `SLSP_#NAME?_1`. Không gặp `#REF!`/`#VALUE!` trong 94 file mẫu (0/94).
- **Cột nằm bên phải cột cuối cùng có map với bất kỳ `Ten_Excel` nào (orphan columns):** 362 lượt (= đúng bằng số `COL_N`, vì phần lớn orphan chính là các cột không xác định) trên **94/94 file**. Ví dụ điển hình lặp lại ở hầu hết file, đúng như dự đoán ở planning-time: BOM2 → `SLSP`, `SLSP_LỌC_IN`, `SLSP_PLDG` (+ đôi khi 1 cột tính toán ẩn dạng `PB_Carb_P2_-_25x1220x2440_(mm)_...`); BOM3 → `Hình_ảnh`, `SLSP`, `SLSP_LỌC_IN`; BOM4 → `SLSP`, `SLSP_LỌC_IN`; BOM5 → `Khối_lượng_tinh`, `SLSP`, `SLSP_LỌC_IN`.

### EV-S04 — Header → sql_col matching (B5, E3, E4, F2)

Dùng lại đúng 2-pass mà `_resolve_detail_row` dùng (`main_window.py:9202-9254`): pass 1 exact/norm-exact, pass 2 suffix-match (`endswith`).

- **Cột Excel chỉ khớp được qua suffix-match (không khớp exact):** 71 lượt trên **71/94 file**. 100% các trường hợp là cùng 1 cặp: `SLg_Tên_Vật_Tư` (header có prefix "SLg_" do merge với ô header-cha) khớp `Tên vật tư` → `ItemName` — đúng ví dụ nêu ở planning-time.
- **Mapping row (`Ten_Excel`) không khớp header thật nào, trong TOÀN BỘ 94 file (0/94 khớp):**
  - BOM2: `Chiều dài chỉ cạnh` — 94/94 file không khớp.
  - BOM3: `Slg` — 188 lượt không khớp trên 94 file (2 mapping row cùng dùng `Ten_Excel="Slg"`: `Quantity` và `Quantity9` — cả 2 đều không khớp ở mọi file → 94×2).
  - BOM4: `Slg` — 94/94 file không khớp (1 mapping row: `Quantity`).
  - BOM5: `Khối lượng` (80/94), `Ghi chú` (80/94), `Chiều dài chỉ cạnh` (80/94), `Diện tích (m2) - Sơn hoàn thiện` (80/94 — **xem ghi chú §1 dưới bảng `feda97e`/`6beb486`: label này bị bỏ sót khi vá BOM2, còn BOM5 vẫn giữ label cũ không khớp header Excel thật**), `Diện tích (m2)  - Sơn lót` (57/94). (80/94, không phải 94/94, vì 14/94 file không có sheet BOM5 — xem `EV-S01` — nên các field BOM5 chỉ xét được trên 80 file có sheet).

Đây là bằng chứng trực tiếp cho D-02: những field này hoặc chưa từng nhận dữ liệu thật từ 94 file mẫu (verdict cần cân nhắc BỎ/SỬA ở Phase 2, không phải GIỮ mặc định), hoặc — trường hợp `Diện tích...Sơn hoàn thiện` ở BOM5 — là 1 lỗi mapping đã biết nhưng vá thiếu.

### EV-S05 — STT mixed-dtype & STT nguyên 0 (C10)

- **Section có STT mixed-dtype (int/float/str trộn lẫn trong 1 cột STT):** 352/470 (khớp đúng `EV-S00`).
- **Dòng có STT nguyên = 0:** **0/94 file** trong dữ liệu ĐÃ PARSE (dataframe sau `parse_bom_file()`). Lý do: `_parse_sheet` (`bom_parser.py:545-547`) chủ động `continue` (loại bỏ) mọi dòng có `stt == "0"` NGAY TRONG bước parse, trước khi trả `df` — nên STT=0 không bao giờ sống sót tới dataframe cuối cùng để mixed-dtype ảnh hưởng tới nó. Đây là hành vi chủ đích, không phải thiếu bằng chứng.

### EV-S06 — Footer keyword & dữ liệu bị cắt bỏ theo footer (C2)

Tái hiện logic `_parse_sheet` dòng 548-555 (khớp `FOOTER_KEYWORDS`, sau đó bỏ MỌI dòng phía dưới, kể cả không rỗng).

- Từ khóa khớp: `Tổng cộng` — 174 lượt trên **94/94 file** (ví dụ `2511_SYCAMORE_CT01/1. CUA TOILET KHUYET TAT KHU VUC COMMON_DH114a.xlsm`, section BOM2); `Kiểm duyệt` — 188 lượt trên **94/94 file** (ví dụ file trên, section BOM3). (Từ khóa `"duyệt"` đứng riêng đã bị loại khỏi `FOOTER_KEYWORDS` từ `817bbbe`/v2.2.15 — xem `bom_parser.py:50-53` — nên không xuất hiện ở đây; đúng như code hiện tại.)
- **Số dòng có dữ liệu thật nằm dưới điểm khớp footer, bị loại bỏ vĩnh viễn:** 2.573 dòng trên **94/94 file** (trung bình ~27 dòng/file; ví dụ `2511_SYCAMORE_CT01/1. CUA TOILET KHUYET TAT KHU VUC COMMON_DH114a.xlsm`, section BOM2). Đây là hệ quả trực tiếp, đo được, của cơ chế "khớp footer 1 lần → bỏ hết phần dưới" (dòng 554-555: `if _footer_seen: continue`) — không phải giả thuyết.

### EV-S07 — STT chữ cái / số La Mã / GIAO RỜI (C3, C4, C5, I2)

- **Dòng section-header STT dạng chữ cái (A, B, C...), bị giữ lại cho Fill-Forward nhưng KHÔNG insert:** 1.083 dòng trên **94/94 file**.
- **Dòng STT số La Mã (I, II, III...), rơi qua nhánh data-row (không bị coi là section-header):** 49 dòng trên **31/94 file**. Toàn bộ 49 dòng đều là `I` hoặc `II` — không gặp La Mã lớn hơn II trong 94 file mẫu.
- **Dòng GIAO RỜI (STT dạng chữ cái nhưng có kích thước vật lý → xử lý như data row, không phải section-header):** 16 dòng trên **8/94 file**. Toàn bộ đều là STT="I" (không "II" trừ 1 file `2606_DRC_CT02/43.DEM_CA49.xlsm` có cả "I" và "II" ở cả BOM2 và BOM4) — ví dụ: `2573_LAGOME_CT07/10.2.STUDY - DESK_SD02_TH05.xlsm` (BOM2, STT="I"), `2606_DRC_CT02/13.SWIVEL ARMCHAIR_AS23.xlsm` (BOM2 và BOM4, STT="I").

### EV-S08 — Dòng bị loại theo rule identity / rule chữ-hoa-toàn-bộ (C7, C8, C9)

- **Rule identity (`_is_data_row` — không có giá trị thật ở bất kỳ cột "Tên/Mã/STT/Name/Code/Item" nào):** 169.486 dòng trên **94/94 file** (đã đếm bằng script riêng không giới hạn để có tổng chính xác). **Quan trọng: 0/169.486 dòng này có STT khác rỗng hoặc bất kỳ nội dung Tên/Mã nào** — toàn bộ là dòng Excel "used range" vượt xa vùng dữ liệu thật do định dạng ô (viền/màu nền) kéo dài hàng nghìn dòng trống phía dưới bảng thật, KHÔNG PHẢI mất dữ liệu thật. Số dòng nhiều nhất ở 1 file: `2606_DRC_CT02/22.1. SOFA AS48.B.1.xlsm` và `22.2. SOFA AS48.B.2.xlsm` (2.128 dòng mỗi file, `35. BARSTOOL AS21.xlsm` 2.127 dòng). Rule này hoạt động đúng như thiết kế (không sai), nhưng cho thấy chi phí lặp thừa mỗi lần import (duyệt qua hàng trăm nghìn dòng định dạng-rỗng).
- **Rule "tên chữ-hoa-toàn-bộ là tiêu đề nhóm" (ví dụ "VẬT TƯ CHÍNH"):** **0/94 file** — rule này CHỈ tồn tại trong `_parse_section_excel_rows` (`bom_parser.py:396-403`), pipeline dùng RIÊNG cho THDM. BOM dùng `_parse_sheet` (`bom_parser.py:461-588`), hàm này KHÔNG có rule chữ-hoa-toàn-bộ nào cả — về mặt cấu trúc code, BOM không thể bị loại dòng vì lý do này. Đây là phát hiện quan trọng: giả định "BOM và THDM dùng chung 1 rule-set parse" (gợi ý trong đặc tả D-01/canonical_refs) KHÔNG đúng ở điểm này — 2 pipeline khác nhau thật sự.

### EV-S09 — BTP predicate hits, BOM2 (F3-F12)

Tái hiện đúng `main_window.py:9717-9765` (đọc ở Task 1), áp trên `df` đã parse của từng file (cùng cấu trúc cột mà app thật dùng).

- **Rule 2 (STT nguyên + "+" trong Tên vật tư):** 12 dòng trên **6/94 file**.
- **Rule 1a (Tên vật tư rỗng, có Tên chi tiết thật):** 744 dòng trên **91/94 file**.
- **Rule 1b (Tên vật tư rỗng, Tên chi tiết cũng rỗng, nhưng có dòng con thập phân):** **0/94 file**. Trong toàn bộ 94 file mẫu, không có trường hợp nào rơi vào nhánh này riêng biệt — mọi dòng cha rỗng-Tên-vật-tư có con đều đồng thời có Tên chi tiết thật (Rule 1a bao trùm hết). Không phải lỗi đo — dữ liệu thật của 94 file không tạo ra case này.
- **Trường hợp loại trừ (rỗng cả Tên vật tư lẫn Tên chi tiết, không con):** 4 dòng trên **1/94 file**.
- **Dòng con (mảnh cắt) mà dòng cha đã có vật tư thật (không bị coi BTP, không cho SP tự tạo mã):** 80 dòng trên **18/94 file**.
- **STT trùng lặp giữa các section trong cùng 1 sheet (case mà `cb2c1d7`/v2.2.18 vá):** 53 giá trị STT trùng trên **13/94 file** — xác nhận bằng dữ liệu thật rằng bug này (trước khi vá) là hiện tượng phổ biến (14% số file mẫu), không phải edge case hiếm.

### EV-S10 — Giá trị placeholder trong Tên vật tư / Tên chi tiết (F4, E10, E12)

`_PLACEHOLDER_VALS = {'_', '--', '-', 'x', 'n/a'}` (`main_window.py:9659`). Chỉ giá trị `_` xuất hiện thật trong 94 file mẫu: 63 lượt trên **12/94 file**, toàn bộ ở cột Tên vật tư (không gặp ở Tên chi tiết). Các placeholder khác (`--`, `-`, `x`, `n/a`) — **0/94 file**, không xuất hiện trong tập mẫu này (không có nghĩa code không cần xử lý — chỉ có nghĩa 94 file hiện tại không chứng minh được các case đó).

### EV-S11 — Giá trị non-str trong cột kiểu text (K1)

Đếm giá trị int/float chảy vào field khai `kieu_dl` thuộc nhóm text (`varchar`/`nvarchar`/...) khi lookup KHÔNG bật (điều kiện `not kl` ở `main_window.py:9274-9275` — đúng dòng quyết định có ép kiểu hay không). **Tối thiểu 5.000 lượt** (giới hạn thu thập của script ở mức này để tránh JSON quá lớn — số thật cao hơn) trên **42/94 file** (ví dụ `2511_SYCAMORE_CT01/1. CUA TOILET KHUYET TAT KHU VUC COMMON_DH114a.xlsm`, section BOM2, cột `ItemType`, STT="1", kiểu `int`), chỉ ở 2 cột: `ItemType` và `Code` — cả hai đều có `Ten_Excel = "STT"` (đọc trực tiếp STT dạng số từ Excel). Kiểu Python gặp: `int` (4.188/5.000 mẫu) và `float` (812/5.000 mẫu). **Đối chiếu với `EV-S13` (Fill-Forward)**: giá trị `ItemType` non-str này bị Fill-Forward ghi đè vô điều kiện trước khi insert (xem `EV-S13`) nên phần lớn KHÔNG chảy tới DB dưới dạng số — nhưng `Code` (sql_col riêng, KHÔNG nằm trong danh sách Fill_Forward=1) thì không có cơ chế ghi đè nào tương tự; đây là câu hỏi mở, xem `Q-04` ở §5 (không tự kết luận `Code` có bị `.strip()`-crash hay không nếu chưa đọc được cách trường này được dùng tiếp — có thể chỉ dùng để lookup, không insert trực tiếp).

### EV-S12 — global_meta xung đột giữa các sheet (A4)

`parse_bom_file()` áp dụng "sheet đầu tiên thắng" (`bom_parser.py:786-790`: `if _k not in global_meta: global_meta[_k] = _v`). Đo trực tiếp bằng cách gọi `_parse_sheet()` (hàm thật) riêng cho từng sheet BOM2-5 của mỗi file, so sánh giá trị `metadata` giữa các sheet cùng 1 khóa.

**229 xung đột trên 93/94 file** (hầu như mọi file đều có ít nhất 1 khóa meta khác nhau giữa các sheet). Theo khóa: `Kích thước` (80 lần), `Số lượng` (79 lần), `Hoàn thiện` (69 lần), `Vật liệu` (1 lần). Ví dụ thật (`2511_SYCAMORE_CT01/1. CUA TOILET KHUYET TAT KHU VUC COMMON_DH114a.xlsm`): khóa `Kích thước` = "Theo bản vẽ" ở BOM2/BOM3/BOM4 nhưng = "DÀI" ở BOM5; khóa `Số lượng` = "1 Hệ" ở BOM2/BOM3/BOM4 nhưng = "9" ở BOM5 — "sheet đầu tiên thắng" (BOM2) khiến giá trị "9" (số lượng thật riêng của BOM5) bị bỏ qua hoàn toàn dù khác nghĩa rõ rệt với "1 Hệ".

### EV-S13 — Fill-Forward: xung đột giữa parse-time và insert-time (C6, I4)

Trường Fill_Forward=1 DUY NHẤT của cả 4 section BOM2-5 là `ItemType`, với `Ten_Excel = "STT"` (đọc trực tiếp mapping — không suy đoán). Nghĩa là: giá trị Fill-Forward capture được từ dòng section-header (`main_window.py:9800-9806`) chính là STT dạng chữ (vd "A") của dòng header đó, còn MỖI dòng dữ liệu bên dưới cũng luôn tự có STT riêng khác rỗng (đó là định nghĩa của 1 data row).

Quan trọng — khác với dự đoán ban đầu ở đặc tả plan (dẫn `bom_parser.py:378-383`, hàm `_parse_section_excel_rows`): hàm đó là pipeline **THDM riêng**, KHÔNG PHẢI hàm BOM dùng (BOM dùng `_parse_sheet`, không có bước merge Fill-Forward nào ở tầng parse cả — `_parse_sheet` trả `df` thô, không cột `ItemType`). Toàn bộ Fill-Forward cho BOM diễn ra Ở TẦNG INSERT (`_generate_bom_details`), không phải ở tầng parse. Tại đó: dòng `main_window.py:9823-9826` ghi đè **VÔ ĐIỀU KIỆN**, không kiểm tra `row_vals` đã có giá trị hay chưa (khác hẳn `ff_always_override` — khái niệm đó chỉ tồn tại ở nhánh THDM). Do đó: giá trị `ItemType` mà Pass 1b tự resolve cho từng dòng từ chính cột STT của dòng đó (xem `EV-S11` — luôn là số, vd "1", "1.1") **luôn luôn bị ghi đè bởi giá trị STT của section-header gần nhất** (vd "A") tại bước insert, miễn là có ít nhất 1 section-header phía trên. Với 94/94 file đều có tối thiểu 1 dòng section-header dạng chữ cái ở mỗi section (`EV-S07`, ví dụ `2511_SYCAMORE_CT01/1. CUA TOILET KHUYET TAT KHU VUC COMMON_DH114a.xlsm` section BOM2 có dòng header STT="A"), kết luận: **việc Pass 1b tự resolve `ItemType` từ STT của chính dòng dữ liệu là code chết (dead code) trên thực tế cho gần như 100% dòng dữ liệu BOM** — kết quả luôn bị ghi đè trước khi insert. Đây là phát hiện cấu trúc quan trọng cho Phase 2: khi port sang `core/`, không cần port bước resolve `ItemType` per-row từ Excel, chỉ cần port đúng cơ chế "STT chữ cái section-header gần nhất → ItemType" (mà bản chất là 1 dạng category/nhóm, không phải Fill-Forward theo nghĩa thông thường).

### EV-S14 — Header meta: ExcelCell Y3 & "Số lượng" (A7)

Đo bằng cách gọi `parse_bom_file(path, meta_keys=meta_keys, cell_specs=cell_specs_header)` với đúng `meta_keys`/`cell_specs` build từ mapping thật (`build_meta_keys_from_mapping`, `build_cell_specs_from_mapping`) — script chính (đo EV-S01..S11) gọi `parse_bom_file(path)` không truyền 2 tham số này nên `global_meta` rỗng ở đó; phải đo lại riêng cho đúng hành vi thật của app.

- **Ô `Y3` (Nguồn_DL=ExcelCell cho field Quantity, `e8a5007`/v2.2.13) đọc được giá trị non-rỗng:** **94/94 file** (ví dụ `2511_SYCAMORE_CT01/1. CUA TOILET KHUYET TAT KHU VUC COMMON_DH114a.xlsm`; khớp mẫu `sample_rows` đã thấy ở `EV-S00`: `'Y3': 3`).
- **Khóa `Số lượng` hiện diện trong `global_meta`:** **94/94 file** (ví dụ file trên — luôn có, dù giá trị có thể xung đột giữa sheet, xem `EV-S12`).

### EV-M01 — Census mapping production qua `load_mapping()`

Đọc trực tiếp `CK_Mapping_v5.xlsx` hiện hành (khớp production v2.2.25) bằng `load_mapping(MAPPING_FILE)`, không qua file mẫu.

**Số record theo section:** HEADER=52, BOM2=35, BOM3=28, BOM4=32, BOM5=30.

**NguonDL theo section:**
| Section | HeThong | Excel | CoDinh | UILookup | SP | ExcelCell |
|---|---|---|---|---|---|---|
| HEADER | 4 | 12 | 29 | 3 | 3 | 1 |
| BOM2 | 8 | 21 | 5 | 1 | 0 | 0 |
| BOM3 | 8 | 14 | 5 | 1 | 0 | 0 |
| BOM4 | 8 | 17 | 6 | 1 | 0 | 0 |
| BOM5 | 8 | 16 | 5 | 1 | 0 | 0 |

**Mac_dinh macro thật gặp** (không suy đoán macro nào không có trong dữ liệu — `NEWID`/`NULL` literal KHÔNG xuất hiện trong mapping hiện hành, dù đặc tả plan liệt kê chúng như macro có thể có): `NOW` (HEADER×4, BOM2-5×2 mỗi section), `EMPTY` (BOM2-5, 2-3 mỗi section), `creator` (mọi section, 1 mỗi section — ứng với macro `CREATOR` viết hoa trong code, `_mac_upper = mac.upper()` nên không phân biệt hoa/thường), `creator_employee` (HEADER×1), `AUTO_INC` (BOM2-5×1 mỗi section), cùng các hằng số cố định (`0`,`1`,`-1`,`2`..`5`,`10`,`A01`,`BOM`,`SX`,`L2`) và tên field parent để copy (`ItemId0`, `DetailRowId_SO`, `EffectiveDate`, `ParentBizDocId`, `bom_product_id`).

**Kieu_lookup theo section:** HEADER — `exact_code`(2), `fuzzy_code`(5), `lookup`(1), `sp_rowid`(1), `sp_version`(1). BOM2 — `nvl_or_btp`(1). BOM3 — `code_then_name`(1), `fuzzy_name`(1). BOM4 — `exact_name`(1). BOM5 — `fuzzy_name`(1).

**Fill_Forward=1:** chỉ `ItemType` ở cả 4 section BOM2-5 (xem `EV-S13`); 0 field ở HEADER.

**Bien_doi:** `UPPER` — đúng 1 field mỗi section BOM2-5 (không có `LOWER`/`TRIM` nào trong mapping hiện hành, dù `73081dc` khai UPPER/LOWER/TRIM transform nói chung — chỉ UPPER thực sự được dùng cho BOM). 0 ở HEADER.

**SP_HOOK đang active (`isactive='1'`):** đúng 3 dòng, cả 3 đều gọi `B10_Boho.dbo.usp_B20BOM_Create_ItemCode` (khớp chuỗi synonym xác nhận ở `SQL-03`), sự kiện `BeforeInsertBatch`, điều kiện `EMPTY(ItemId)`, cho BOM2/BOM3/BOM4 (KHÔNG có hook cho BOM5 — BOM5 không có cơ chế tự tạo mã BTP qua SP). `outputfields` giống nhau cả 3: `ItemId,ItemName,Unit`.

**SP_CONFIG active (`isactive='1'`, không tính THDM):** 3 dòng cho HEADER — `Ws_Id` (lookup `B00Branch`), `RowId` (`usp_sys_CreateSttBySeq`), `Version` (`usp_B20BOM_AutoVersion`). 1 dòng `isactive='0'` (tắt): `Code` → `usp_sys_AutoNewCode` — SP này được khai báo trong mapping nhưng KHÔNG active, khớp với finding #4 của `docs/audit/AUDIT-2026-09-21.md` ("usp_sys_AutoNewCode does not exist or is not accessible") — mapping đã tắt nó từ trước, không phải audit phát hiện ra 1 lỗi mới.

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

### SQL-02 — Trùng mã BTP tái dùng (E9)

Dùng đúng predicate của `_find_existing_btp_code` (`main_window.py:9116-9121`: `IsActive = 1`, `Name`, `Code LIKE item_code0 + '.%'`, `ORDER BY Id DESC`). "item_code0 prefix" = phần Code trước dấu chấm cuối cùng (ví dụ `LO.CH.AS11.21.0005` → prefix `LO.CH.AS11.21`, `.0005` là số thứ tự do USP sinh).

```sql
SELECT TOP 20 Code, Name, IsActive FROM B20Item WHERE Code LIKE '%.%' AND IsActive = 1 ORDER BY Id DESC
```
**Kết quả:** xác nhận đúng mẫu (`LO.CH.AS11.21.0005`, `TP_0004.2.2.0052`, …).

```sql
SELECT COUNT(*) AS collision_groups, SUM(n) AS total_dup_rows
FROM (
  SELECT LEFT(Code, LEN(Code) - CHARINDEX('.', REVERSE(Code))) AS item_code0_prefix, Name, COUNT(*) AS n
  FROM B20Item WHERE IsActive = 1 AND Code LIKE '%.%'
  GROUP BY LEFT(Code, LEN(Code) - CHARINDEX('.', REVERSE(Code))), Name
  HAVING COUNT(*) > 1
) x
```
**Kết quả:** `collision_groups = 477`, `total_dup_rows = 1607`.

```sql
SELECT TOP 10 LEFT(Code, LEN(Code) - CHARINDEX('.', REVERSE(Code))) AS item_code0_prefix, Name, COUNT(*) AS n, MIN(Id) AS min_id, MAX(Id) AS max_id
FROM B20Item WHERE IsActive = 1 AND Code LIKE '%.%'
GROUP BY LEFT(Code, LEN(Code) - CHARINDEX('.', REVERSE(Code))), Name
HAVING COUNT(*) > 1
ORDER BY n DESC
```
**Kết quả (top 10 nhóm trùng, n = số Id active cùng prefix+Name):** `RO.FF.02A.R.1.1` / "" / n=24; `SP_065.1` / "CHI-CANH_03_005" / n=20; `SP_065.1` / "KTHT-HONG_03_001" / n=20; `SP_065.1` / "KTHT-HONG_03_002" / n=20; `SP_065.1` / "KTHT-HONG_04_001" / n=20; `SP_065.1` / "KTHT-HONG_04_002" / n=20; `SP_065.1` / "KTHT-NOC_05_001" / n=20; `SP_065.1` / "P1A-TU" / n=20; `SP_065.1` / "P1B-KHUNGCHAN" / n=20; `SP_065.1` / "P1C-KHUNGDUOI" / n=20.

Với `ORDER BY Id DESC` + `TOP 1`, "tái dùng mã BTP" của E9 luôn chọn Id lớn nhất trong nhóm — xác định (deterministic), khác hẳn G1 (không có `ORDER BY`). 477 nhóm/1607 dòng trùng cho thấy việc "phải chọn" xảy ra thường xuyên, nhưng quy tắc chọn ở đây rõ ràng, không phải rủi ro cùng loại với G1.

### SQL-03 — Metadata usp_B20BOM_Create_ItemCode (H8, F7-F11)

```sql
SELECT name, type_desc FROM sys.objects WHERE name LIKE '%B20BOM_Create_ItemCode%'
```
**Kết quả:** 1 dòng — `name=usp_B20BOM_Create_ItemCode, type_desc=SYNONYM` (trong DB hiện tại `B10_Boho_Data`).

```sql
SELECT name, base_object_name FROM sys.synonyms WHERE name = 'usp_B20BOM_Create_ItemCode'
```
**Kết quả:** `base_object_name = [B10_Boho].[dbo].[usp_B20BOM_Create_ItemCode]` — synonym trỏ sang **DB khác** (`B10_Boho`, không phải `B10_Boho_Data` đang kết nối, cũng không phải `BOMTool`).

```sql
SELECT p.parameter_id, p.name, t.name AS type_name, p.max_length, p.is_output
FROM [B10_Boho].sys.parameters p
JOIN [B10_Boho].sys.types t ON p.user_type_id = t.user_type_id
WHERE p.object_id = OBJECT_ID('[B10_Boho].[dbo].[usp_B20BOM_Create_ItemCode]')
ORDER BY p.parameter_id
```
**Kết quả:** 9 tham số — `@_B20BOMDetail` (xml), `@_B20BOMDetail2` (xml), `@_B20BOMDetail1` (xml), `@_ItemId0` (int), `@_ProductId` (int), `@_BizDocId_SO` (varchar 24), `@_ParentBizDocId` (varchar 524), `@_DetailRowId_SO` (varchar 524), `@_BranchCode` (varchar 24). Khớp đúng tên/thứ tự tham số mà SP_HOOK ở mapping truyền vào (`EV-M01`).

```sql
SELECT OBJECT_DEFINITION(OBJECT_ID('[B10_Boho].[dbo].[usp_B20BOM_Create_ItemCode]')) AS proc_def
```
**Kết quả:** `NULL` — không có quyền `VIEW DEFINITION` trên login hiện tại cho đối tượng này. Không suy đoán nội dung SP khi không xem được định nghĩa.

### SQL-04 — Phân bố EmployeeId trên B20BOM (J1, J2)

Ngày commit `9dc3c92` (v2.2.22) = `2026-09-21` (xác nhận qua `git log -1 --format=%ad --date=short 9dc3c92`). Cột `EmployeeId` (int), `CreatedAt` (datetime) đã xác nhận tồn tại đúng tên qua `INFORMATION_SCHEMA.COLUMNS` trước khi query.

```sql
SELECT CASE WHEN CreatedAt < '2026-09-21' THEN 'before_9dc3c92' ELSE 'on_or_after_9dc3c92' END AS period,
       EmployeeId, COUNT(*) AS n
FROM B20BOM
GROUP BY CASE WHEN CreatedAt < '2026-09-21' THEN 'before_9dc3c92' ELSE 'on_or_after_9dc3c92' END, EmployeeId
ORDER BY period, n DESC
```
**Kết quả:**
- TRƯỚC (2026-09-21): `EmployeeId=1` → 264; `NULL` → 201; `5` → 61; `62189` → 23; `303` → 22; `70529` → 3; `70053` → 2.
- SAU (≥2026-09-21): `EmployeeId=1` → 48; `NULL` → 16; `62189` → 10; `62207` → 4; `62185` → 4; `21387` → 3; `62285` → 1.

`EmployeeId=1` vẫn là giá trị phổ biến nhất SAU khi `9dc3c92` vá, nhưng sau vá đã xuất hiện dải Id nhân viên thật đa dạng hơn (62189/62207/62185/21387/62285) — trước vá gần như chỉ có 1/NULL/5. **Chưa xác nhận** `EmployeeId=1` sau vá là hợp lệ (nhân viên Id=1 có thật) hay vẫn là 1 fallback khi không ai được chọn — xem `Q-02` ở §5.

### SQL-05 — ParentDetailRowId_SO so với công thức Mục số (J4)

Ngày commit `78b36fc` = `2026-09-22`.

```sql
SELECT TOP 20 Id, CreatedAt, ParentDetailRowId_SO, DetailRowId_SO, ParentBizDocId, BizDocId_SO
FROM B20BOM WHERE CreatedAt >= '2026-09-22' ORDER BY Id DESC
```
**Kết quả:** 18 dòng trả về (chưa đủ lịch sử sau-vá để có 20). 17/18 dòng có `ParentDetailRowId_SO == DetailRowId_SO` (đúng theo công thức Mục số). 1 dòng (`Id=687`, tạo `2026-09-22T10:38:43`) có `ParentDetailRowId_SO == ParentBizDocId == "11036174FO"` — đúng mẫu lỗi CŨ, xảy ra SAU commit vá.

```sql
SELECT SUM(CASE WHEN ParentDetailRowId_SO = DetailRowId_SO THEN 1 ELSE 0 END) AS matches_detail_row,
       SUM(CASE WHEN ParentDetailRowId_SO = ParentBizDocId THEN 1 ELSE 0 END) AS matches_order_code,
       SUM(CASE WHEN ParentDetailRowId_SO NOT IN (DetailRowId_SO, ParentBizDocId) THEN 1 ELSE 0 END) AS other,
       COUNT(*) AS total
FROM B20BOM WHERE CreatedAt >= '2026-09-22'
```
**Kết quả (SAU 78b36fc):** `matches_detail_row=17, matches_order_code=1, other=0, total=18` → tỷ lệ sai 1/18 ≈ 5.6%.

Cùng query với `WHERE CreatedAt < '2026-09-22'` (TRƯỚC): `matches_detail_row=529, matches_order_code=49, other=0, total=644` → tỷ lệ sai 49/644 ≈ 7.6%.

Tỷ lệ sai KHÔNG giảm về 0 sau `78b36fc` (7.6% → 5.6%, mẫu sau-vá nhỏ — 18 dòng — nên chưa chắc chắn về thống kê, nhưng dấu hiệu đủ rõ để nêu câu hỏi). Nghi ngờ còn 1 nhánh code khác chưa được `78b36fc` bao phủ (ví dụ luồng THDM, hoặc 1 loại sản phẩm cụ thể) — xem `Q-03` ở §5.

### SQL-06 — Dấu vết import dở dang (K3)

Cột liên kết `B20BOMDetail.BOMId → B20BOM.Id` (xác nhận theo mẫu query THDM có sẵn trong `bom_parser.py` — `_thdm_load_bom_qty_dict` JOIN `B20BOMDetail bd JOIN B20BOM bom ON bom.Id = bd.BOMId`).

```sql
SELECT COUNT(*) AS headers_with_zero_details FROM B20BOM h WHERE NOT EXISTS (SELECT 1 FROM B20BOMDetail d WHERE d.BOMId = h.Id)
```
**Kết quả:** `8`.

```sql
SELECT COUNT(*) AS total_headers FROM B20BOM
```
**Kết quả:** `662` (8/662 ≈ 1.2%).

```sql
SELECT TOP 8 h.Id, h.CreatedAt, h.BizDocId_SO FROM B20BOM h WHERE NOT EXISTS (SELECT 1 FROM B20BOMDetail d WHERE d.BOMId = h.Id) ORDER BY h.CreatedAt DESC
```
**Kết quả:** cả 8 dòng đều CŨ (2026-03-11 → 2026-06-29), không có dòng gần đây — **kết quả tích cực**, không có dấu hiệu import dở dang đang diễn ra. Chi tiết: `Id=345` (2026-06-29 20:50, đơn `21035622FO`) — cùng đơn hàng này có 4 header mồ côi (345,344,343,342, cùng `21035622FO`, gợi ý các lần import thử-lại-thất-bại lặp lại); `Id=285,284` (2026-06-19, đơn `21035472FO`); `Id=270` (2026-06-16, đơn `21035548FO`); `Id=22` (2026-03-11, đơn `11034336FO`). Đối chiếu với commit/rollback ở `main_window.py:10840-10870`/`9937-9957` (đọc ở read_first Task 3): các header mồ côi này nhất quán với 1 import bị lỗi giữa chừng SAU KHI insert header nhưng TRƯỚC KHI insert đủ detail, không có transaction bao trùm cả 2 bước.

### SQL-07 — Mức dùng mã MKT thật (G3, G4, G5)

```sql
SELECT m.ItemTypeSX_Parent, COUNT(*) AS n
FROM B20BOMDetail bd JOIN [BOMTool].[dbo].[vB20Item_MKT] m ON m.Id = bd.ItemId
GROUP BY m.ItemTypeSX_Parent ORDER BY n DESC
```
**Kết quả (toàn bộ lịch sử, số dòng BOM chi tiết thật dùng mã MKT theo từng `ItemTypeSX_Parent`):** `I`=1373, `F`=652, `A`=265, `C`=254, `NULL`=72, `G`=67, `D`=50, `H`=48, `B`=21. Tổng = 2802 dòng.

**Số liệu quyết định mức rủi ro của G1:** các nhóm `C` (254 dòng), `F` (652 dòng), `I` (1373 dòng) và `NULL` (72 dòng) **chính là 4 nhóm trùng khóa** phát hiện ở `SQL-01`. Tổng 254+652+1373+72 = **2351/2802 dòng (84%)** dùng MKT fallback đã đi qua 1 nhóm có collision không xác định thứ tự. Đây là bằng chứng định lượng cho Mức rủi ro = Cao của G1 (không phải rủi ro lý thuyết — là đa số lưu lượng MKT fallback thật). Chi tiết theo từng Id trong nhóm C và NULL đã có ở `SQL-01(d)`.

### SQL-08 — Trùng tên B20Item (E8, E12)

```sql
SELECT COUNT(*) AS dup_name_groups, SUM(n) AS total_rows
FROM (SELECT Name, COUNT(*) AS n FROM B20Item WHERE IsActive = 1 GROUP BY Name HAVING COUNT(*) > 1) x
```
**Kết quả:** `dup_name_groups = 7192`, `total_rows = 22519`.

```sql
SELECT TOP 10 Name, COUNT(*) AS n FROM B20Item WHERE IsActive = 1 GROUP BY Name HAVING COUNT(*) > 1 ORDER BY n DESC
```
**Kết quả (top 10):** "GÓI 1/1 SHOR DINING CHAIR"=278; "GÓI 1/1 Chip White Oak Dining Chair - Camp Olive"=166; "GÓI 1/4 MD1-TỦ DƯỚI"=138; "GÓI 1/2 Mặt bàn"=125; "GÓI 2/2 Khung chân"=125; "GÓI 1/1 Chip White Oak Dining Chair - Camp Stone"=116; "GÓI 1/1 Chip White Oak Barstool - Camp Navy"=81; "GÓI 1/1 Chip White Oak Dining Chair - Camp Salt"=80; "GÓI 1/1 Chip Walnut Dining Chair - Camp Navy"=74; "Chỉ khung bao ngang"=71.

Đây là số trùng tên TOÀN DANH MỤC (population rộng — nhiều sản phẩm dùng chung tên linh kiện phổ thông như "Chỉ khung bao ngang"), khác quy mô với `SQL-02` (477 nhóm đã scope theo `item_code0` prefix — nhóm nhỏ hơn, đúng ngữ cảnh tái dùng mã BTP của 1 sản phẩm cụ thể). Không gộp 2 số liệu này làm một. Đây chính là population mà logic tie-break `1694b70` (§1) được viết ra để xử lý.

## 4. Bảng kiểm kê & verdict (D-05)

### 4.A Nhận diện sheet/section

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **A1** `_load_sheet_config` (`bom_parser.py:57-98`): đọc `_CONFIG` từ mapping file **tươi mỗi lần parse** (không cache); `PermissionError`/`Exception` → log qua `logging.warning` (dùng module `logging` chuẩn, KHÔNG nuốt lỗi) rồi trả `[]` | GIỮ | `bom_parser.py:57-98`; fresh-reload cho phép thêm section BOM6+ không cần restart app; nhánh lỗi CÓ gọi `_log.warning(...)` (đối lập với claim "0 lần gọi logger" của `main_window.py` trong PROJECT.md — file này khác); `EV-S00` (94/94 file parse không lỗi ở bước này), `EV-M01` (census đọc được `_CONFIG` bình thường) | - | Trung bình |
| **A2** `_detect_sheet_type` (`bom_parser.py:100-126`): match tên sheet bằng regex word-boundary (`contains` phải khớp hết, `excludes` không được khớp cái nào) → section; không khớp → `UNKNOWN` | GIỮ | `bom_parser.py:100-126`; `EV-S02` (sheet `BOM_Phan I` luôn UNKNOWN ở 94/94 file vì `_CONFIG` không đăng ký "I" — đúng thiết kế, không phải bug) | - | Thấp |
| **A3** `parse_bom_file` (`bom_parser.py:766-769`): sheet `hidden`/`veryHidden` bị `continue` TRƯỚC bước phân loại — không bao giờ vào `skipped_sheets` | GIỮ | `bom_parser.py:766-769`; `EV-S02` (sheet tham chiếu ẩn `TC_CUA`, `TC_KH`, `LOẠI 1/4/5/6` luôn hidden ở 94/94 file, đúng ý đồ — chỉ là sheet tra cứu phụ trợ, không phải section dữ liệu) | - | Thấp |
| **A4** `parse_bom_file` (`bom_parser.py:786-790`): `global_meta` gộp theo quy tắc **"sheet đầu tiên thắng"** (`if _k not in global_meta: global_meta[_k] = _v`) — không phân biệt sheet nào là nguồn đúng cho từng khóa | SỬA | `bom_parser.py:786-790`; `EV-S12` — 229 xung đột trên 93/94 file; ví dụ thật: khóa `Số lượng` = "1 Hệ" (BOM2, thắng) nhưng BOM5 có giá trị khác nghĩa thật "9" bị bỏ qua hoàn toàn | Thay "sheet đầu tiên thắng" toàn cục bằng resolve theo khóa: mỗi khóa meta (`Kích thước`, `Số lượng`, `Hoàn thiện`, `Vật liệu`) khai rõ **section nguồn ưu tiên** (không phải luôn BOM2) trong `_CONFIG` hoặc mapping — hiệu ứng quan sát được: `EV-S12`'s 229 xung đột giảm về số case thật cần cảnh báo cho business thay vì âm thầm chọn sai. Section nào là nguồn đúng cho khóa nào là quyết định nghiệp vụ — xem `Q-06` ở §5. Requirement: MỚI | Cao |
| **A5** `parse_bom_file` (`bom_parser.py:815-825`): mọi section trong `_CONFIG` không tìm thấy sheet khớp đều nhận cảnh báo `"Không tìm thấy sheet …"` như nhau — không phân biệt section bắt buộc/tùy chọn | SỬA | `bom_parser.py:815-825`; `EV-S01` — `BOM5` thiếu ở 14/94 file (0/94 thiếu BOM2/3/4), nhận cảnh báo y hệt section bắt buộc dù dữ liệu mẫu cho thấy BOM5 là tùy chọn trong thực tế | Thêm cờ "bắt buộc"/"tùy chọn" cho từng section trong `_CONFIG` (cột mới, ví dụ `Required`); chỉ đẩy `warnings.append(w)` khi section bắt buộc thiếu — section tùy chọn (đề xuất: BOM5) chỉ hiện trong UI như thông tin, không phải cảnh báo. Section nào tùy chọn ngoài BOM5 là câu hỏi nghiệp vụ — xem `Q-05` ở §5. Requirement: FIX-04 | Trung bình |
| **A6** `parse_bom_file` (`bom_parser.py:724-760`): workbook mở **3 lần** riêng biệt (cached `wb`, formula-aware `wb_live`, hidden-state `_wb_hid`); `except Exception: wb_live=None` (735) và `except Exception: _hidden_rows_map/_hidden_cols_map={}` (757) đều KHÔNG gọi logger | SỬA | `bom_parser.py:724-760`; `EV-S00` (94/94 file hiện tại parse không lỗi — chưa có bằng chứng crash, nhưng khớp `CONCERNS.md` "Large Excel Files Loaded Into Memory Twice" và là điểm OBS-01 vi phạm) | Thêm `log_fn`/`logging.warning` vào 2 khối `except` (735, 757) trước khi fallback `None`/`{}`, để lỗi mở lần 2/3 (file hỏng, hết bộ nhớ) không biến mất hoàn toàn; gộp 3 lần mở thành tối đa 2 (cached + 1 lần đọc kết hợp formula+hidden-state) là cải tiến hiệu năng riêng, không bắt buộc cho Phase 2 — chỉ phần logging là correctness-relevant. Requirement: OBS-01 | Trung bình |
| **A7** Trích xuất meta đầu phiếu: `_resolve_formula` (128-143), `_merge_meta_rows` (144-161), `_extract_meta` (162-188), `_extract_cell_meta` (189-211), đọc 15 dòng đầu mỗi sheet (770-781), ô `Y3` cố định cho `Quantity` (`Nguon_DL=ExcelCell`, commit `e8a5007`/v2.2.13) | GIỮ | `bom_parser.py:128-211`, `770-781`; `EV-S14` — ô `Y3` đọc được giá trị 94/94 file, khóa `Số lượng` hiện diện trong `global_meta` 94/94 file | - | Thấp |

### 4.B Nhận diện header

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **B1** `_find_header_row` (`bom_parser.py:213-219`): dòng đầu tiên có ≥2 hit với `HEADER_ANCHORS` (`"MÃ SP","STT","Tên chi tiết","Tên Vật Tư","Mã chi tiết","Mã vật tư"`) được chọn làm header row | GIỮ | `bom_parser.py:213-219`; `EV-S00` (94/94 file parse ra `df` có cột thật, hàm ý header row được tìm thấy đúng ở mọi file — nếu sai, section đó sẽ rơi vào nhánh `hidx is None` và trả `df` rỗng, không khớp baseline) | - | Thấp |
| **B2** `_build_headers` (`bom_parser.py:233-239`): forward-fill nhãn dòng r1 qua merged cell **không giới hạn ranh giới phải** — cột tính toán/ẩn nằm cạnh bảng thật cũng bị "thừa hưởng" nhãn gần nhất | SỬA | `bom_parser.py:233-239`; `EV-S03` — 362 lượt cột không xác định (`COL_N`)/orphan trên 94/94 file; hầu hết `COL_N`/toàn-rỗng đã tự dọn ở `_parse_sheet:585-586`, nhưng cột **có tên thật do forward-fill** (vd `SLSP`, `SLSP_LỌC_IN`, `SLSP_PLDG` — lặp lại nhất quán ở BOM2/3/4/5) KHÔNG bị dọn vì có giá trị thật, không rỗng, không phải `COL_*`; `EV-S04` xác nhận các tên này 0/94 khớp bất kỳ `Ten_Excel` nào → rác preview, không lẫn vào INSERT (không mapping row nào tham chiếu chúng) | Gắn provenance-flag cho mỗi cột khi build header (`filled=True` nếu tên đến từ forward-fill, không phải nhãn gốc của chính cột đó); sau khi build, áp dụng thêm điều kiện drop tại `_parse_sheet:585-586`: bỏ cột `filled=True` VÀ không khớp bất kỳ `Ten_Excel` nào trong mapping section đó (tái dùng đúng cơ chế match 2-pass đã có ở `_resolve_detail_row`). Hiệu ứng quan sát được: `EV-S03`'s cột orphan-do-forward-fill giảm về 0 trong khi cột merged-header hợp lệ (đã khai trong `Ten_Excel`, ví dụ `SLg_Tên_Vật_Tư`) vẫn resolve bình thường như `EV-S04` ghi nhận. Requirement: FIX-02 | Trung bình |
| **B3** `_build_headers` (`bom_parser.py:241-248`): ghép `v1_v2` khi cả 2 dòng có nhãn; fallback `COL_{i}` khi không xác định được nhãn nào | GIỮ | `bom_parser.py:241-248`; `EV-S03` (362 lượt `COL_N` trên 94/94 file — đúng như kỳ vọng khi Excel có cột không nhãn, các cột này bị dọn ở `_parse_sheet:585-586`) | - | Thấp |
| **B4** `_build_headers` (`bom_parser.py:249-257`): dedupe tên trùng bằng hậu tố `_N` | GIỮ | `bom_parser.py:249-257`; `EV-S03` (157 lượt cột nhận hậu tố dedupe, ví dụ `Số_lượng_1`, `SLSP_LỌC_IN_1`, trên 94/94 file — cơ chế cần thiết để tránh đè khóa dict ở `_parse_sheet:519`) | - | Thấp |
| **B5** Header Excel → `sql_col` cho BOM **KHÔNG** đi qua `_build_excel_col_map`/`_norm_col` (`bom_parser.py:221-222, 270-294`) — 2 hàm này chỉ được gọi 1 nơi duy nhất trong toàn repo (`bom_parser.py:1094`, luồng THDM). Cơ chế thật của BOM là match 2-pass **inline, lặp lại theo từng dòng dữ liệu** trong `_resolve_detail_row` (`main_window.py:9228-9254`, xem `E3`/`E4`) | SỬA | `bom_parser.py:221-222, 270-294` (không dùng cho BOM — xác nhận bằng grep toàn bộ call site), `main_window.py:9228-9254` (cơ chế thật); `EV-S04` (71/94 file cần suffix-pass mới khớp `Tên vật tư`; các `Ten_Excel` không khớp header thật nào, vd BOM5 `Diện tích...Sơn hoàn thiện`, 80/94) | Trích cơ chế match 2-pass hiện có trong `_resolve_detail_row` thành 1 hàm build-column-map DÙNG CHUNG cho BOM (tương tự `_build_excel_col_map` đã có cho THDM), chạy 1 lần/section thay vì lặp lại cho mỗi dòng × mỗi mapping record. Không đổi kết quả khớp (vẫn đúng những gì `EV-S04` đã đo), chỉ gom logic về 1 chỗ — đây cũng là bước cần thiết để B2's provenance-flag hoạt động. Requirement: ARCH-01 | Trung bình |

### 4.C Parse dòng dữ liệu

**Lưu ý baseline (áp dụng cho toàn bộ §4.C):** đặc tả kế hoạch trích dẫn `_parse_section_excel_rows` (`bom_parser.py:297-407`) làm nguồn cho nhóm C — hàm này đã được `01-01-SUMMARY.md`/`EV-S13`/`EV-S08` xác nhận là **THDM-only** (gọi duy nhất tại `bom_parser.py:1143`, luồng THDM). Hàm BOM thật là `_parse_sheet` (`bom_parser.py:461-588`), đọc lại và trích dẫn dưới đây; các dòng có khác biệt với mô tả gốc của đặc tả đều ghi rõ.

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **C1** Bỏ dòng hoàn toàn trống: `_parse_sheet:516-517` (`if all(c is None for c in row): continue`) | GIỮ | `bom_parser.py:516-517`; `EV-S00` (94/94 file parse ổn định, nhất quán với việc dòng trống bị loại đúng trước khi vào các bước sau) | - | Thấp |
| **C2** `FOOTER_KEYWORDS` + dừng-sau-footer: `_parse_sheet:548-555` (khớp từ khóa case-insensitive trên TOÀN BỘ text nối của cả dòng → set cờ `_footer_seen` → mọi dòng phía sau bị bỏ vĩnh viễn, kể cả có dữ liệu thật); từ khóa `"duyệt"` đứng riêng đã bị loại bỏ bởi `817bbbe`/v2.2.15 sau khi gây mất dữ liệu hàng loạt | SỬA | `bom_parser.py:548-555`, `bom_parser.py:48-53` (danh sách `FOOTER_KEYWORDS` hiện tại); `EV-S06` — 2.573 dòng dữ liệu thật bị loại vĩnh viễn trên 94/94 file (~27 dòng/file, ví dụ `2511_SYCAMORE_CT01/...xlsm` BOM2); commit `817bbbe` (§1) — cùng cơ chế match-substring-toàn-dòng này đã từng gây "mất dữ liệu hàng loạt" thật với từ khóa `"duyệt"` | **[Cập nhật theo Q-08, business owner đã chốt 2026-09-25 — chấp nhận #3]** Không tiếp tục đoán-theo-từ-khóa (kể cả bản thu hẹp điều kiện). Thay bằng: template Excel bổ sung 1 dấu hiệu kết-thúc-bảng dữ liệu TƯỜNG MINH (ví dụ: 1 ô/hàng đặc biệt đánh dấu "HẾT DỮ LIỆU", hoặc khai số dòng dữ liệu tối đa trong `_CONFIG` per-section) — parser dừng đọc đúng tại dấu hiệu đó, không cần khớp `FOOTER_KEYWORDS` nữa. Nếu file thiếu dấu hiệu này → REJECT rõ ràng, yêu cầu user bổ sung, không tự đoán rồi im lặng cắt dữ liệu như hiện tại. `EV-S06`'s 2.573 dòng là baseline để xác nhận không còn dòng thật nào bị cắt oan sau khi đổi cơ chế. Thiết kế cụ thể dấu hiệu này (hình thức, vị trí) do Phase 2 quyết khi làm việc với đội vận hành file mẫu. Requirement: MỚI | Cao |
| **C3** GIAO RỜI: dòng STT chữ cái có kích thước vật lý (DÀY/RỘNG/DÀI) → xử lý như data row, không phải section header: `_parse_sheet:530-544` (`_has_dims` check) | GIỮ | `bom_parser.py:530-544`; `EV-S07` — 16 dòng GIAO RỜI trên 8/94 file, toàn bộ nhận diện đúng (STT="I", 1 file có cả "I" và "II") | - | Thấp |
| **C4** Dòng section-header chữ cái (A, B, C...): `_parse_sheet:538-543` — được **append thẳng vào `rows_out`/DataFrame** (không capture Fill-Forward ngay tại đây, không bị loại khỏi df ở bước parse) | GIỮ | `bom_parser.py:538-543`; `EV-S07` (1.083 dòng section-header trên 94/94 file), `EV-S13` (xác nhận Fill-Forward thật của BOM chạy Ở TẦNG INSERT — `_generate_bom_details`, không phải parse) — **khác mô tả gốc của đặc tả** (trích `_parse_section_excel_rows:368-375`, hàm THDM có capture Fill-Forward ngay tại parse và không insert dòng header); `_parse_sheet` không làm vậy — dòng header vẫn nằm trong `df` cho tới khi bị lọc ở `main_window.py:9786-9806` (xem `I2`/`I3`) | - | Thấp |
| **C5** STT số La Mã (I, II...) luôn fall-through như data row: `_parse_sheet:538-540` (`if _is_roman_numeral(stt): pass`) — **không có tham số `roman_as_hdr`** như bản THDM (`_parse_section_excel_rows`), Roman luôn là data row cho BOM, không có cách bật ngược lại | GIỮ | `bom_parser.py:538-540`; `EV-S07` — 49 dòng trên 31/94 file, toàn bộ "I"/"II", không có trường hợp nào cần xử lý như header trong dữ liệu mẫu | - | Thấp |
| **C6** Fill-Forward tại tầng parse: **không tồn tại trong `_parse_sheet`** cho BOM — điều kiện `ff_always_override or not rd.get(sql_col)` mà đặc tả trích (`bom_parser.py:378-383`) chỉ có trong `_parse_section_excel_rows` (THDM-only, xem lưu ý baseline) | GIỮ | `bom_parser.py:378-383` (code tồn tại nhưng KHÔNG chạy cho BOM), `EV-S13` (xác nhận BOM không merge Fill-Forward ở tầng parse; cơ chế thật nằm ở tầng insert, xem `I3`/`I4`) | - | Thấp |
| **C7** Bỏ dòng không có identity: `_parse_sheet:577-578` (`if not _is_data_row(rd, key_cols): continue`), `key_cols` = mọi header chứa `"Tên/Mã/STT/Name/Code/Item"` (`bom_parser.py:510-511`) — rộng hơn mô tả gốc của đặc tả (chỉ `ItemId`/`ItemName`, đúng cho THDM `skip_no_item`, không đúng cho BOM) | GIỮ | `bom_parser.py:510-511, 577-578`; `EV-S08` — 169.486 dòng bị loại trên 94/94 file, xác nhận 0/169.486 dòng có bất kỳ giá trị Tên/Mã/STT thật nào (100% noise định dạng Excel used-range, không phải mất dữ liệu thật) | - | Thấp |
| **C8** Rule chữ-hoa-toàn-bộ ("VẬT TƯ CHÍNH" → bỏ qua trừ khi STT số): **không tồn tại trong `_parse_sheet`** — chỉ có ở `_parse_section_excel_rows:396-403` (THDM-only, xem lưu ý baseline) | GIỮ | `bom_parser.py:396-405` (code tồn tại nhưng KHÔNG áp dụng cho BOM); `EV-S08` — xác nhận trực tiếp bằng cấu trúc code (0/94, không phải thiếu mẫu) rằng BOM không có rule này; dòng tên viết-hoa-toàn-bộ có dữ liệu thật đi qua bình thường như mọi data row khác (qua `C7`) | - | Thấp |
| **C9** `_is_data_row` (`bom_parser.py:260-266`): coi chuỗi `'0'`/`'None'`/`'nan'`/rỗng là "không có giá trị" ở các cột key | GIỮ | `bom_parser.py:260-266`; `EV-S08` (0 false-positive trên dữ liệu Tên/Mã thật), `EV-S05` (dòng STT=0 đã bị loại sớm hơn ở `bom_parser.py:546-547` trước khi tới `_is_data_row`, nên nhánh `'0'` ở đây chủ yếu bảo vệ các cột Mã/Tên có giá trị literal "0", không phải STT) | - | Thấp |
| **C10** Chuẩn hóa STT chống integer-0: **3 bản triển khai độc lập** — `bom_parser.py:521-524` (trong `_parse_sheet`), `main_window.py:9624-9629` (`_get_stt_pre`, dùng cho BTP detection), `main_window.py:9771-9776` (`_get_stt`, dùng cho vòng lặp insert chính) — cả 3 logic giống hệt nhau (`"0" if raw in (0,0.0) else "" if rỗng else str(raw).strip()`) nhưng KHÔNG dùng chung code | SỬA | `bom_parser.py:521-524`, `main_window.py:9624-9629`, `main_window.py:9771-9776`; `EV-S05` — 352/470 section có cột STT mixed-dtype (int/float/str trộn), giải thích vì sao cả 3 nơi đều phải tự phòng thủ riêng trước cùng 1 rủi ro dữ liệu | Trích thành 1 hàm dùng chung `normalize_stt(raw) -> str` (ứng viên tự nhiên cho `core/sanitizer.py` cùng `safe_str`/`safe_int`/`safe_date`), gọi từ cả 3 điểm. Hiệu ứng quan sát được: không đổi kết quả (vẫn khớp `EV-S05`'s 352/470 section) nhưng loại bỏ rủi ro 3 bản sao lệch nhau sau lần sửa kế tiếp — đúng lớp lỗi mà `cb2c1d7` (§1, STT trùng giữa section) đã từng phải vá vì 1 chỗ tính STT không nhất quán với chỗ khác. Requirement: ARCH-04 | Trung bình |

### 4.D Resolve theo NguonDL (_resolve_row_mapping)

**Lưu ý baseline (áp dụng cho §4.D):** `_resolve_row_mapping` (`bom_parser.py:1259-1379`) được gọi ở đúng 2 nơi trong repo — `main_window.py:6588` (THDM, `get_excel_val=lambda rec: excel_row.get(...)`) và `main_window.py:9170` (BOM, qua `_resolve_detail_row`, với `mapping_recs=_non_excel` đã LỌC BỎ mọi record `nguon_dl=='Excel'` VÀ `get_excel_val=None`). Do đó nhánh `nguon=='Excel'` của hàm này (1300-1329) **không bao giờ thực thi có ý nghĩa cho BOM** — cơ chế Excel thật của BOM nằm hoàn toàn ở `_resolve_detail_row` Pass 1b riêng (`main_window.py:9172-9418`, xem `E5`/`E6`/`E7`). Các nhánh `CoDinh`/`HeThong`/`UILookup`/`MucLookup` VẪN thực thi cho BOM (đều nằm trong `_non_excel`).

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **D1** `skip_nguon` mặc định `{'SP','TinhToan'}` (`bom_parser.py:1283, 1295-1297`): record có `nguon_dl` nằm trong set này → `row_out.setdefault(sql_col, None)` rồi bỏ qua | GIỮ | `bom_parser.py:1283, 1295-1297`; `EV-M01` — BOM2-5 DETAIL có 0 record `nguon_dl='SP'` (bảng NguonDL theo section) và 0 record `TinhToan` ở bất kỳ section nào → nhánh này hiện là no-op cho BOM (không có gì để skip), nhưng vẫn đúng/an toàn nếu mapping tương lai thêm field SP/TinhToan vào BOM DETAIL | - | Thấp |
| **D2** `Excel` + macro `Mac_dinh` fallback (EMPTY/NOW/CREATOR/literal auto-cast): `bom_parser.py:1300-1319` | BỎ | `bom_parser.py:1300-1319`; xác nhận bằng grep toàn repo — `_resolve_row_mapping` chỉ có 2 call site, chỉ THDM (`main_window.py:6588-6591`) truyền `get_excel_val`; BOM (`main_window.py:9170`) lọc bỏ record Excel VÀ truyền `get_excel_val=None` → nhánh 1300-1329 không chạy cho BOM; `EV-M01` (BOM2-5 có 21/14/17/16 record `Excel` — toàn bộ resolve qua `_resolve_detail_row` Pass 1b, không qua đây) | Lý do bỏ: nhánh này chỉ phục vụ THDM (`_resolve_thdm_thvt_row`), không có dữ liệu BOM nào đi qua. Thay thế: khi port `_resolve_row_mapping` (subset BOM) vào `core/`, không cần mang nhánh `Excel` theo — cơ chế Excel thật của BOM là Pass 1b trong `_resolve_detail_row`, xem `E5`/`E6`/`E7` để có Cách sửa thật cho phần này. THDM giữ nguyên bản riêng, không đụng | Thấp |
| **D3** `Excel` → `Bien_doi` UPPER/LOWER/TRIM: `bom_parser.py:1320-1327` | BỎ | `bom_parser.py:1320-1327`; cùng bằng chứng call-site như `D2` — nhánh nằm trong cùng khối `if nguon == 'Excel':` không chạy cho BOM; `EV-M01` (`Bien_doi=UPPER` có 1 field/section BOM2-5 — áp dụng thật qua `_resolve_detail_row:9294-9305`, xem `E7`, không qua đây) | Lý do bỏ: cùng lý do `D2` — nhánh Excel/Bien_doi của `_resolve_row_mapping` là THDM-only. Thay thế: `Bien_doi` cho BOM đã có cơ chế riêng ở `_resolve_detail_row:9294-9305` (`E7`) — không cần port nhánh này | Thấp |
| **D4** `CoDinh` NULL/EMPTY/int/float/str auto-cast: `bom_parser.py:1332-1341` | GIỮ | `bom_parser.py:1332-1341`; `EV-M01` — BOM2-5 có 5-6 record `CoDinh`/section, đều đi qua `_non_excel` nên nhánh này THẬT SỰ thực thi cho BOM (khác `D2`/`D3`) | - | Thấp |
| **D5a** `HeThong`: `mac=='NOW'` hoặc `sql_col in ('CreatedAt','ModifiedAt')` → `now` | GIỮ | `bom_parser.py:1346`; `EV-M01` (macro `NOW` gặp thật: HEADER×4, BOM2-5×2/section) | - | Thấp |
| **D5b** `HeThong`: `mac=='AUTO_INC'` → `builtin_order` | GIỮ | `bom_parser.py:1348`; `EV-M01` (macro `AUTO_INC` gặp ở BOM2-5×1/section) | - | Thấp |
| **D5c** `HeThong`: `sql_col in ('BOMId','BizDocId')` hoặc `mac=='NEWID'` → `doc_id` (placeholder `'__BOMID__'`, thay thế thật lúc sinh SQL, xem `_sv` tại `main_window.py:9911`) | GIỮ | `bom_parser.py:1350`, `main_window.py:9911`; `EV-M01` (mapping không có macro `NEWID` literal nào thật, nhánh này chủ yếu phục vụ `sql_col in ('BOMId','BizDocId')`) | - | Thấp |
| **D5d** `HeThong`: `sql_col in parent_fields` → copy `parent_row.get(sql_col)` — `parent_fields` cho BOM = `PARENT_FIELDS` (`main_window.py:9147-9150`: `BranchCode`,`EffectiveDate`,`FinishedDate`,`ItemId0`,`DetailRowId_SO`,`ParentBizDocId`,`ProductionProcessId`) `\| {'BOMDetailType'}` | GIỮ | `bom_parser.py:1352`, `main_window.py:9147-9150, 9162`; `EV-M01` (các field parent-copy này — `ItemId0`, `DetailRowId_SO`, `ParentBizDocId`, `bom_product_id` — là macro `Mac_dinh` thật gặp trong mapping production); đúng giá trị copy phụ thuộc HEADER đã resolve đúng trước đó — `DetailRowId_SO`/`ParentBizDocId` đúng hay sai do `J4`/`J5` quyết định, không phải bản thân cơ chế copy này | - | Thấp |
| **D5e** `HeThong`: `mac` khớp tên field của `parent_row` → copy giá trị field đó (`mac` dùng như tên field, không phải literal) | GIỮ | `bom_parser.py:1354-1356`; `EV-M01` (macro đặt tên field cha gặp thật: `ItemId0`, `DetailRowId_SO`, `EffectiveDate`, `ParentBizDocId`, `bom_product_id`) | - | Thấp |
| **D5f** `HeThong`: `mac` là literal cố định (không thuộc `NULL/EMPTY/NOW/AUTO_INC/NEWID`) → giá trị fallback cứng | GIỮ | `bom_parser.py:1357-1358`; `EV-M01` (hằng số cố định gặp thật: `0,1,-1,2..5,10,A01,BOM,SX,L2`) | - | Thấp |
| **D5g** `HeThong`: else → `None` | GIỮ | `bom_parser.py:1359-1360`; `EV-M01` (mọi record `HeThong` production khớp 1 trong các nhánh `D5a`-`D5f` — không ghi nhận field `HeThong` nào thật rơi vào nhánh else); hành vi mặc định an toàn khi không khớp nhánh nào ở trên | - | Thấp |
| **D6** `UILookup`: `row_out[sql_col] = ui_values.get(mac)` | GIỮ | `bom_parser.py:1364-1366`; `EV-M01` (BOM UILookup=1/section, `mac='creator'`); `_ctx['ui_values']` của BOM chỉ truyền key `'creator'` (`main_window.py:9163-9165`, xem `E1`) — khớp đúng, không có `UILookup` nào khác được BOM khai báo | - | Thấp |
| **D7** `MucLookup`: đọc raw Muc key từ Excel row, chỉ được set bởi `_thdm_expand_muc_rows` (THDM) | BỎ | `bom_parser.py:1371-1377`; `EV-M01` — bảng NguonDL theo section không liệt kê `MucLookup` ở HEADER/BOM2-5 (0 record) → không có dữ liệu BOM nào cần nhánh này | Lý do bỏ: `MucLookup` là cơ chế THDM thuần (mở rộng dòng Mục), 0 field BOM/HEADER tham chiếu theo `EV-M01`. Thay thế: không port nhánh `MucLookup` vào `core/` khi trích BOM-scope của `_resolve_row_mapping`; giữ nguyên trong pipeline THDM hiện tại (ngoài phạm vi, không đụng) | Thấp |
| **D8** `TinhToan`: bị skip tại `D1` (`skip_nguon`); không có bước nào trong luồng BOM (`_resolve_detail_row` Pass 2, `main_window.py:9420-9481`) tính lại field `TinhToan` sau đó — Pass 2 chỉ xử lý `nguon_dl=='SP'` (`main_window.py:9425, 9436`) | GIỮ | `main_window.py:9420-9481, 8958` (nhánh `TinhToan` chỉ tồn tại ở `_resolve_header_field`/HEADER, resolve trong "pass 2" — nhưng đó là comment cho `SP`, không có code case riêng cho `TinhToan`); `EV-M01` — không có record `TinhToan` nào ở bất kỳ section BOM/HEADER hiện hành; grep toàn `main_window.py` xác nhận không có bước tính `TinhToan` nào sau `_resolve_row_mapping` cho BOM | - | Thấp |

### 4.E Resolve dòng chi tiết BOM (_resolve_detail_row)

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **E1** Pass 1a (`main_window.py:9152-9170`): field non-Excel → `_resolve_row_mapping(_non_excel, _ctx, get_excel_val=None)` với `_bom_parent` (parent_row + `BOMDetailType`); `ctx['ui_values']` chỉ truyền key `'creator'` | GIỮ | `main_window.py:9152-9170`; `EV-M01` (BOM UILookup=1/section, `mac='creator'` — khớp đúng key duy nhất được truyền, xem `D6`) | - | Thấp |
| **E2** Compound key `Col1\|Col2\|@ParentField` (`main_window.py:9191-9226`): Pass 1 exact/norm-exact, Pass 2 suffix-match — đã vá thiếu Pass 2 bởi commit `3e679fc` | GIỮ | `main_window.py:9191-9226`; `EV-S04` (cơ chế suffix-match đã kiểm chứng đúng); commit `3e679fc` (§1, Chuỗi đổi hướng liên quan `D`/`E`/`G`) | - | Thấp |
| **E3** Single-key Pass 1 exact/norm-exact (`main_window.py:9227-9240`) | GIỮ | `main_window.py:9227-9240`; `EV-S04` | - | Thấp |
| **E4** Single-key Pass 2 suffix-match cho header bị merge dưới nhãn cha (`main_window.py:9243-9254`) | GIỮ | `main_window.py:9243-9254`; `EV-S04` — 71/94 file cần Pass 2 để khớp `SLg_Tên_Vật_Tư` → `Tên vật tư`; không ghi nhận trường hợp false-positive nào trong 94 file mẫu (mọi match Pass 2 đều đúng theo kiểm tra chéo), nhưng cơ chế "first-suffix-match-wins" chưa được đo với 2 header cùng suffix-match 1 target — rủi ro lý thuyết, chưa xác nhận | - | Trung bình |
| **E5** Excel rỗng → macro `Mac_dinh` (`main_window.py:9261-9272`): `EMPTY` luôn gán `raw=''` — KHÔNG kiểm tra `kd in ('date','datetime')` như `_resolve_row_mapping`/`D2` (vốn trả `None` cho kiểu ngày) | SỬA | `main_window.py:9261-9272` (Pass 1b thật của BOM) so với `bom_parser.py:1309-1310` (`D2`, hiện không chạy cho BOM nhưng là đặc tả gốc cùng ý nghĩa); `EV-M01` (macro `EMPTY` gặp thật BOM2-5×2-3/section) — 2 triển khai KHÔNG khớp nhau ở đúng điểm `D2` xử lý riêng cho kiểu ngày | Thêm kiểm tra `kd in ('date','datetime')` vào nhánh `EMPTY` của Pass 1b — trả `None` thay vì `''` cho field kiểu ngày, đồng nhất với `_resolve_row_mapping`. Hiệu ứng quan sát được: loại bỏ khả năng 1 field ngày có macro `EMPTY` đi thẳng chuỗi rỗng `''` vào tham số pyodbc (đường export SQL đã an toàn qua `_sv`, nhưng đường insert trực tiếp — `main_window.py:9946-9949` — truyền `row_vals[c]` thẳng, không qua `_sv`). Requirement: MỚI | Trung bình |
| **E6** Ép kiểu (date qua `_DATE_FMTS`, number, int) — thất bại thì `raw=None` **âm thầm, không log** (`main_window.py:9276-9292`) | SỬA | `main_window.py:9276-9292`; `EV-S11` (giá trị non-str thật sự chảy tới vùng code này từ Excel — xác nhận nhánh ép kiểu có dữ liệu thật đi qua, không phải lý thuyết) | Thêm log (`log_fn`/`self._log`, theo đúng mẫu đã có ở `_log_bom_lookup_fail`) khi mọi định dạng ép kiểu đều thất bại, trước khi gán `raw=None` — ghi rõ field, giá trị gốc, section, dòng. Hiệu ứng quan sát được: field lặng lẽ thành `NULL` trong Bravo hiện tại có nguồn gốc, thay vì không truy vết được (đúng lớp lỗi mà `45031a8` từng phải thêm traceback để chẩn đoán `8a07639`). Requirement: OBS-01, OBS-02 | Trung bình |
| **E7** `Bien_doi` (UPPER/LOWER/TRIM) áp dụng cho BOM (`main_window.py:9297-9304`) — đã vá bởi commit `1954780` (trước đó khai báo nhưng không chạy) | GIỮ | `main_window.py:9297-9304`; `EV-M01` (`Bien_doi=UPPER`, 1 field/section BOM2-5); commit `1954780` (§1) | - | Thấp |
| **E8** `code_then_name` lookup — ưu tiên Code (unique) trước khi fallback Name (`main_window.py:9308-9333`), cache riêng theo `bang_master` dựng ở `_build_bom_detail_caches` (`main_window.py:9080-9091`) | GIỮ | `main_window.py:9080-9091, 9308-9333`; `SQL-08` (7.192 nhóm trùng tên B20Item, xác nhận vì sao ưu tiên Code trước là thiết kế đúng); commit `8354cef`, `7166b36` (§1) | - | Thấp |
| **E9** `nvl_or_btp` nhánh BTP → `_find_existing_btp_code` (`main_window.py:9101-9125`, gọi tại `9336-9352`) — `ORDER BY Id DESC` + `TOP 1`, xác định (deterministic) | GIỮ | `main_window.py:9101-9125, 9336-9352`; `SQL-02` — 477 nhóm/1.607 dòng trùng `item_code0` prefix, luôn chọn `Id` lớn nhất một cách nhất quán (khác hẳn `G1` — có `ORDER BY`, không nondeterministic); commit `639d938` (§1) | - | Trung bình |
| **E10** `nvl_or_btp` nhánh NVL với `_no_popup=True` và guard `_PLACEHOLDER_VALS` (`main_window.py:9353-9374`) | GIỮ | `main_window.py:9353-9374`; `EV-S10` — 63 lượt placeholder `'_'` trên 12/94 file, lọc đúng; commit `006fc61`, `449972a`, `113c443` (§1) | - | Thấp |
| **E11** `_log_bom_lookup_fail` (`main_window.py:7300-7317`), gọi tại `9375-9381` khi có Tên vật tư thật nhưng không resolve được `ItemId` — chỉ ghi file (an toàn gọi từ background thread) | GIỮ | `main_window.py:7300-7317, 9375-9381`; `EV-S10` (bối cảnh: dòng có `_vt_val` thật không thuộc `_PLACEHOLDER_VALS` mà vẫn `raw=None` là đúng lúc hàm này được gọi) | - | Thấp |
| **E12** `fuzzy_name` mặc định, Tier 1/2/3 + tie-break; Tier 3 tắt cho BOM4 (`main_window.py:9382-9400`) | GIỮ | `main_window.py:9382-9400`; `SQL-08` (7.192 nhóm trùng tên — population mà tie-break `1694b70` được viết ra để xử lý), `EV-S10`; commit `1694b70`, `8926a2b` (§1) | - | Trung bình |
| **E13** Fallback ĐVT đọc từ `B20Item` khi Excel ĐVT rỗng (`main_window.py:9406-9416`) | GIỮ | `main_window.py:9406-9416`; `EV-S00` (baseline — không ghi nhận lỗi liên quan Unit trong 94 file) | - | Thấp |
| **E14** Pass 2 SP fields (lookup/multi-output/single, `main_window.py:9424-9479`), `except Exception as e2` CÓ gọi `self._log(...)` (không nuốt lỗi) | GIỮ | `main_window.py:9424-9479`; `EV-M01` — BOM2-5 DETAIL có **0 record `nguon_dl='SP'`** hiện hành (bảng NguonDL theo section) → Pass 2 hiện là code chết cho BOM (không có field nào để lặp), nhưng an toàn/forward-compatible nếu mapping tương lai thêm field SP vào BOM DETAIL; khác `H`-group (SP_HOOK, event-driven theo `isactive`) — đây là cơ chế field-level `nguon_dl='SP'` riêng, chưa từng được BOM dùng | - | Thấp |

### 4.F BTP detection (BOM2)

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **F1** Scope chỉ BOM2: `section == 'BOM2' and stt_col is not None` (`main_window.py:9634`) | GIỮ | `main_window.py:9634`; `EV-S09` | - | Thấp |
| **F2** Tìm cột Tên vật tư/Tên chi tiết, gồm suffix-match (`main_window.py:9635-9657`) | GIỮ | `main_window.py:9635-9657`; `EV-S04` | - | Thấp |
| **F3** `_sec_id_seq` scope theo section qua `SECTION_STT_PATTERN` (`main_window.py:9662-9674`), commit `cb2c1d7` | GIỮ | `main_window.py:9662-9674`; `EV-S09` — 53 giá trị STT trùng giữa section trên 13/94 file (14%), xác nhận bug mà `cb2c1d7` vá là hiện tượng phổ biến, không phải edge case hiếm; commit `cb2c1d7` (§1, Chuỗi #5) | - | Thấp |
| **F4** `_is_real_vt` với `_PLACEHOLDER_VALS` (`main_window.py:9676-9680`), commit `113c443` | GIỮ | `main_window.py:9676-9680`; `EV-S10` | - | Thấp |
| **F5** `_find_parent_vt` — quét ngược tìm cha theo STT lồng thập phân, đệ quy nếu cha cũng rỗng (`main_window.py:9696-9704`) | GIỮ | `main_window.py:9696-9704`; `EV-S09` | - | Thấp |
| **F6** `_find_parent_vt` — fallback quét ngược trong cùng section khi không có dòng "X" trần (đánh số phẳng kiểu AS27, `main_window.py:9705-9715`), commit `3fd7985` | GIỮ | `main_window.py:9705-9715`; `EV-S09` — cơ chế heuristic "dòng gần nhất có Tên vật tư thật trong cùng section", chưa đo được trường hợp false-positive (2 nhóm phẳng liền kề bị lẫn cha) trong 94 file mẫu; commit `3fd7985` (§1) | - | Trung bình |
| **F7** Rule 2: STT nguyên + `"+"` trong Tên vật tư → container BTP thật (`main_window.py:9727-9730`) | SỬA | `main_window.py:9727-9730`; `EV-S09` — 12 dòng trên 6/94 file; commit `ebc9afe`, `63f6e6a` (§1, Chuỗi #3) | **[Q-08, chốt 2026-09-25 — chấp nhận #1]** Không tiếp tục đoán qua dấu "+". Đọc trực tiếp cột tường minh mới "Loại dòng: NVL / BTP" — dòng có giá trị "BTP" thì xử lý như container BTP, không cần suy luận từ nội dung Tên vật tư nữa. `EV-S09`'s 12 dòng hiện tại là baseline đối chiếu khi có file mẫu dùng cột mới. Requirement: MỚI | Trung bình |
| **F8** Rule 1a: Tên vật tư rỗng + có Tên chi tiết thật → BTP, áp dụng mọi tầng STT (`main_window.py:9732-9748`) | SỬA | `main_window.py:9732-9748`; `EV-S09` — 744 dòng trên 91/94 file (rule phổ biến nhất trong nhóm F); commit `fb137cc` (§1, Chuỗi #4 — mở rộng phạm vi từ chỉ-STT-nguyên sang mọi tầng) | **[Q-08, chốt 2026-09-25 — chấp nhận #1]** Cùng hướng với `F7`: đọc cột "Loại dòng" tường minh thay vì đoán từ (Tên vật tư rỗng + có Tên chi tiết). `EV-S09`'s 744 dòng/91 file là population lớn nhất bị ảnh hưởng — ưu tiên cao khi thiết kế cột mới ở Phase 2 vì đây là rule phổ biến nhất. Requirement: MỚI | Trung bình |
| **F9** Rule 1b: Tên vật tư rỗng + Tên chi tiết cũng rỗng nhưng có dòng con thập phân → BTP (cùng khối code với `F8`, `main_window.py:9737-9748`, điều kiện `_has_child`) | SỬA | `main_window.py:9737-9748`; `EV-S09` — **0/94 file**: không có trường hợp nào rơi riêng vào nhánh này (luôn bị `F8` bao trùm trước) — logic OR đúng, chỉ chưa được dữ liệu thật xác nhận độc lập | **[Q-08, chốt 2026-09-25 — chấp nhận #1]** Cùng hướng `F7`/`F8` — thay bằng đọc cột "Loại dòng" tường minh. 0/94 file hiện dùng riêng nhánh này nên rủi ro portover thấp. Requirement: MỚI | Thấp |
| **F10** Trường hợp loại trừ: Tên vật tư rỗng + Tên chi tiết rỗng + không con → KHÔNG phải BTP (comment đặc tả `main_window.py:9610-9612`; nhánh false ngầm định khi `not _ct_empty or _has_child` sai, không thêm vào `_btp_stt_set`) | SỬA | `main_window.py:9610-9612, 9734-9748`; `EV-S09` — 4 dòng trên 1/94 file; các dòng này rơi về `nvl_or_btp` NVL path (`E10`) với `_vt_val` rỗng → `raw=None`, sau đó có thể nhận `G1`/`G3` MKT fallback nếu không bị `_btp_skip_mkt_set` loại | **[Q-08, chốt 2026-09-25 — chấp nhận #1]** Với cột "Loại dòng" tường minh, "loại trừ" trở thành đơn giản: dòng không đánh dấu "BTP" (và không có Tên vật tư) vẫn rơi về NVL path (`E10`) như hiện tại — không cần suy luận phủ định (không tên + không con) nữa, chỉ cần đọc cột trực tiếp. Requirement: MỚI | Trung bình |
| **F11** Cut-piece: dòng con thập phân mà dòng cha (1 cấp trên) ĐÃ có Tên vật tư thật → KHÔNG bị đưa vào `_btp_skip_mkt_set` (vẫn thử tái dùng mã cũ trước, sau đó rơi về MKT fallback thay vì để SP_HOOK tạo mã mới) (`main_window.py:9750-9764`) | SỬA | `main_window.py:9750-9764`; `EV-S09` — 80 dòng trên 18/94 file (~19%); commit `458fdfd`, `3fd7985` (§1, Chuỗi #1, đã vá 2 lần) — dòng cut-piece đúng loại này rơi vào vùng rủi ro `Cao` của `G1`/`G3` (MKT fallback), nên tính đúng của `F11` là điều kiện tiên quyết để giới hạn đúng phạm vi rủi ro đó | **[Q-08, chốt 2026-09-25 — chấp nhận #1]** Đây là rule PHỨC TẠP NHẤT trong nhóm F (đệ quy quét ngược tìm cha qua `F5`/`F6`) — mức độ cột "Loại dòng" mới có thay thế được hoàn toàn `F11`'s quyết định "có nên fallback MKT hay không" hay chỉ thay được phần "có phải BTP không" cần Phase 2 thiết kế kỹ, có thể cần thêm 1 giá trị riêng cho "mảnh cắt" trong cột mới (không chỉ NVL/BTP nhị phân) thay vì suy luận qua `_find_parent_vt`. Đã vá phản ứng 2 lần (`458fdfd`, `3fd7985`) — dấu hiệu rõ nhất cho thấy cách tiếp cận đoán hiện tại không bền, cần chuyển hẳn sang tường minh. Requirement: MỚI | Trung bình |
| **F12** `_btp_key = (_sec_key, stt_v)` truyền `is_btp_row` vào `_resolve_detail_row` (`main_window.py:9808-9821`) | GIỮ | `main_window.py:9808-9821`; `EV-S09` (glue nhất quán với `F3`'s section-scoped key) | - | Thấp |

### 4.G MKT fallback

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **G1** Build `_mkt_cache` từ `vB20Item_MKT`: `SELECT ItemTypeSX_Parent, Id` không `ORDER BY` → dict, dòng sau ghi đè dòng trước | SỬA | `main_window.py:9529-9538` (build cache), `main_window.py:9835-9839` (áp dụng lookup `_mkt_cache.get(_item_type) or _mkt_cache.get(None)`); `SQL-01` — 4 nhóm trùng `ItemTypeSX_Parent` (C=3, F=2, I=2, NULL=2) trên 15 dòng của view (9 khóa distinct); `SQL-01(d)` cho thấy "người thắng" thực tế đổi giữa các lần chạy (`MKT_VAI`=27 dòng vs `MKT_DA`=227 dòng cùng nhóm C; `MKT_COKHI`=9 vs `MKT_KINH`=63 cùng nhóm NULL) và `MKT_SIMILI` không bao giờ thắng dù luôn có trong view; commit `e92c060` (tạo `_mkt_cache` lần đầu, 2026-08-11, chưa từng được sửa lại tới v2.2.25) | **[Cập nhật theo Q-01, business owner chốt 2026-09-25]** Business owner sẽ tự dọn dữ liệu gốc: tạo mã MKT gộp mới (vd `MKT_VAIDASIMILI`) thay thế `MKT_SIMILI`/`MKT_VAI`/`MKT_DA` cho nhóm `C` TRƯỚC (nhóm `F`/`I`/`NULL` xử lý sau, chưa có lịch). Code Phase 2 vẫn PHẢI thêm `ORDER BY` xác định (vd `CreatedAt DESC`) khi build `_mkt_cache` làm phòng vệ (defense-in-depth) cho các nhóm CHƯA được gộp — không được coi việc dọn dữ liệu của nhóm `C` là đủ để bỏ qua `ORDER BY` ở code, vì `F`/`I`/`NULL` vẫn còn collision thật cho tới khi business dọn tiếp. Theo D-04 (§1): commit `e92c060` tạo `_mkt_cache` lần đầu KHÔNG có `ORDER BY` và chưa từng được sửa lại tới v2.2.25 — đây không phải đảo ngược 1 fix trước đó, mà là đóng 1 khoảng trống có từ lúc tạo | Cao |
| **G2** Build `_mkt_cache` bọc trong `except Exception: pass` — thất bại (view mất quyền truy cập, mất kết nối tạm thời) bị nuốt HOÀN TOÀN, không log, cache rơi về `{}` rỗng (`main_window.py:9536-9538`) | SỬA | `main_window.py:9529-9538`; `SQL-01` (xác nhận `vB20Item_MKT` là view thật, truy cập bình thường trong điều kiện chuẩn — nên 1 lần thất bại là bất thường đáng ghi log, không phải trạng thái mong đợi); `SQL-07` (cache này phục vụ 84% lưu lượng MKT-fallback thật — 2.351/2.802 dòng) | Thêm log (`self._log(...)`) trong khối `except` trước khi để cache rỗng — 1 lần build cache thất bại (mất kết nối/quyền) hiện tại khiến MỌI dòng lẽ ra nhận MKT-fallback trong lần import đó âm thầm không nhận ItemId, không có dấu vết chẩn đoán, trên đúng code path đã xác nhận chiếm 84% lưu lượng MKT-fallback (`SQL-07`). Requirement: OBS-01 | Cao |
| **G3** Điều kiện áp dụng MKT fallback: `not ItemId and _mkt_cache and _btp_key not in _btp_skip_mkt_set` (`main_window.py:9835`) | GIỮ | `main_window.py:9835`; `SQL-07` — 2.802 dòng lịch sử dùng MKT fallback, điều kiện gate đúng theo thiết kế (chỉ áp dụng khi chưa có `ItemId` và không thuộc tập bị loại trừ bởi `F11`) | - | Cao |
| **G4** Key lookup `_mkt_cache.get(_item_type) or _mkt_cache.get(None)` — catch-all khi `ItemType` không khớp key nào (`main_window.py:9836-9839`) | GIỮ | `main_window.py:9836-9839`; `SQL-01`(a) — nhóm `NULL` có 2 dòng trùng (`MKT_COKHI` Id=421889, `MKT_KINH` Id=424817); `SQL-07` — nhóm `NULL` chiếm 72 dòng lịch sử. Cơ chế fallback-về-bucket-chung tự nó hợp lý (thiết kế đúng), nhưng bucket `None` KẾ THỪA đúng lỗi collision của `G1` — sửa ở `G1`, không cần sửa riêng ở đây | - | Cao |
| **G5** MKT Unit fallback: bỏ qua cho BOM4, chấp nhận giá trị `'-'` (khác `E13` — fallback Unit thông thường loại trừ `'-'`) (`main_window.py:9847-9856`) | GIỮ | `main_window.py:9847-9856`; `SQL-07` (bối cảnh: nhánh này chỉ chạy cho dòng đã nhận `ItemId` từ MKT fallback, tập con của 2.802 dòng lịch sử); commit `8a07639` (§1 — đã vá `.strip()` crash tại đúng dòng 9847 bằng `str(...) or ''` trước khi gọi `.strip()`, KHÔNG liên quan tới bug `G1`, xem ghi chú dưới bảng commit ở §1); `except Exception: pass` ở 9855-9856 vẫn không log — sẽ được liệt kê trong sweep `K2` ở §8, không lặp lại riêng ở đây | - | Thấp |

### 4.H SP_HOOK

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **H1** `mapping_loader.py:133-140`: `SP_CONFIG.Fallback` chuẩn hóa `'1.0'→'1'`; `mapping_loader.py:147-150`: `SP_HOOK` chỉ giữ dòng `isactive=='1'` NGAY LÚC LOAD mapping (không phải lúc dùng) | GIỮ | `mapping_loader.py:133-140, 147-150`; `EV-M01` — 3 dòng `SP_HOOK` active sau filter, 1 dòng `SP_CONFIG` (`Code`) có `isactive='0'` bị tắt đúng thiết kế | - | Thấp |
| **H2** `_run_row_sp_hooks` Condition `EMPTY()`/`NOTEMPTY()` (`bom_parser.py:422-427`) | GIỮ | `bom_parser.py:422-427`; `EV-M01` (Condition `EMPTY(ItemId)` gặp thật ở cả 3 dòng SP_HOOK active) | - | Thấp |
| **H3** Thay tham số `{field}`/`None` (`bom_parser.py:434-448`) | GIỮ | `bom_parser.py:434-448`; `EV-M01` | - | Thấp |
| **H4** Gọi procedure bằng f-string dựng từ tên SP trong mapping (`bom_parser.py:450`, `cur.execute(f'EXEC {h_sp} ' + ...)`) | GIỮ | `bom_parser.py:450`; `EV-M01` — tên SP (`h_sp`) đến từ `SP_HOOK.SP_Name` trong file mapping cục bộ (`CK_Mapping_v5.xlsx`), do đội dev/admin quản lý trực tiếp, KHÔNG phải input từ người dùng cuối UI (khác lớp `f"SELECT TOP 0 [{col_name}] FROM {table}"` mà `CONCERNS.md` đã flag — nơi input gần biên người dùng hơn). Ranh giới tin cậy đúng: file mapping = cấu hình admin, không phải input động | - | Thấp |
| **H5** Phân phối output field của SP trở lại `row_vals` (`bom_parser.py:451-456`) | GIỮ | `bom_parser.py:451-456`; `EV-M01` (`OutputFields="ItemId,ItemName,Unit"` cho cả 3 dòng SP_HOOK active) | - | Thấp |
| **H6** Exception trong hook → gọi `log_fn` CHỈ KHI được truyền vào (`bom_parser.py:457-459`) — mặc định tham số `log_fn=None`; grep toàn repo xác nhận 2 call site: BOM (`main_window.py:9859-9862`) LUÔN truyền `log_fn` thật (`self._log(...)`); THDM (`main_window.py:7127-7128`) truyền tường minh `log_fn=None` (im lặng hoàn toàn — ngoài phạm vi BOM, không sửa ở đây) | GIỮ | `bom_parser.py:410, 457-459`, `main_window.py:9859-9862, 7127-7128`; `EV-M01` — với BOM, hook exception luôn được log qua `self._log`, không bao giờ rơi vào nhánh im lặng | - | Thấp |
| **H7** Per-row `BeforeInsert` hook lọc theo section (`main_window.py:9584-9588`, gọi tại `9858-9862`) | GIỮ | `main_window.py:9584-9588, 9858-9862`; `EV-M01` — danh sách `mapping.get('SP_HOOK', [])` ĐÃ được lọc `isactive=='1'` từ lúc `load_mapping()` (`H1`), nên filter per-row này KHÔNG cần re-check `isactive` — so với batch filter (`H8`) có re-check tường minh `isactive=='1'` lần 2 (`main_window.py:9878`): re-check thứ 2 chỉ là double-guard vô hại (defensive), không phải bằng chứng 2 nơi xử lý khác nhau | - | Thấp |
| **H8** `BeforeInsertBatch` gom theo `sp_name` → `_run_batch_hook_grouped` (`main_window.py:9872-9891`, hàm tại `9965-10070+`): `usp_B20BOM_Create_ItemCode`, Condition `EMPTY(ItemId)`, case blank-`ItemName` cho XML khi `'+'` hoặc placeholder (commit `e414085`) | GIỮ | `main_window.py:9872-9891, 9965-10052`; `SQL-03` (xác nhận `usp_B20BOM_Create_ItemCode` là SYNONYM trỏ `B10_Boho.dbo...`, tham số khớp đúng những gì `Params` mapping truyền — nhưng `OBJECT_DEFINITION` trả `NULL`, không xem được nội dung SP thật); `EV-M01` (3 dòng active, section BOM2/3/4, không có hook cho BOM5); commit `e414085` (§1, Chuỗi #3) | - | Trung bình |

### 4.I Fill-Forward lúc insert

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **I1** `ff_fields` lấy từ mapping `fill_forward=='1'` + `ten_excel` (`main_window.py:9593-9597`) | GIỮ | `main_window.py:9593-9597`; `EV-M01` — `Fill_Forward=1` chỉ đúng field `ItemType` ở cả 4 section BOM2-5, 0 ở HEADER | - | Thấp |
| **I2** Nhận diện section-header lúc insert, với override GIAO RỜI theo kích thước vật lý (`main_window.py:9786-9799`) — **tái tính toán ĐỘC LẬP** cùng 1 quy tắc (`SECTION_STT_PATTERN` + dims-check) mà `_parse_sheet` (`C3`/`C4`/`C5`) đã tính 1 lần ở tầng parse | SỬA | `main_window.py:9786-9799` so với `bom_parser.py:526-544` (`_parse_sheet`, cùng quy tắc, code hoàn toàn tách biệt); `EV-S07` — kết quả 2 nơi hiện khớp nhau (1.083 dòng section-header, 16 dòng GIAO RỜI) nhưng không có gì đảm bảo 2 bản sao không lệch nhau sau 1 lần sửa chỉ ở 1 nơi — đúng lớp rủi ro mà `cb2c1d7` (§1) đã từng phải vá khi `_sec_id_seq` lệch giữa các bước | Gộp `SECTION_STT_PATTERN`-match + dims-check thành 1 hàm dùng chung (`is_section_header_row(stt, row, dim_cols) -> bool`), gọi từ cả `_parse_sheet` và `_generate_bom_details` thay vì 2 bản độc lập. Hiệu ứng quan sát được: không đổi kết quả hiện tại (vẫn khớp `EV-S07`), nhưng loại bỏ khả năng 2 nơi lệch nhau sau lần sửa kế tiếp. Requirement: ARCH-01 | Trung bình |
| **I3** Capture Fill-Forward tại dòng section-header (`main_window.py:9800-9806`) | GIỮ | `main_window.py:9800-9806`; `EV-S07`, `EV-S13` | - | Thấp |
| **I4** Ghi đè Fill-Forward **VÔ ĐIỀU KIỆN** khi insert: `row_vals[sql_col_ff] = current_ff[sql_col_ff]` (`main_window.py:9823-9826`), không kiểm tra `row_vals` đã có giá trị hay chưa | GIỮ | `main_window.py:9823-9826`; `EV-S13` — so với `C6` (nhánh `ff_always_override or not rd.get(sql_col)` của `_parse_section_excel_rows`, THDM-only, không chạy cho BOM ở tầng parse): BOM hoàn toàn KHÔNG có bước Fill-Forward nào ở tầng parse (`C6`), toàn bộ Fill-Forward của BOM diễn ra Ở TẦNG INSERT tại đúng dòng `I4` này, và ghi đè vô điều kiện 100% thời gian — khác hẳn ngữ nghĩa "chỉ ghi đè khi ô đang trống, trừ khi `ff_always_override`" mà `C6`/THDM mô tả. `EV-S13` xác nhận: field Fill-Forward duy nhất của BOM là `ItemType` (`Ten_Excel="STT"`), và MỌI dòng dữ liệu đều nằm dưới ít nhất 1 section-header (`EV-S07`: 94/94 file có ≥1 dòng section-header/section) → giá trị `ItemType` tự resolve từ STT riêng của chính dòng đó (qua Pass 1b, xem `EV-S11`) **luôn luôn** bị ghi đè bởi STT chữ cái của section-header gần nhất. Đây không phải "Fill-Forward khi thiếu giá trị" theo nghĩa thông thường — bản chất là "gán category theo nhóm section", vốn ĐÚNG THIẾT KẾ cho ý nghĩa của `ItemType`, không phải bug. Theo đúng hướng dẫn plan-wide: dữ liệu cho thấy giá trị riêng của dòng LUÔN bị ghi đè → câu hỏi này được đưa ra business owner xác nhận, xem `### Q-07` ở §5 | - | Trung bình |

### 4.J Trường header: EmployeeId / ParentDetailRowId_SO / đơn hàng

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **J1** `_on_creator_change` (`main_window.py:4779-4800`): cache miss `UserId` → query DB qua `_lookup_user_id_by_emp`; `_current_creator_employee_id` lưu RIÊNG, không qua fallback UserId (commit `9dc3c92`) | SỬA | `main_window.py:4779-4800`; `SQL-04` — sau `9dc3c92` (2026-09-21), `EmployeeId=1` vẫn là giá trị phổ biến nhất (48/82 dòng có EmployeeId) dù đã xuất hiện dải Id nhân viên thật đa dạng hơn; commit `9dc3c92` (§1, Chuỗi #6) | **[Cập nhật theo Q-02, business owner chốt 2026-09-25]** `EmployeeId=1` là tài khoản admin/hệ thống mặc định BAN ĐẦU cho user chưa có account — nay chính sách đã đổi, MỌI user phải có account/EmployeeId thật. Danh sách chọn nhân viên (dropdown nguồn của `_on_creator_change`) phải lọc CHỈ hiển thị user đã được khai báo `EmployeeId` — loại bỏ khả năng chọn/rơi về trạng thái không-EmployeeId dẫn tới `Id=1`. Requirement: MỚI | Trung bình |
| **J2** `_resolve_header_field` UILookup `creator`/`creator_employee` (`main_window.py:8941-8944`); nhánh `product_id`/`order_id`/`period_id` là THDM, chỉ ghi nhận, không đụng | SỬA | `main_window.py:8941-8944`; `SQL-04` — cùng nghi vấn `EmployeeId=1` như `J1` | Cùng hướng `J1` (Q-02) — field này CHỈ đọc `ui_values['creator']`/`['creator_employee']` đã được UI set trước đó, nên khi `J1` lọc đúng danh sách nhân viên, `J2` tự động nhận giá trị đúng theo, không cần sửa logic riêng ở đây. Requirement: MỚI (gộp cùng J1) | Trung bình |
| **J3** `_resolve_header_field` `CoDinh`/`HeThong` cho HEADER (`main_window.py:8918-8937`); `HeThong` chỉ xử lý `mac=='NOW'`, còn lại luôn `None` — ĐƠN GIẢN HƠN `D5a-g` (không có `AUTO_INC`/`BOMId`/`parent_fields`/copy-theo-tên) | GIỮ | `main_window.py:8918-8937`; `EV-M01` — cả 4 record `HeThong` của HEADER đều dùng macro `NOW`, không có record nào cần `AUTO_INC`/copy-theo-parent (HEADER là tài liệu gốc, không có "parent" hay `builtin_order` để copy) → nhánh đơn giản này ĐỦ và ĐÚNG cho dữ liệu mapping hiện hành, không phải thiếu sót | - | Thấp |
| **J4** `ParentDetailRowId_SO` phải theo công thức `"Mục số\|@ParentBizDocId"` giống `DetailRowId_SO`, không được bằng thẳng mã đơn hàng (`main_window.py:10434-10436, 10599-10602`, commits `432fa70`, `78b36fc`) | SỬA | `main_window.py:10434-10436, 10599-10602`; `SQL-05` — tỷ lệ dòng sai (bằng mã đơn hàng thay vì công thức) giảm từ 7,6% (644 dòng, trước vá) xuống 5,6% (18 dòng, sau vá `78b36fc`) nhưng **KHÔNG về 0%**; commit `78b36fc` (§1, Chuỗi #7) đã vá 1 điểm nhưng bằng chứng cho thấy còn nhánh khác chưa được phủ | Rà lại TOÀN BỘ nơi `ParentDetailRowId_SO`/`DetailRowId_SO` được gán (không chỉ 2 điểm đã đọc ở `10434-10436`/`10599-10602`) bằng kỹ thuật enum call-site giống đã áp dụng cho `D2`/`D3`/`B5` trong tài liệu này — tìm nhánh còn gán thẳng `ParentBizDocId` thay vì qua công thức. Gốc rễ CHÍNH XÁC là gì chưa xác nhận được bằng bằng chứng hiện có — xem `Q-03` ở §5 (đã ghi từ Plan 01). Theo D-04 (§1): commit `78b36fc` đã vá đúng 1 nhánh gán sai, nhưng `SQL-05` xác nhận residual 5,6% dòng sai VẪN xuất hiện sau vá — đề xuất này khác `78b36fc` ở chỗ mở rộng phạm vi audit ra TOÀN BỘ call-site thay vì tin rằng 1 điểm đã vá là đủ. Requirement: MỚI | Cao |
| **J5** `ParentBizDocId`/`BizDocId_SO` lấy thẳng từ đơn hàng đã chọn trên dropdown `_bom_selected_order_id` (`main_window.py:10437-10440, 10596-10598`, commits `f977b79`, `a9a891a`) | GIỮ | `main_window.py:10437-10440, 10596-10598`; `SQL-05` — 17/18 (sau vá) và 529/644 (trước vá) dòng có `ParentDetailRowId_SO == DetailRowId_SO` đúng — các trường hợp sai đo ở `SQL-05` đều là lỗi của `J4` (field khác), không phải bản thân phép gán trực tiếp này | - | Thấp |
| **J6** `BOMDetailType` suy từ `_CONFIG.view_insert` + mapping `CoDinh` (`main_window.py:9495-9507`, `SECTION_TO_TYPE`) | GIỮ | `main_window.py:9495-9507`; `EV-M01` (cấu trúc `_CONFIG`/mapping xác nhận khớp code) | - | Thấp |

### 4.K Xuyên suốt: sanitize, exception, insert

| Logic/rule con | Verdict (GIỮ/SỬA/BỎ) | Bằng chứng | Cách sửa (nếu SỬA) | Mức rủi ro |
|---|---|---|---|---|
| **K1** `.strip()` không được bảo vệ trên giá trị thô từ Excel/DB — 69 vị trí trong phạm vi luồng BOM (xem `### 8.1`), commit `8a07639` đã vá phản ứng đúng 1 vị trí (`main_window.py:9847`, xem `G5`) sau khi gây crash Import thật | SỬA | `main_window.py:9847` (vị trí đã vá phản ứng), đầy đủ 69 vị trí còn lại tại `### 8.1` (file:dòng:hàm); `EV-S11` — giá trị non-str (int/float) thật sự chảy vào field kiểu text ở ≥5.000 lượt/42 file, xác nhận rủi ro không phải lý thuyết; commit `8a07639` (§1) là bằng chứng crash thật đã xảy ra đúng lớp lỗi này | Thay từng vị trí ở `### 8.1` bằng `sanitizer.safe_str()` (`core/sanitizer.py`, ARCH-04) — hàm dùng chung ép `str()` trước khi `.strip()`, thay vì để nguyên receiver không chắc chắn kiểu. Hiệu ứng quan sát được: lớp crash `8a07639` đã vá 1 điểm phản ứng, `K1` đóng nốt các điểm còn lại chủ động thay vì chờ crash tiếp theo mới vá. Requirement: FIX-03, ARCH-04 | Cao |
| **K2** `except` không log và không re-raise — 43 khối trong phạm vi luồng BOM (xem `### 8.2`) | SỬA | `main_window.py:9855` (ví dụ đại diện, `_generate_bom_details`), đầy đủ 43 khối tại `### 8.2` (file:dòng:hàm:loại exception); `EV-S11` (nguồn dữ liệu thật — giá trị non-str/parse-fail — mà nhiều khối `except` này bảo vệ, ví dụ `_resolve_detail_row:9270,9282,9287,9292` — bị nuốt không log); trùng khớp trực tiếp với root-cause đã ghi nhận ở `PROJECT.md`/`.planning/codebase/CONCERNS.md` ("416 exception handler không log" ở `main_window.py`, "0 lần gọi logger" — nguyên nhân gốc chuỗi bug vá vụn v2.2.21→v2.2.25) | Với mỗi vị trí ở `### 8.2`: thêm `logger.exception(...)`/`self._log(...)` trước khi tiếp tục, HOẶC nếu im lặng thật sự an toàn (ví dụ probe optional-dependency) thì ghi rõ comment giải thích tại sao — theo đúng cách tiếp cận đã dùng cho `K1`. Requirement: OBS-01, OBS-02 | Cao |
| **K3** Insert từng dòng riêng lẻ vào `B20BOMDetail` (`main_window.py:9937-9957`, không batch/bulk); header + tất cả detail được insert trong CÙNG 1 transaction, `conn.commit()` 1 lần ở cuối, `conn.rollback()` khi lỗi (`_run_insert_bg`, `main_window.py:10839-10874`) | SỬA | `main_window.py:9937-9957, 10839-10874`; `SQL-06` — 8/662 header (1,2%) không có detail nào, NHƯNG cả 8 đều CŨ (2026-03-11→06-29) — **đính chính bằng chứng**: ghi chú gốc của `SQL-06` ở §3 suy đoán "không có transaction bao trùm cả 2 bước" dựa trên hiện tượng orphan-header; đọc lại code trực tiếp cho thấy `_run_insert_bg` HIỆN TẠI đã bọc cả insert header + `_generate_bom_details` + `commit()` trong 1 transaction — 8 dòng orphan là lịch sử, không phải bằng chứng thiếu transaction ở code hiện hành | Phần transaction header+detail (đã đúng, giữ nguyên khi port). Phần còn thiếu là ARCH-03 (Staging First): hiện tại insert thẳng vào `B20BOM`/`B20BOMDetail` thật, không qua bảng Staging → Validate trước. `core/bravo_staging.py` (Phase 2) nên coi transaction hiện có là bước "Transfer" cuối cùng của pattern, thêm bước Staging+Validate TRƯỚC nó. `SQL-06`'s 8 orphan lịch sử là baseline để xác nhận không phát sinh orphan mới sau khi thêm Staging. Requirement: ARCH-03 | Trung bình |
| **K4** `_sv` — render giá trị Python → SQL literal cho `export_only` (`main_window.py:9906-9935`), escape `'` → `''`, phân biệt kiểu theo `col_kieu` | GIỮ | `main_window.py:9906-9935`; `EV-S11` (giá trị non-str phải chảy qua đúng nhánh phân kiểu ở đây khi xuất SQL text) — chỉ dùng cho đường xuất preview/export, đường insert thật dùng tham số hóa (`cur.execute(sql_exec, exec_vals)`, `K3`), không đi qua `_sv`, không có rủi ro SQL injection ở đường insert thật | - | Thấp |
| **K5** Điều kiện bỏ qua section: không thuộc `SECTION_TO_TYPE`, `df` rỗng, có cột `'Lỗi'`, không có `detail_recs` (`main_window.py:9558-9569`) | GIỮ | `main_window.py:9558-9569`; `EV-S01` — 14/94 file thiếu `BOM5` sinh `df` chứa cột `'Lỗi'` (xem `A5`), điều kiện này đúng chặn không cố insert dòng `'Lỗi'` như dữ liệu thật | - | Thấp |

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

**Trả lời:** Business owner (2026-09-25) đề xuất giải pháp KHÁC với 3 phương án đưa ra — không chọn ưu tiên-code, mà **dọn dữ liệu gốc**: tạo 1 mã MKT gộp mới (ví dụ `MKT_VAIDASIMILI`) ở nhóm cha, thay thế cho `MKT_SIMILI`/`MKT_VAI`/`MKT_DA` (nhóm `C`) — tương tự cách `MKT_CHI` đã làm cho 1 nhóm khác. **Phạm vi đợt này: chỉ nhóm `C` trước**; nhóm `F`, `I`, `NULL` xử lý sau, chưa có lịch cụ thể. Hệ quả cho Phase 2: (a) code vẫn nên thêm `ORDER BY` cho `G1` làm phòng vệ (defense-in-depth) cho các nhóm CHƯA được gộp (`F`, `I`, `NULL`) trong lúc chờ business dọn tiếp; (b) một khi nhóm `C` chỉ còn 1 mã, `G1`'s dict-collision cho nhóm này tự nhiên biến mất không cần logic ưu tiên; (c) việc tạo mã gộp + gán lại `ItemTypeSX_Parent` cho `vB20Item_MKT` là thao tác dữ liệu do business/admin tự làm trong Bravo, KHÔNG phải việc của Phase 2 code.

### Q-02

**Bối cảnh (`SQL-04`, EmployeeId):** sau khi `9dc3c92` (v2.2.22) vá lỗi EmployeeId hard-code=1, dữ liệu thật SAU vá vẫn cho `EmployeeId=1` là giá trị phổ biến nhất (48/82 dòng có EmployeeId, so với dải Id nhân viên thật khác chỉ 1-10 dòng mỗi Id). **Câu hỏi:** `EmployeeId=1` sau vá có phải là 1 nhân viên thật (ví dụ tài khoản admin/hệ thống) hay vẫn là 1 nhánh fallback cũ chưa được `9dc3c92` bao phủ hết? Cần business hoặc DBA xác nhận Id=1 trong bảng nhân viên tương ứng là ai.

**Trả lời:** Business owner (2026-09-25) xác nhận: `EmployeeId=1` là **tài khoản admin/hệ thống mặc định ban đầu**, dùng làm fallback cho user chưa có account khai báo. Tuy nhiên **chính sách hiện tại đã đổi — mọi user đều phải có account/EmployeeId thật**. Yêu cầu cho Phase 2: danh sách chọn nhân viên (employee lookup dropdown) phải **lọc chỉ hiển thị user đã được khai báo `EmployeeId`** — loại bỏ khả năng rơi về `Id=1` một cách ngầm định. Đây là 1 yêu cầu cụ thể mới, nên ghi nhận vào Phase 2 cùng nhóm `J1`/`J2` (không chỉ là câu hỏi làm rõ dữ liệu — có hành động sửa đi kèm).

### Q-03

**Bối cảnh (`SQL-05`, ParentDetailRowId_SO):** sau khi `78b36fc` (v2.2.23) vá công thức ParentDetailRowId_SO, tỷ lệ dòng sai theo mẫu CŨ (bằng mã đơn hàng thay vì công thức Mục số) giảm từ 7.6% (644 dòng, trước vá) xuống 5.6% (18 dòng, sau vá) — KHÔNG về 0%. Mẫu sau-vá còn nhỏ (18 dòng) nên chưa chắc chắn về thống kê. **Câu hỏi:** có luồng import nào khác (THDM, hoặc 1 loại sản phẩm/section cụ thể) không đi qua đúng nhánh code mà `78b36fc` đã vá? Cần business xác nhận nguồn gốc của (các) dòng sai còn lại sau vá, hoặc executor cần thêm thời gian điều tra nếu Phase 2 quyết định port field này.

**Trả lời:** Business owner (2026-09-25): chưa biết nguồn gốc cụ thể — **để Phase 2 điều tra thêm** theo đúng đề xuất ở `Cách sửa` của `J4` (rà toàn bộ call-site gán `ParentDetailRowId_SO`/`DetailRowId_SO`, không chỉ 2 điểm đã đọc ở Plan 01).

### Q-04

**Bối cảnh (`EV-S11`/`EV-S13`, cột `Code` nhận giá trị non-str từ STT):** trường `Code` (BOM5, `Ten_Excel="STT"`, kiểu khai `varchar`) nhận giá trị int/float trực tiếp từ Excel ở phần lớn các dòng (không có cơ chế Fill-Forward ghi đè như `ItemType`, xem `EV-S13`). Chưa xác định được `Code` có được insert trực tiếp (rủi ro `.strip()`-crash giống `8a07639`) hay chỉ dùng làm khóa lookup nội bộ (không insert, không gọi `.strip()`). **Câu hỏi cho Phase 2 (không phải business — kỹ thuật):** đọc tiếp luồng dùng `BOM5.Code` sau `_resolve_detail_row` trước khi chốt verdict cho field này.

**Trả lời:** (chờ Phase 2 — câu hỏi kỹ thuật, không cần business owner)

### Q-05

**Bối cảnh (`A5`, cảnh báo thiếu section):** `EV-S01` xác nhận `BOM5` thiếu ở 14/94 file mẫu thật (0/94 thiếu BOM2/3/4) — dữ liệu thật cho thấy BOM5 là section tùy chọn trong thực tế, nhưng code hiện tại cảnh báo `"Không tìm thấy sheet …"` giống hệt nhau cho MỌI section, không phân biệt bắt buộc/tùy chọn. **Câu hỏi:** ngoài BOM5, có section nào khác (hiện tại hoặc dự kiến thêm — BOM6, BOM7...) mà business coi là tùy chọn không? Trả lời câu này quyết định cột `Required`/danh sách tùy chọn cụ thể sẽ khai trong `_CONFIG` ở Phase 2 cho `A5`/`FIX-04`.

**Trả lời:** Business owner (2026-09-25): **chỉ BOM5 là tùy chọn** — BOM2/3/4 vẫn bắt buộc như hiện tại. Phase 2 chỉ cần đánh dấu `Required=0` cho đúng `BOM5` trong `_CONFIG` (`A5`/`FIX-04`), không cần thêm section nào khác vào danh sách tùy chọn.

### Q-06

**Bối cảnh (`A4`, xung đột `global_meta` "sheet đầu tiên thắng"):** `EV-S12` — 229 xung đột trên 93/94 file giữa các khóa meta (`Kích thước`, `Số lượng`, `Hoàn thiện`, `Vật liệu`) đọc được từ các sheet BOM2-5 khác nhau của cùng 1 file; cơ chế hiện tại luôn lấy giá trị của sheet BOM2 (sheet đầu tiên duyệt qua) làm giá trị cuối cùng, kể cả khi sheet khác (vd BOM5) có giá trị khác nghĩa thật (ví dụ `Số lượng`="1 Hệ" ở BOM2 nhưng ="9" ở BOM5). **Câu hỏi:** với mỗi khóa meta hiện có (`Kích thước`, `Số lượng`, `Hoàn thiện`, `Vật liệu`), section nào là NGUỒN ĐÚNG khi các sheet có giá trị khác nhau? (Ví dụ: `Số lượng` nên luôn lấy từ BOM2, hay từ section nào đang active/được chọn insert?) Trả lời câu này quyết định nội dung cụ thể của "Cách sửa" cho `A4` ở Phase 2.

**Trả lời:** (chờ business owner)

### Q-07

**Bối cảnh (`I4`, Fill-Forward `ItemType` ghi đè vô điều kiện):** `EV-S13` xác nhận field Fill-Forward duy nhất của BOM (`ItemType`) LUÔN bị ghi đè bởi STT chữ cái của section-header gần nhất khi insert, bất kể dòng dữ liệu đó tự có giá trị `ItemType` gì (resolve từ STT riêng của chính nó). 94/94 file có ít nhất 1 dòng section-header/section, nên override này thực thi 100% thời gian trong dữ liệu mẫu — nói cách khác, `ItemType` không hoạt động như "điền khi thiếu" theo nghĩa Fill-Forward thông thường mà như "gán category theo nhóm section" (STT chữ cái A/B/C... → category). **Câu hỏi:** đây có đúng là ý nghĩa nghiệp vụ mong muốn của `ItemType` không (mỗi dòng thuộc nhóm section nào quyết định `ItemType`, không phải giá trị STT số riêng của dòng)? Nếu đúng, `I4` port nguyên trạng vào `core/`. Nếu KHÔNG — nếu có kịch bản business muốn STT số riêng của dòng override lại nhóm section — cần đặc tả lại quy tắc trước khi Phase 2 viết `core/`.

**Trả lời:** Business owner (2026-09-25) xác nhận: **đúng ý nghĩa nghiệp vụ mong muốn** — `ItemType` = phân loại theo nhóm section cha (chữ cái A/B/C... gần nhất phía trên dòng đó), KHÔNG phải giá trị tự tính từ STT riêng của từng dòng. `I4` port nguyên trạng vào `core/`, không cần đặc tả lại.

### Ghi chú: các dòng SỬA/BỎ chỉ có bằng chứng EV-M (census mapping), theo quy tắc bằng chứng (D-02/D-03)

Các dòng dưới đây có Bằng chứng **chỉ** dựa trên `EV-M01` (census mapping production) + đọc/đối chiếu code, KHÔNG có dữ liệu mẫu (`EV-Snn`) hay SQL thật (`SQL-nn`) trực tiếp chứng minh — liệt kê tường minh ở đây theo đúng quy tắc "SỬA/BỎ chỉ có bằng chứng EV-M phải được liệt kê ở §5":

- **`D2`** (BỎ — nhánh `Excel` của `_resolve_row_mapping` không chạy cho BOM): bằng chứng là grep toàn repo xác nhận call-site (cross-reference code, không phải dữ liệu mẫu) + `EV-M01`.
- **`D3`** (BỎ — nhánh `Bien_doi` của `_resolve_row_mapping` không chạy cho BOM): cùng bằng chứng call-site như `D2`.
- **`D7`** (BỎ — `MucLookup` không có record BOM/HEADER nào): bằng chứng là `EV-M01` census (0 record `MucLookup`), không có mẫu Excel nào chứng minh trực tiếp field bị ảnh hưởng (vì không có field nào cả).
- **`E5`** (SỬA — `EMPTY` macro không phân biệt kiểu ngày ở Pass 1b): bằng chứng là so sánh code trực tiếp giữa `_resolve_detail_row` và `_resolve_row_mapping` + `EV-M01` (tần suất macro `EMPTY` gặp thật); KHÔNG có sample/SQL nào xác nhận đã có field ngày BOM thực sự bị lỗi này — rủi ro tiềm ẩn (latent), chưa quan sát được crash thật trong 94 file mẫu hoặc SQL thật.

### Q-08 — Chiến lược "ép chuẩn Excel, không đoán" cho các logic đoán-vì-file-lộn-xộn

**Bối cảnh:** Rà soát lại §4 sau khi bảng verdict đã ra, phát hiện 1 nhóm logic có đặc điểm chung: tồn tại KHÔNG PHẢI vì nghiệp vụ cần nó, mà vì code đang **đoán ý nghĩa dữ liệu từ 1 file Excel không có chuẩn tường minh**, thay vì bắt buộc user sửa file/template cho đúng chuẩn (đúng tinh thần "Ép chuẩn User — Chặn cứng Data rác" đã thống nhất từ đầu dự án). Đây là ghi chú thảo luận, **CHƯA đổi verdict** trong §4 — business owner quyết định ở Plan 03 sau khi bàn kỹ từng điểm:

1. **Nhóm F — BTP detection (F7-F11, `main_window.py:9717-9765`)**: đoán 1 dòng trống-Tên-vật-tư có phải bán-thành-phẩm không, dựa trên có Tên chi tiết/dòng con/dấu "+"/dòng cha có vật tư thật hay không (đệ quy quét ngược). Đề xuất thảo luận: **BỎ toàn bộ heuristic**, thay bằng 1 cột tường minh trong template Excel (ví dụ "Loại dòng: NVL / BTP") — reject dòng thiếu cột này thay vì đoán.
2. **B2/B5/E4 — suffix-match 2 lần khớp header bị merge** (71/94 file phải dựa vào đoán hậu tố, ví dụ `SLg_Tên_Vật_Tư` → "Tên vật tư"): đề xuất thay đoán-theo-hậu-tố bằng 1 danh sách alias tường minh khai trong mapping, reject tên cột không nằm trong danh sách.
3. **C2 — footer detection bằng từ khóa** (2.573 dòng bị cắt vĩnh viễn trên 94/94 file, cùng lớp lỗi `817bbbe` từng vá 1 lần rồi tái diễn): đề xuất thêm dấu hiệu kết-thúc-bảng tường minh trong template thay vì đoán theo từ khóa.
4. **A4 — `global_meta` "sheet đầu tiên thắng"** (229 xung đột/93 file, vd "Số lượng"="1 Hệ" ở BOM2 nhưng ="9" ở BOM5, giá trị khác nghĩa thật bị bỏ qua âm thầm): đề xuất cảnh báo bắt buộc khi sheet xung đột thay vì tự chọn sheet đầu.
5. **E10 — giá trị placeholder rác trong Tên vật tư/Tên chi tiết** (`_`, `--`, `-`, `x`, `n/a`; 63 lần trên 12/94 file): đề xuất bắt ô trống thật, cảnh báo nếu gặp ký tự lạ, không coi các ký tự này là "= rỗng" ngầm định.
6. **`#NAME?` lỗi công thức Excel lọt vào tên cột** (`EV-S03`, 2/94 file, ví dụ cột `SLSP_#NAME?`): đề xuất reject thẳng file có lỗi công thức ở vùng header thay vì cố đọc tiếp.

**Trả lời:** Business owner đã quyết định (2026-09-25):

- **Chấp nhận #1 và #3** — đưa vào yêu cầu chuẩn hóa template Excel:
  - #1: thêm cột tường minh "Loại dòng: NVL / BTP" trong BOM2, thay thế heuristic đoán ở nhóm F.
  - #3: thêm dấu hiệu kết-thúc-bảng tường minh, thay thế đoán-theo-từ-khóa ở `C2`.
- **BỎ #2, #4, #5, #6** — không bắt user thay đổi file thêm; code tiếp tục tự xử lý các trường hợp này như bảng §4 gốc đã đề xuất (suffix-match cho header, resolve-theo-khóa cho `global_meta`, coi placeholder là rỗng, tiếp tục đọc header dù có lỗi công thức) — KHÔNG đổi verdict/Cách sửa hiện có của `B2`/`B5`/`E4`/`A4`/`E10` trong §4.

Hệ quả: verdict + "Cách sửa" của nhóm `F` (F7-F11) và `C2` trong §4 được cập nhật lại theo hướng #1/#3 — xem nội dung đã sửa trực tiếp trong bảng. Thiết kế chi tiết cột "Loại dòng" mới tương tác thế nào với logic phụ trợ hiện có của BTP (`F3` section-scope, `F5`/`F6` quét-ngược-tìm-cha, `F11` cut-piece) — **để Phase 2 thiết kế cụ thể khi viết `core/`**, không tự quyết ở đây.

## 6. Tổng hợp & ánh xạ sang Phase 2

**Tổng số dòng §4:** 90 (đúng bằng số ID bắt buộc tối thiểu — không thêm dòng phụ ngoài danh sách bắt buộc của Plan 02).

**GIỮ: 64 · SỬA: 23 · BỎ: 3** (cập nhật 2026-09-25 — 2 đợt: (1) quyết định Q-08: `F7`-`F11` GIỮ→SỬA, thay heuristic đoán BTP bằng đọc cột "Loại dòng" tường minh; `C2` đổi "Cách sửa" sang dấu hiệu kết-thúc-bảng tường minh. (2) trả lời Q-01..Q-07: `J1`/`J2` GIỮ→SỬA — lọc danh sách nhân viên chỉ hiện user có `EmployeeId` khai báo (Q-02); `G1` giữ SỬA nhưng "Cách sửa" cập nhật theo hướng business tự dọn dữ liệu gốc cho nhóm `C` + code vẫn thêm `ORDER BY` phòng vệ cho nhóm còn lại (Q-01); `A5`/`I4` không đổi verdict, chỉ xác nhận phạm vi/đúng thiết kế (Q-05/Q-07). Số gốc ban đầu (trước mọi cập nhật): GIỮ 71 · SỬA 16 · BỎ 3)

### Bảng ánh xạ SỬA/BỎ → yêu cầu Phase 2

| Row ID | Verdict | Mức rủi ro | Yêu cầu Phase 2 | Ghi chú |
|---|---|---|---|---|
| `A4` | SỬA | Cao | MỚI | `global_meta` "sheet đầu tiên thắng" → resolve theo khóa; quy tắc cụ thể chờ `Q-06` |
| `A5` | SỬA | Trung bình | FIX-04 | Thêm cờ bắt buộc/tùy chọn cho section trong `_CONFIG`; danh sách tùy chọn cụ thể chờ `Q-05` |
| `A6` | SỬA | Trung bình | OBS-01 | Thêm log cho 2 khối `except` nuốt lỗi khi mở workbook lần 2/3 |
| `B2` | SỬA | Trung bình | FIX-02 | Provenance-flag cho cột forward-fill + drop cột orphan không khớp `Ten_Excel` nào |
| `B5` | SỬA | Trung bình | ARCH-01 | Trích cơ chế match 2-pass trong `_resolve_detail_row` thành hàm build-column-map dùng chung, chạy 1 lần/section |
| `C2` | SỬA | Cao | MỚI | Thêm dấu hiệu kết-thúc-bảng tường minh trong template (Q-08 #3, đã chốt); reject file thiếu dấu hiệu — không còn đoán qua `FOOTER_KEYWORDS` |
| `F7` | SỬA | Trung bình | MỚI | Đọc cột "Loại dòng: NVL/BTP" tường minh (Q-08 #1, đã chốt) thay vì đoán qua dấu "+" |
| `F8` | SỬA | Trung bình | MỚI | Đọc cột "Loại dòng" tường minh (Q-08 #1) — population lớn nhất bị ảnh hưởng (744 dòng/91 file) |
| `F9` | SỬA | Thấp | MỚI | Đọc cột "Loại dòng" tường minh (Q-08 #1) — 0/94 file dùng riêng nhánh này |
| `F10` | SỬA | Trung bình | MỚI | Đọc cột "Loại dòng" tường minh (Q-08 #1) — loại trừ đơn giản hóa, không cần suy luận phủ định |
| `F11` | SỬA | Trung bình | MỚI | Cần Phase 2 thiết kế kỹ (Q-08 #1) — có thể cần thêm giá trị riêng cho "mảnh cắt" trong cột mới, không chỉ nhị phân NVL/BTP; đã vá phản ứng 2 lần (`458fdfd`, `3fd7985`) |
| `C10` | SỬA | Trung bình | ARCH-04 | Gộp 3 bản `normalize_stt` độc lập thành 1 hàm `core/sanitizer.py` |
| `D2` | BỎ | Thấp | — | Nhánh `Excel` của `_resolve_row_mapping` không chạy cho BOM (chỉ THDM) — không port |
| `D3` | BỎ | Thấp | — | Nhánh `Bien_doi` của `_resolve_row_mapping` không chạy cho BOM (chỉ THDM) — không port |
| `D7` | BỎ | Thấp | — | `MucLookup` không có record BOM/HEADER nào — không port |
| `E5` | SỬA | Trung bình | MỚI | Đồng nhất xử lý macro `EMPTY` cho field ngày giữa Pass 1b và `_resolve_row_mapping` |
| `E6` | SỬA | Trung bình | OBS-01, OBS-02 | Log khi ép kiểu date/number/int thất bại, trước khi `raw=None` |
| `G1` | SỬA | Cao | FIX-01 | Thêm `ORDER BY` xác định khi build `_mkt_cache`; quy tắc ưu tiên cụ thể chờ `Q-01` |
| `G2` | SỬA | Cao | OBS-01 | Log khi build `_mkt_cache` thất bại (hiện `except Exception: pass` hoàn toàn im lặng) |
| `I2` | SỬA | Trung bình | ARCH-01 | Gộp rule nhận diện section-header (parse-time và insert-time) thành 1 hàm dùng chung |
| `J4` | SỬA | Cao | MỚI | Rà lại toàn bộ nơi gán `ParentDetailRowId_SO`/`DetailRowId_SO`; gốc rễ residual 5,6% lỗi chờ `Q-03` |
| `K1` | SỬA | Cao | FIX-03, ARCH-04 | Thay 69 vị trí `.strip()` không bảo vệ (§8.1) bằng `sanitizer.safe_str()` |
| `K2` | SỬA | Cao | OBS-01, OBS-02 | Thêm log cho 43 khối `except` không log (§8.2) |
| `K3` | SỬA | Trung bình | ARCH-03 | Thêm bước Staging+Validate trước transaction insert hiện có (đã đúng, giữ nguyên làm bước Transfer) |
| `J1` | SỬA | Trung bình | MỚI | Lọc dropdown chọn nhân viên chỉ hiện user đã khai báo `EmployeeId` (Q-02, đã chốt) |
| `J2` | SỬA | Trung bình | MỚI | Ăn theo `J1` — không cần sửa logic riêng, tự nhận giá trị đúng khi `J1` lọc đúng |

### FIX-01..FIX-04 — xác nhận có mặt

- **FIX-01** (`_mkt_cache` ORDER BY): dòng `G1`.
- **FIX-02** (`_build_headers` forward-fill vô hạn): dòng `B2`.
- **FIX-03** (`.strip()` → `sanitizer.safe_str()`): dòng `K1`.
- **FIX-04** (cảnh báo section tùy chọn): dòng `A5`.

Không có FIX nào bị bằng chứng bác bỏ — cả 4 đều được xác nhận là vấn đề thật bằng dữ liệu mẫu/SQL thật (`EV-S01`, `EV-S03`, `EV-S06`(gián tiếp qua cùng cơ chế với K1's crash evidence)/`EV-S11`, `SQL-01`).

### Danh sách các dòng Mức rủi ro = Cao

`A4`, `C2`, `G1`, `G2`, `G3`, `G4`, `J4`, `K1`, `K2` — 9 dòng. `G3`/`G4` là GIỮ (cơ chế áp dụng/lookup đúng thiết kế) nhưng thừa hưởng rủi ro Cao của `G1` vì cùng nằm trên code path MKT-fallback đã xác nhận chiếm 84% lưu lượng thật (`SQL-07`).

## 7. Xác nhận của business owner

_(Điền ở Plan 03 — ghi lại xác nhận bằng văn bản, kèm ngày, từ người dùng trong chat. Im lặng, auto-advance hay yolo mode không tính là xác nhận.)_

## 8. Phụ lục

Nguồn: `tests/tmp/ba_bom_flow_sweep.py` — script AST một-lần, read-only (không import/chạy code mục tiêu, không ghi vào bất kỳ file nào dưới `v3/Tools/`), quét đúng các hàm luồng BOM liệt kê ở Task 3 Step 1 của `01-02-PLAN.md` (loại trừ `_thdm_*`, helper THDM-only, và dialog helper). Kết quả dump ra `tests/tmp/ba_bom_flow_sweep.json`, đọc 1 lần để viết §8.1/§8.2 dưới đây; cả 2 file tạm đã bị xoá ở Task 3 Step 6 (xem xác nhận cuối tài liệu).

### 8.1 Vị trí .strip() rủi ro (K1)

Tổng **69 vị trí `.strip()` không được bảo vệ tường minh** (receiver là Name/Attribute/Subscript/Call khác `str()`, hoặc `(x or ...).strip()` không ép `str()` trước — CHÍNH LÀ lớp lỗi mà commit `8a07639` đã vá phản ứng tại 1 điểm `main_window.py:9847`, xem `G5`) và **41 vị trí đã tự bảo vệ** (`str(x).strip()` ép kiểu tường minh — không thể crash bất kể kiểu dữ liệu đầu vào).

**bom_parser.py** — unguarded: 16, guarded (đã ép `str()`): 22

| Dòng | Hàm | Snippet |
|---|---|---|
| 82 | `_load_sheet_config` | `contains  = [x.strip().upper() for x in str(row[3] or '').split(',') if x.strip(…` |
| 83 | `_load_sheet_config` | `excludes  = [x.strip().upper() for x in str(row[4] or '').split(',') if x.strip(…` |
| 422 | `_run_row_sp_hooks` | `cond       = hook.get('condition', '').strip()` |
| 431 | `_run_row_sp_hooks` | `h_sp   = hook.get('sp_name', '').strip()` |
| 432 | `_run_row_sp_hooks` | `h_par  = hook.get('params', '').strip()` |
| 433 | `_run_row_sp_hooks` | `h_outs = [f.strip() for f in hook.get('outputfields', '').split(',') if f.strip(…` |
| 436 | `_run_row_sp_hooks` | `ph = ph.strip()` |
| 441 | `_run_row_sp_hooks` | `v = v.strip()` |
| 425 | `_run_row_sp_hooks` | `should_run = not row_vals.get(cond[6:-1].strip())` |
| 427 | `_run_row_sp_hooks` | `should_run = bool(row_vals.get(cond[9:-1].strip()))` |
| 440 | `_run_row_sp_hooks` | `k = k.strip().lstrip('@')` |
| 1305 | `_resolve_row_mapping` | `or (isinstance(val, str) and val.strip() == ''))` |
| 1327 | `_resolve_row_mapping` | `val = val.strip()` |

**mapping_loader.py** — unguarded: 6, guarded (đã ép `str()`): 0

| Dòng | Hàm | Snippet |
|---|---|---|
| 78 | `_load_section_rows` | `df = df[df['Section'].astype(str).str.strip() == section]` |
| 234 | `build_meta_keys_from_mapping` | `ten = rec.get('ten_excel', '').strip()` |
| 243 | `build_meta_keys_from_mapping` | `part = part.strip()` |
| 252 | `build_meta_keys_from_mapping` | `part = part.strip()` |
| 269 | `build_cell_specs_from_mapping` | `ten = rec.get('ten_excel', '').strip()` |
| 272 | `build_cell_specs_from_mapping` | `coord = ten.split('\|')[0].strip()` |

**main_window.py** — unguarded: 47, guarded (đã ép `str()`): 19

| Dòng | Hàm | Snippet |
|---|---|---|
| 4800 | `_on_creator_change` | `self.cmb_creator.set(selected_name.split("\|")[0].strip())` |
| 8969 | `_resolve_header_field` | `for _p in [p.strip() for p in ten_excel.split('\|')]:` |
| 8984 | `_resolve_header_field` | `_te = _te.strip()` |
| 9020 | `_resolve_header_field` | `s = raw.strip()` |
| 9084 | `_build_bom_detail_caches` | `for _ss1 in [f.strip() for f in re.split(r'[\|,]', ss) if f.strip()]:` |
| 9445 | `_resolve_detail_row` | `sp_name2  = cfg2.get('sp_name', '').strip()` |
| 9446 | `_resolve_detail_row` | `params_s2 = cfg2.get('params', '').strip()` |
| 9196 | `_resolve_detail_row` | `_t = _t.strip()` |
| 9258 | `_resolve_detail_row` | `or (isinstance(raw, str) and raw.strip() == ''))` |
| 9314 | `_resolve_detail_row` | `_code_val, _name_val = _code_val.strip(), _name_val.strip()` |
| 9430 | `_resolve_detail_row` | `out_fs = [f.strip() for f in cfg2.get('outputfields', '').split(',') if f.strip(…` |
| 9447 | `_resolve_detail_row` | `fallback2 = cfg2.get('fallback', '').strip() or None` |
| 9448 | `_resolve_detail_row` | `out_fs2   = [f.strip() for f in cfg2.get('outputfields', '').split(',') if f.str…` |
| 9298 | `_resolve_detail_row` | `_bien_doi = _nan_str(rec.get('bien_doi', '')).strip().upper()` |
| 9346 | `_resolve_detail_row` | `_vt_val, _ct_val = _vt_val.strip(), _ct_val.strip()` |
| 9406 | `_resolve_detail_row` | `if sql_col == 'Unit' and (raw is None or (isinstance(raw, str) and not raw.strip…` |
| 9304 | `_resolve_detail_row` | `raw = raw.strip()` |
| 9387 | `_resolve_detail_row` | `and raw.strip() in {'_', '--', '-', 'x', 'n/a'}` |
| 9456 | `_resolve_detail_row` | `ph = ph.strip()` |
| 9459 | `_resolve_detail_row` | `kh = kh.strip().lstrip('@'); vh = vh.strip()` |
| 9988 | `_run_batch_hook_grouped` | `cond       = hook.get('condition', '').strip()` |
| 9989 | `_run_batch_hook_grouped` | `xml_param  = hook.get('xmlparam', '').strip()` |
| 9990 | `_run_batch_hook_grouped` | `xml_tag    = hook.get('xmltag', '').strip()` |
| 9991 | `_run_batch_hook_grouped` | `xml_fields = [f.strip() for f in hook.get('xmlfields', '').split(',') if f.strip…` |
| 9992 | `_run_batch_hook_grouped` | `h_outs     = [f.strip() for f in hook.get('outputfields', '').split(',') if f.st…` |
| 10006 | `_run_batch_hook_grouped` | `cond_field = cond[6:-1].strip()` |
| 10059 | `_run_batch_hook_grouped` | `ph = ph.strip()` |
| 10063 | `_run_batch_hook_grouped` | `k = k.strip().lstrip('@'); v = v.strip()` |
| 10009 | `_run_batch_hook_grouped` | `cond_field = cond[9:-1].strip()` |
| 10047 | `_run_batch_hook_grouped` | `('+' in _fval or _fval.strip() in {'_', '--', '-', 'x', 'n/a'}):` |
| 10633 | `_header_resolve_bg` | `sp_name  = sp_cfg.get('sp_name', '').strip()` |
| 10634 | `_header_resolve_bg` | `params_s = sp_cfg.get('params', '').strip()` |
| 10679 | `_header_resolve_bg` | `_vsql   = _v.get('sql', '').strip()` |
| 10680 | `_header_resolve_bg` | `_vparam = _v.get('params', '').strip()` |
| 10681 | `_header_resolve_bg` | `_warn_m = _v.get('warningmessage', '').strip()` |
| 10635 | `_header_resolve_bg` | `fallback = sp_cfg.get('fallback', '').strip() or None` |
| 10685 | `_header_resolve_bg` | `_params = [row.get(p.strip()) for p in _vparam.split(',') if p.strip()]` |

### 8.2 except không log (K2)

Tổng **43 khối `except` không gọi logger và không re-raise**, trong đúng phạm vi các hàm luồng BOM (không tính THDM/dialog). "Không log" = thân khối không chứa lời gọi nào tới `_log`/`logger`/`logging`/`log_fn`/`self._log` VÀ không có `raise` nào — bao gồm cả `except: pass` và `except Exception as e:` mà biến `e` không được dùng để log (chỉ dùng nội bộ, ví dụ format traceback không ghi ra đâu).

**bom_parser.py** — except không log: 12

| Dòng | Hàm | Loại exception | Snippet |
|---|---|---|---|
| 204 | `_extract_cell_meta` | `ValueError` | `except ValueError:` |
| 565 | `_parse_sheet` | `(ValueError, TypeError)` | `except (ValueError, TypeError): return True` |
| 702 | `_is_encrypted_excel` | `zipfile.BadZipFile` | `except zipfile.BadZipFile:` |
| 706 | `_is_encrypted_excel` | `Exception` | `except Exception:` |
| 694 | `_is_encrypted_excel` | `Exception` | `except Exception:` |
| 735 | `parse_bom_file` | `Exception` | `except Exception:` |
| 757 | `parse_bom_file` | `Exception` | `except Exception:` |
| 1339 | `_resolve_row_mapping` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |
| 1341 | `_resolve_row_mapping` | `(ValueError, TypeError)` | `except (ValueError, TypeError): row_out[sql_col] = mac` |
| 1317 | `_resolve_row_mapping` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |
| 1319 | `_resolve_row_mapping` | `(ValueError, TypeError)` | `except (ValueError, TypeError): val = mac` |
| 565 | `_is_non_numeric` | `(ValueError, TypeError)` | `except (ValueError, TypeError): return True` |

**mapping_loader.py** — except không log: 3

| Dòng | Hàm | Loại exception | Snippet |
|---|---|---|---|
| 37 | `_load_config` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |
| 160 | `load_mapping` | `Exception` | `except Exception as e:` |
| 139 | `load_mapping` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |

**main_window.py** — except không log: 28

| Dòng | Hàm | Loại exception | Snippet |
|---|---|---|---|
| 7316 | `_log_bom_lookup_fail` | `Exception` | `except Exception:` |
| 8928 | `_resolve_header_field` | `(ValueError, TypeError)` | `except (ValueError, TypeError): pass` |
| 8930 | `_resolve_header_field` | `(ValueError, TypeError)` | `except (ValueError, TypeError): pass` |
| 9042 | `_resolve_header_field` | `(ValueError, TypeError)` | `except (ValueError, TypeError): raw = None` |
| 9047 | `_resolve_header_field` | `(ValueError, TypeError)` | `except (ValueError, TypeError): raw = None` |
| 9030 | `_resolve_header_field` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |
| 9124 | `_find_existing_btp_code` | `Exception` | `except Exception:` |
| 9415 | `_resolve_detail_row` | `Exception` | `except Exception:` |
| 9287 | `_resolve_detail_row` | `(ValueError, TypeError)` | `except (ValueError, TypeError): raw = None` |
| 9292 | `_resolve_detail_row` | `(ValueError, TypeError)` | `except (ValueError, TypeError): raw = None` |
| 9270 | `_resolve_detail_row` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |
| 9282 | `_resolve_detail_row` | `(ValueError, TypeError)` | `except (ValueError, TypeError): pass` |
| 9272 | `_resolve_detail_row` | `(ValueError, TypeError)` | `except (ValueError, TypeError): raw = mac` |
| 9537 | `_generate_bom_details` | `Exception` | `except Exception:` |
| 9553 | `_generate_bom_details` | `Exception` | `except Exception:` |
| 9504 | `_generate_bom_details` | `(ValueError, TypeError)` | `except (ValueError, TypeError):` |
| 9855 | `_generate_bom_details` | `Exception` | `except Exception:` |
| 10869 | `_run_insert_bg` | `Exception` | `except Exception as e:` |
| 10890 | `_run_insert_bg` | `Exception` | `except Exception:` |
| 10872 | `_run_insert_bg` | `Exception` | `except Exception:` |
| 10883 | `_run_insert_bg` | `Exception` | `except Exception:` |
| 10706 | `_header_resolve_bg` | `Exception` | `except Exception:` |
| 10641 | `_header_resolve_bg` | `Exception` | `except Exception as e:` |
| 10699 | `_header_resolve_bg` | `Exception` | `except Exception as _ve:` |
| 10811 | `_on_header_resolved` | `Exception` | `except Exception:` |
| 10725 | `_on_header_resolved` | `Exception` | `except Exception:` |
| 10730 | `_on_header_resolved` | `Exception` | `except Exception:` |
| 10799 | `_on_header_resolved` | `Exception` | `except Exception:` |
