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

### Q-02

**Bối cảnh (`SQL-04`, EmployeeId):** sau khi `9dc3c92` (v2.2.22) vá lỗi EmployeeId hard-code=1, dữ liệu thật SAU vá vẫn cho `EmployeeId=1` là giá trị phổ biến nhất (48/82 dòng có EmployeeId, so với dải Id nhân viên thật khác chỉ 1-10 dòng mỗi Id). **Câu hỏi:** `EmployeeId=1` sau vá có phải là 1 nhân viên thật (ví dụ tài khoản admin/hệ thống) hay vẫn là 1 nhánh fallback cũ chưa được `9dc3c92` bao phủ hết? Cần business hoặc DBA xác nhận Id=1 trong bảng nhân viên tương ứng là ai.

### Q-03

**Bối cảnh (`SQL-05`, ParentDetailRowId_SO):** sau khi `78b36fc` (v2.2.23) vá công thức ParentDetailRowId_SO, tỷ lệ dòng sai theo mẫu CŨ (bằng mã đơn hàng thay vì công thức Mục số) giảm từ 7.6% (644 dòng, trước vá) xuống 5.6% (18 dòng, sau vá) — KHÔNG về 0%. Mẫu sau-vá còn nhỏ (18 dòng) nên chưa chắc chắn về thống kê. **Câu hỏi:** có luồng import nào khác (THDM, hoặc 1 loại sản phẩm/section cụ thể) không đi qua đúng nhánh code mà `78b36fc` đã vá? Cần business xác nhận nguồn gốc của (các) dòng sai còn lại sau vá, hoặc executor cần thêm thời gian điều tra nếu Phase 2 quyết định port field này.

### Q-04

**Bối cảnh (`EV-S11`/`EV-S13`, cột `Code` nhận giá trị non-str từ STT):** trường `Code` (BOM5, `Ten_Excel="STT"`, kiểu khai `varchar`) nhận giá trị int/float trực tiếp từ Excel ở phần lớn các dòng (không có cơ chế Fill-Forward ghi đè như `ItemType`, xem `EV-S13`). Chưa xác định được `Code` có được insert trực tiếp (rủi ro `.strip()`-crash giống `8a07639`) hay chỉ dùng làm khóa lookup nội bộ (không insert, không gọi `.strip()`). **Câu hỏi cho Phase 2 (không phải business — kỹ thuật):** đọc tiếp luồng dùng `BOM5.Code` sau `_resolve_detail_row` trước khi chốt verdict cho field này.

## 6. Tổng hợp & ánh xạ sang Phase 2

_(Điền ở Plan 02.)_

## 7. Xác nhận của business owner

_(Điền ở Plan 03 — ghi lại xác nhận bằng văn bản, kèm ngày, từ người dùng trong chat. Im lặng, auto-advance hay yolo mode không tính là xác nhận.)_

## 8. Phụ lục

_(Điền ở Plan 02, nếu cần.)_
