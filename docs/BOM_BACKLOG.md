# BOM Backlog — BOHO Import Tool

Cập nhật: 2026-10-19. Phạm vi: chỉ BOM (THĐM và phần khác giữ nguyên).
Trạng thái: `XONG` | `CODE-CHƯA-BUILD` | `CHƯA-LÀM` | `CHỜ-QUYẾT` | `PHÍA-BẠN` (làm trong Bravo/DB).

Quy tắc: không build/push/release khi chưa được báo. Bản hiện hành: v2.2.25 (25/09/2026).

## A. Đã release
| Việc | Phiên bản |
|---|---|
| Bỏ `ParentDetailRowId_SO` khỏi 3 danh sách shortcut | đã release |
| EmployeeId (người tạo) dùng UILookup `creator_employee` | v2.2.22 |
| Sửa strip Unit ở MKT unit fallback | đã release |
| Ghi traceback vào `import_crash.log` trong `_run_insert_bg` | đã release |

## B. Đã sửa code, chưa build (`Tools/views/main_window.py`, chưa commit)
| # | Việc | Trạng thái | Ghi chú |
|---|---|---|---|
| B1 | `_generate_bom_details`: bỏ MKT fallback cho BOM3 loại H/E2 và BOM2 loại H (`_sp_creates_btp`) để SP Bravo tạo mã BTP | CODE-CHƯA-BUILD | Mới kiểm tra cú pháp, chưa import thử thật |
| B2 | `_find_existing_btp_code`: chỉ tái dùng mã khi tiền tố khớp đúng `ItemCode0.MụcSố`; không có Mục thì không tái dùng | CODE-CHƯA-BUILD | Mới kiểm tra cú pháp |
| B3 | Lấy Mục số từ `B30BizDocDetailSO.BuiltinOrder0` theo `DetailRowId_SO`, truyền `muc_so` vào `_resolve_detail_row` | CODE-CHƯA-BUILD | Đi kèm B2 |

## C. Code chưa làm
| # | Việc | Trạng thái |
|---|---|---|
| C1 | `_mkt_cache`: khóa `C` và `NULL` ghi đè lẫn nhau, thứ tự không cố định | CHƯA-LÀM |
| C2 | Thay các `except: pass` bằng log rõ ràng | CHƯA-LÀM |
| C3 | MKT_KINH dùng sai (23/29 dòng không phải kính); đã xử lý riêng H/E2, các loại khác chưa rà | CHƯA-LÀM |
| C4 | Nguyên nhân gốc khiến MKT fallback gán sai (chưa chứng minh) | CHƯA-LÀM |
| C5 | Chạy thử B1–B3 với import BOM thật (H/E2) | CHƯA-LÀM |
| C6 | Bộ test cho `_find_existing_btp_code`, nhánh bỏ MKT, `_mkt_cache` (dùng file BOM mẫu thật) | CHƯA-LÀM |
| C7 | Tách dần logic BOM khỏi `main_window.py` sang `Tools/services/` | CHỜ-QUYẾT |

## C2. Rule BTP đã chốt (19/10/2026) — chưa code
| # | Rule | Trạng thái |
|---|---|---|
| R1 | Cột `BTP` trong sheet BOM (đọc trực tiếp từ header Excel, không cần khai báo mapping vì không có cột SQL). Nhập `1` = dòng BTP | CODE-CHƯA-BUILD |
| R2 | Dòng `BTP = 1`: tái dùng mã BTP cùng tiền tố+tên nếu có, không thì SP tạo mã mới; gửi ItemName rỗng cho SP; không gán MKT. Lưu ý: BOM3 chỉ H/E2 mới được SP tạo, loại khác bị R6 chặn | CODE-CHƯA-BUILD |
| R3 | Nhóm H và E2: tự tạo mã BTP, không cần đánh dấu (đã code, chưa build — mục B1) | CODE-CHƯA-BUILD |
| R4 | ĐVT bắt buộc có: Excel -> ĐVT mã trong danh mục; cả hai trống thì chặn import, không dùng mặc định | CHƯA-LÀM |
| R5 | Chuẩn hóa tên BTP (rỗng, xuống dòng, khoảng trắng) trước khi gọi SP | CHƯA-LÀM |
| R6 | Kiểm tra kết quả SP: dòng BTP (H/E2, BOM2 H, hoặc BTP=1) thiếu mã thì dừng import, báo danh sách dòng (`_run_batch_hook_grouped`, `services/bom_rules.py`) | CODE-CHƯA-BUILD |

Đã loại bỏ (không làm): cảnh báo lệch tiền tố khi import (do người dùng tự xử lý), kiểm tra khớp ĐVT khi tái dùng (cùng Mục không trùng tên khác quy cách), THĐM đi theo BOM (việc của Bravo).

## D. Dữ liệu BOM (đơn hàng cụ thể)
| # | Việc | Trạng thái |
|---|---|---|
| D1 | DRC (11036174FO): MO57 BOM 598, dòng H 49284 `LO.TB.MO57.50.0002` còn tiền tố cũ, BOM v1 693 cũng dùng mã này. Đã đề nghị script đổi mã | CHỜ-QUYẾT |
| D2 | LAGOME (11036066FO): script tạo mã BTP cho 10 dòng kính (BOM 545, 547, 549, 550, 552, 553) đã gửi, chưa xác nhận có chạy | CHỜ-QUYẾT |
| D3 | LAGOME: kính dùng mã SP sinh hay mã kính có sẵn (`2573.x.H001`, `TP07037`) | CHỜ-QUYẾT |
| D4 | LAGOME: chưa kiểm tra tiền tố lệch và mã dùng chung | CHƯA-LÀM |
| D5 | `B00Config` `M_ParentId_GT_Default`/`M_ParentId_GCKL_Default`: `ParentId` 62419/62407 của mã H/E2 mới trông sai | PHÍA-BẠN |
| D6 | ĐVT trên `B20Item` của một số mã (`LO.CH.AS27…`, `LO.SO.AS31…`, `LO.TB.MO58…`) đang `-`/`1`, BOM ghi CÁI | PHÍA-BẠN |
| D7 | Deactivate mã MKT/mã dùng chung cũ trong `B20Item` | PHÍA-BẠN |

## E. Ghi chú kỹ thuật cần nhớ
- "BOM mới nhất": `ROW_NUMBER() PARTITION BY (ItemId0, ParentBizDocId, DetailRowId_SO, TypeB) ORDER BY TRY_CAST(Version AS int) DESC, CreatedAt DESC, Id DESC`, `IsActive=1`.
- Loại trừ dự án test: ProductId1 424909, 428759.
- Không có quyền UPDATE `B20Item`; thay đổi catalog do người dùng chạy script.
- SP `usp_B20BOM_Create_ItemCode`: mã dạng `ItemCode0.BuiltinOrder0.000N`; BOM2 xét ItemId rỗng và (tên rỗng hoặc ItemType H); BOM3 xét ItemType H/E2 và ItemId rỗng.

## Rule BTP mới (BOM2)
- Dòng BOM2 là BTP khi **Đvt = CỤM** (không phân biệt hoa thường/dấu). Tên vật tư không còn là điều kiện; dòng CỤM có Tên vật tư thường (không trống, không có "+") chỉ bị cảnh báo, không chặn.
- BTP: để trống ItemId cho usp_B20BOM_Create_ItemCode tạo mã, hoặc tái dùng mã cũ theo Tên chi tiết; không fallback MKT.
- Tên vật tư trống/"+" mà Đvt khác CỤM, hoặc BTP thiếu Tên chi tiết → lỗi Kiểm tra (chặn Import), không có rule dự phòng.
- Đã bỏ: cột BTP=1, rule dòng con/Tên chi tiết, rule "+" riêng STT nguyên, phân biệt container/mảnh cắt. Giữ: nhóm SP tạo mã theo ItemType (BOM3: H,E2; BOM2: H).

## Nhóm vật tư — bảng quy đổi (sheet NHOM_QUY_DOI trong CK_Mapping_v5.xlsx)
- Quy đổi theo TÊN nhóm (mẫu cũ A..E và mẫu mới D..H cùng quy về loại chuẩn). Cột: Section, Ten_Nhom, Chu_Nhom, Loai_Chuan, Ma_Tam, Cach_Xu_Ly (MKT/SP/NHOM_CHA).
- ItemType của dòng = Loai_Chuan của nhóm; mã tạm MKT lấy theo Ma_Tam. Nhóm lạ → lỗi Kiểm tra (không còn mặc định MKT_KINH). Nhóm con La Mã (I, II…) thừa hưởng nhóm chữ phía trên.
- Dòng không tra được mã theo Tên vật tư (kết quả none) → cảnh báo "sẽ dùng mã tạm MKT_xxx" ở bước Kiểm tra (không chặn).
- Chưa có sheet NHOM_QUY_DOI → giữ cách cũ (theo ItemType).
- Cần xác nhận: C (VAIDASIMILI/VAI/DA/SIMILI), F (VTP/KEO), I (BAOBI/BB) đang chọn mặc định ở cột Ma_Tam.
