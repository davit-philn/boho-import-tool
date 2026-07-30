/* ============================================================
   UPDATE ImportTemplateDetail — Chuẩn hoá ColKey theo Layout XML
   Nguồn: B20Item_Layout_Fields.xlsx (BRAVO 10)
   Ngày : 2026-07-27
   ============================================================
   Các cột sai được phát hiện:
     CostingId       → ItemGroupId       (Bộ định khoản)
     ItemGroupCode   → ParentId0         (Nhóm VT cấp 1 — NHÓM THÔNG TIN SẢN PHẨM)
     ItemGroupCode   → ParentId          (Nhóm VT cấp 2 chi tiết — TIỀN TỐ)
     NormRate        → B20ItemInfo.Height (Độ dày = ChildTable B20ItemInfo.Height)
     WoodType2       → WoodType          (B20Item chỉ có 1 cột WoodType)
     Attr_Supplier   → vB20ItemVendor.VendorId (ChildTable nhà cung cấp)
   ============================================================ */

USE [BOMTool];
GO

-- ── 1. BỘ ĐỊNH KHOẢN: CostingId → ItemGroupId ─────────────────────────────
-- Layout field #13: "Bộ định khoản" = ItemGroupId (BravoLookupBox, B20Item)
UPDATE ImportTemplateDetail
SET ColKey = N'ItemGroupId'
WHERE ColKey  = N'CostingId'
  AND ColName = N'BỘ ĐỊNH KHOẢN';

SELECT @@ROWCOUNT AS [1_BoDinhKhoan_updated];
GO

-- ── 2. MÃ NHÓM VẬT TƯ (NHÓM THÔNG TIN): ItemGroupCode → ParentId0 ─────────
-- Layout field #34: "Nhóm vật tư (cấp 1)" = ParentId0 (BravoLookupBox, B20Item)
-- Đây là nhóm cha cấp 1 (top-level group), nhập bằng Mã → hệ thống resolve sang Id
UPDATE ImportTemplateDetail
SET ColKey = N'ParentId0'
WHERE ColKey   = N'ItemGroupCode'
  AND ColGroup = N'NHÓM THÔNG TIN SẢN PHẨM'
  AND ColName  = N'MÃ NHÓM VẬT TƯ';

SELECT @@ROWCOUNT AS [2_MaNhomVatTu_updated];
GO

-- ── 3. NHÓM VẬT TƯ CHI TIẾT (TIỀN TỐ — NVL01): ItemGroupCode → ParentId ───
-- Layout field #35: "Nhóm vật tư chi tiết (cấp 2)" = ParentId (BravoLookupBox, B20Item)
UPDATE ImportTemplateDetail
SET ColKey = N'ParentId'
WHERE ColKey   = N'ItemGroupCode'
  AND ColGroup = N'TIỀN TỐ'
  AND ColName  = N'NHÓM VẬT TƯ CHI TIẾT';

SELECT @@ROWCOUNT AS [3_NhomVatTuChiTiet_NVL01_updated];
GO

-- ── 4. NHÓM VẬT TƯ (TIỀN TỐ — NVL03, NVL04, NVL05, NVL06,
--       NVL07, NVL08, NVL09, NVL10): ItemGroupCode → ParentId ──────────────
-- Cùng ý nghĩa: nhóm cha cấp 2 cho các template còn lại
UPDATE ImportTemplateDetail
SET ColKey = N'ParentId'
WHERE ColKey   = N'ItemGroupCode'
  AND ColGroup = N'TIỀN TỐ'
  AND ColName  = N'NHÓM VẬT TƯ';

SELECT @@ROWCOUNT AS [4_NhomVatTu_TIENTO_updated];
GO

-- ── 5. ĐỘ DÀY: NormRate → B20ItemInfo.Height ──────────────────────────────
-- Layout field #36 (Tab Tiền tố): "Chiều cao (độ dày)" = B20ItemInfo.Height
-- B20ItemInfo là ChildTable 1-1, truy cập qua JOIN khi query
UPDATE ImportTemplateDetail
SET ColKey   = N'B20ItemInfo.Height',
    DataType = N'DECIMAL'
WHERE ColKey  = N'NormRate'
  AND ColName = N'ĐỘ DÀY';

SELECT @@ROWCOUNT AS [5_DoDay_updated];
GO

-- ── 6. NHÓM VL 2: WoodType2 → WoodType ────────────────────────────────────
-- B20Item chỉ có 1 cột WoodType (field #48: "Loại" = WoodType).
-- Không tồn tại WoodType2 → ánh xạ về cùng cột WoodType.
-- Lưu ý: khi hiện Treeview, cột NHÓM VL 2 sẽ bị bỏ qua (duplicate vs NHÓM VL 1).
UPDATE ImportTemplateDetail
SET ColKey = N'WoodType'
WHERE ColKey  = N'WoodType2'
  AND ColName = N'NHÓM VL 2';

SELECT @@ROWCOUNT AS [6_WoodType2_updated];
GO

-- ── 7. NHÀ CUNG CẤP: Attr_Supplier → vB20ItemVendor.VendorId ──────────────
-- Layout Child Table: vB20ItemVendor → VendorId (BravoLookupBox)
-- Không nằm trong B20Item chính, là bảng con 1-N nhà cung cấp của vật tư.
UPDATE ImportTemplateDetail
SET ColKey = N'vB20ItemVendor.VendorId'
WHERE ColKey  = N'Attr_Supplier'
  AND ColName = N'NHÀ CUNG CẤP';

SELECT @@ROWCOUNT AS [7_NhaCungCap_updated];
GO

-- ══ Kiểm tra kết quả ══════════════════════════════════════════════════════════

-- Xem toàn bộ mapping sau khi fix (theo template NVL01 làm mẫu)
SELECT
    d.ColIndex,
    d.ColGroup,
    d.ColName,
    d.ColKey,
    d.DataType,
    d.IsNamePart,
    d.NameConcatOrder
FROM ImportTemplateDetail d
JOIN ImportTemplate t ON t.Id = d.TemplateId
WHERE t.TemplateCode = N'NVL01'
ORDER BY d.ColIndex;

-- Kiểm tra còn ColKey nào chưa được chuẩn hoá (vẫn dùng tên cũ)
SELECT t.TemplateCode, d.ColName, d.ColKey
FROM ImportTemplateDetail d
JOIN ImportTemplate t ON t.Id = d.TemplateId
WHERE d.ColKey IN (N'CostingId', N'NormRate', N'WoodType2', N'Attr_Supplier')
   OR (d.ColKey = N'ItemGroupCode')   -- kiểm tra còn sót không
ORDER BY t.TemplateCode, d.ColIndex;
