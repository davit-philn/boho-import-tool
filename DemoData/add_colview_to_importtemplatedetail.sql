/* ============================================================
   Thêm cột ColView vào ImportTemplateDetail
   ColView = tên cột lookup trong vB20Item để HIỂN THỊ
   ColKey  = tên cột gốc trong B20Item để IMPORT
   Nguồn mapping: vB20Item_LookupMapping.xlsx
   ============================================================ */

USE [BOMTool];
GO

-- ── Thêm cột ColView nếu chưa tồn tại ─────────────────────────────────────────
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.ImportTemplateDetail')
      AND name = N'ColView'
)
BEGIN
    ALTER TABLE dbo.ImportTemplateDetail
    ADD ColView NVARCHAR(100) NULL;
    PRINT 'Đã thêm cột ColView.';
END
ELSE
    PRINT 'Cột ColView đã tồn tại, bỏ qua ALTER TABLE.';
GO

-- ── UPDATE ColView theo mapping ColKey → lookup column trong vB20Item ──────────
-- Chỉ set khi ColView còn NULL (an toàn khi chạy lại nhiều lần)

-- ── Nhóm 1: trường từ B20Item ─────────────────────────────────────────────────
UPDATE dbo.ImportTemplateDetail SET ColView = N'ItemId0_Name'
    WHERE ColKey = N'ItemId0'             AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ProductId1_Name'
    WHERE ColKey = N'ProductId1'          AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'UnitName'
    WHERE ColKey = N'Unit'                AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ItemTypeName'
    WHERE ColKey = N'ItemType'            AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ItemGroupName'
    WHERE ColKey = N'ItemGroupId'         AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ItemCatgName'
    WHERE ColKey = N'ItemCatgId'          AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'BranchName'
    WHERE ColKey = N'BranchCode'          AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ParentId0_Name'
    WHERE ColKey = N'ParentId0'           AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ParentId_Name'
    WHERE ColKey = N'ParentId'            AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ItemClassPurchaseName'
    WHERE ColKey = N'ItemClassPurchaseId' AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'CustomField1Name'
    WHERE ColKey = N'CustomFieldId1'      AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'QCItemGroupName'
    WHERE ColKey = N'QCItemGroupId'       AND (ColView IS NULL OR ColView = N'');

-- ── Nhóm 2: trường từ B20ItemInfo (trong vB20Item đã JOIN sẵn) ────────────────
UPDATE dbo.ImportTemplateDetail SET ColView = N'ProductName'
    WHERE ColKey = N'ProductId'           AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'UnitOfLengthName'
    WHERE ColKey = N'UnitOfLength'        AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'UnitOfHeightName'
    WHERE ColKey = N'UnitOfHeight'        AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'UnitOfWidthName'
    WHERE ColKey = N'UnitOfWidth'         AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'UnitOfWeightName'
    WHERE ColKey = N'UnitOfWeight'        AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ProductClass1Name'
    WHERE ColKey = N'ProductClassId1'     AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ProductClass2Name'
    WHERE ColKey = N'ProductClassId2'     AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'SalesTaxName'
    WHERE ColKey = N'SalesTaxCode'        AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ExciseTaxName'
    WHERE ColKey = N'ExciseTaxId'         AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'EnvTaxName'
    WHERE ColKey = N'EnvTaxId'            AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ProductItemName'
    WHERE ColKey = N'ProductItemId'       AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ManufacturerName'
    WHERE ColKey = N'ManufacturerId'      AND (ColView IS NULL OR ColView = N'');

UPDATE dbo.ImportTemplateDetail SET ColView = N'ItemPurchasePriceTypeName'
    WHERE ColKey = N'ItemPurchasePriceTypeId' AND (ColView IS NULL OR ColView = N'');

GO

-- ── Kiểm tra kết quả ──────────────────────────────────────────────────────────
SELECT
    d.TemplateId,
    d.ColIndex,
    d.ColGroup,
    d.ColName,
    d.ColKey,
    d.ColView,
    CASE WHEN d.ColView IS NULL THEN N'(raw value)'
         ELSE N'→ ' + d.ColView
    END AS DisplayFrom,
    d.DataType
FROM dbo.ImportTemplateDetail d
ORDER BY d.TemplateId, d.ColIndex;

-- Kiểm tra các ColKey chưa có ColView (cần bổ sung nếu có)
SELECT DISTINCT d.ColKey, d.ColName
FROM dbo.ImportTemplateDetail d
WHERE d.ColView IS NULL
ORDER BY d.ColKey;
