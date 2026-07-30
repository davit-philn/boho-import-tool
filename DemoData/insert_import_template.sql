-- =========================================================
-- INSERT ImportTemplate + ImportTemplateDetail
-- Source: [BoHo] Chuẩn hóa NVL NMSX.xlsx
-- Generated: 2026-07-27
-- =========================================================
USE [BOMTool];
GO

-- Xóa dữ liệu cũ nếu có (idempotent — chạy lại an toàn)
DELETE FROM ImportTemplateDetail WHERE TemplateId IN
  (SELECT Id FROM ImportTemplate WHERE TemplateCode IN
   (N'NVL01',N'NVL02',N'NVL03',N'NVL04',N'NVL05',
    N'NVL06',N'NVL07',N'NVL08',N'NVL09',N'NVL10',N'TP'));
DELETE FROM ImportTemplate WHERE TemplateCode IN
  (N'NVL01',N'NVL02',N'NVL03',N'NVL04',N'NVL05',
   N'NVL06',N'NVL07',N'NVL08',N'NVL09',N'NVL10',N'TP');
GO

-- ── 1. ImportTemplate ──────────────────────────────────────
INSERT INTO ImportTemplate (TemplateCode,TemplateName,ItemGroupCode,CodePattern,NameGenAuto,IsActive) VALUES
  (N'NVL01',N'Vật tư tấm',N'NVL01',N'NVL01_XXXXXX',1,1),
  (N'NVL02',N'Chỉ cạnh / Nẹp',N'NVL02',N'NVL02_XXXXXX',1,1),
  (N'NVL03',N'Foam',N'NVL03',N'NVL03_XXXXXX',1,1),
  (N'NVL04',N'Vải / Da / Simili',N'NVL04',N'NVL04_XXXXXX',0,1),
  (N'NVL05',N'Cơ khí / Kim loại',N'NVL05',N'NVL05_XXXXXX',1,1),
  (N'NVL06',N'Vật tư Sofa+Đinh+Keo',N'NVL06',N'NVL06_XXXXXX',0,1),
  (N'NVL07',N'Vật tư phụ',N'NVL07',N'NVL07_XXXXXX',0,1),
  (N'NVL08',N'Phụ kiện',N'NVL08',N'NVL08_XXXXXX',0,1),
  (N'NVL09',N'Gỗ',N'NVL09',N'NVL09_XXXXXX',1,1),
  (N'NVL10',N'Bao bì',N'NVL10',N'NVL10_XXXXXX',0,1),
  (N'TP',N'Thầu phụ',N'TP',N'TP_XXXXXX',0,1);
GO

-- ── 2. ImportTemplateDetail ───────────────────────────────

-- === NVL01: Vật tư tấm ===
DECLARE @t_NVL01 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL01');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ CHI TIẾT',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,10,N'TIỀN TỐ',N'CỐT VÁN',N'BoardCore',N'NVARCHAR',1,1);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,11,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',1,2);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,12,N'TIỀN TỐ',N'SỐ MẶT CHÍNH',N'MainFaceCount',N'NVARCHAR',1,3);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,13,N'TIỀN TỐ',N'NHÓM VL 1',N'WoodType',N'NVARCHAR',1,4);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,14,N'TIỀN TỐ',N'LỚP PHỦ CHÍNH',N'MainFinish',N'NVARCHAR',1,5);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,15,N'TIỀN TỐ',N'SỐ MẶT PHỤ',N'SecondaryFaceCount',N'NVARCHAR',1,6);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,16,N'TIỀN TỐ',N'NHÓM VL 2',N'WoodType2',N'NVARCHAR',1,7);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,17,N'TIỀN TỐ',N'LỚP PHỦ PHỤ',N'SecondaryFinish',N'NVARCHAR',1,8);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,18,N'TIỀN TỐ',N'ĐỘ DÀY',N'NormRate',N'DECIMAL',1,9);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,19,N'TIỀN TỐ',N'QUY CÁCH (mm)',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL01,20,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL02: Chỉ cạnh / Nẹp ===
DECLARE @t_NVL02 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL02');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,9,N'TIỀN TỐ',N'CHỈ CẠNH',N'EdgeBanding',N'NVARCHAR',1,1);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,10,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',1,2);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,11,N'TIỀN TỐ',N'QUY CÁCH (mm)',N'Specification',N'NVARCHAR',1,3);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL02,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL03: Foam ===
DECLARE @t_NVL03 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL03');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',1,1);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,10,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',1,2);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,11,N'TIỀN TỐ',N'QUY CÁCH (mm)',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL03,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL04: Vải / Da / Simili ===
DECLARE @t_NVL04 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL04');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,10,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,11,N'TIỀN TỐ',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL04,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL05: Cơ khí / Kim loại ===
DECLARE @t_NVL05 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL05');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',1,1);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,10,N'TIỀN TỐ',N'LOẠI CHẤT LIỆU',N'Material',N'NVARCHAR',1,2);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,11,N'TIỀN TỐ',N'QUY CÁCH (mm)',N'Specification',N'NVARCHAR',1,3);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL05,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL06: Vật tư Sofa+Đinh+Keo ===
DECLARE @t_NVL06 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL06');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,10,N'TIỀN TỐ',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,11,N'TIỀN TỐ',N'LOẠI CHẤT LIỆU',N'Material',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL06,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL07: Vật tư phụ ===
DECLARE @t_NVL07 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL07');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,10,N'TIỀN TỐ',N'QUY CÁCH (mm)',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,11,N'TIỀN TỐ',N'LOẠI CHẤT LIỆU',N'Material',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL07,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL08: Phụ kiện ===
DECLARE @t_NVL08 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL08');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,10,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,11,N'TIỀN TỐ',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL08,13,N'TIỀN TỐ',N'NHÀ CUNG CẤP',N'Attr_Supplier',N'NVARCHAR',0,NULL);
GO

-- === NVL09: Gỗ ===
DECLARE @t_NVL09 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL09');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',1,1);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,10,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',1,2);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,11,N'TIỀN TỐ',N'QUY CÁCH',N'Specification',N'NVARCHAR',1,3);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL09,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === NVL10: Bao bì ===
DECLARE @t_NVL10 INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'NVL10');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,10,N'TIỀN TỐ',N'LOẠI/ CHẤT LIỆU',N'Material',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,11,N'TIỀN TỐ',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_NVL10,12,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- === TP: Thầu phụ ===
DECLARE @t_TP INT = (SELECT Id FROM ImportTemplate WHERE TemplateCode=N'TP');
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,1,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ',N'Code',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,2,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NHÓM VẬT TƯ',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,3,N'NHÓM THÔNG TIN SẢN PHẨM',N'TÊN NGUYÊN VẬT LIỆU',N'Name',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,4,N'NHÓM THÔNG TIN SẢN PHẨM',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,5,N'NHÓM THÔNG TIN SẢN PHẨM',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,6,N'NHÓM THÔNG TIN SẢN PHẨM',N'BỘ ĐỊNH KHOẢN',N'CostingId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,7,N'NHÓM THÔNG TIN SẢN PHẨM',N'MÃ NGÀNH HÀNG',N'ItemCatgId',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,8,N'NHÓM THÔNG TIN SẢN PHẨM',N'LOẠI VẬT TƯ',N'ItemType',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,9,N'TIỀN TỐ',N'NHÓM VẬT TƯ CHI TIẾT',N'ItemGroupCode',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,10,N'TIỀN TỐ',N'QUY CÁCH',N'Specification',N'NVARCHAR',0,NULL);
INSERT INTO ImportTemplateDetail (TemplateId,ColIndex,ColGroup,ColName,ColKey,DataType,IsNamePart,NameConcatOrder) VALUES (@t_TP,11,N'TIỀN TỐ',N'ĐƠN VỊ TÍNH',N'Unit',N'NVARCHAR',0,NULL);
GO

-- ── Kiểm tra kết quả ───────────────────────────────────────
SELECT t.TemplateCode, t.TemplateName, t.NameGenAuto,
       COUNT(d.Id) AS TotalCols,
       SUM(d.IsNamePart) AS NamePartCols
FROM ImportTemplate t
LEFT JOIN ImportTemplateDetail d ON d.TemplateId = t.Id
GROUP BY t.Id, t.TemplateCode, t.TemplateName, t.NameGenAuto
ORDER BY t.TemplateCode;