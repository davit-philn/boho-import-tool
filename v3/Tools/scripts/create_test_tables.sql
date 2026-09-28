-- ============================================================
-- Tạo DB test với cấu trúc ĐÚNG theo BRAVO thực tế
-- Server: CALSIFER\SQLEXPRESS  |  DB: B10_Boho_Data
-- ============================================================

USE [B10_Boho_Data];
GO

-- ============================================================
-- 1. B20Item — danh mục vật tư
-- ============================================================
IF OBJECT_ID('B20Item', 'U') IS NOT NULL DROP TABLE B20Item;
GO
CREATE TABLE B20Item (
    Id                  INT             NOT NULL IDENTITY(1,1) PRIMARY KEY,
    ParentId            INT             NOT NULL DEFAULT 0,
    IsGroup             BIT             NOT NULL DEFAULT 0,
    BranchCode          VARCHAR(3)      NOT NULL DEFAULT '',
    Code                VARCHAR(24)     NOT NULL,
    BarCode             VARCHAR(24)     NULL,
    Name                NVARCHAR(512)   NOT NULL DEFAULT '',
    Unit                NVARCHAR(8)     NOT NULL DEFAULT '',
    ItemType            TINYINT         NOT NULL DEFAULT 0,
    ShortCode           VARCHAR(6)      NULL,
    ItemOrder           INT             NOT NULL DEFAULT 0,
    ItemGroupId         INT             NULL,
    ItemClassPurchaseId INT             NULL,
    IsActive            BIT             NOT NULL DEFAULT 1,
    CreatedBy           INT             NOT NULL DEFAULT 0,
    CreatedAt           SMALLDATETIME   NULL,
    ModifiedBy          INT             NOT NULL DEFAULT 0,
    ModifiedAt          SMALLDATETIME   NOT NULL DEFAULT GETDATE(),
    [timestamp]         ROWVERSION      NOT NULL,
    QCItemGroupId       INT             NULL,
    ItemCatgId          INT             NULL,
    FileName            NVARCHAR(512)   NULL,
    Specification       NVARCHAR(MAX)   NULL,
    ProductId1          INT             NULL,
    IsEquipment         TINYINT         NULL,
    Constyle            NVARCHAR(50)    NULL,
    Isgen               TINYINT         NOT NULL DEFAULT 0,
    ItemId0             INT             NULL,
    CustomFieldId1      INT             NULL,
    ColorId             INT             NULL,
    ProductId           INT             NULL,
    DrawingCode         VARCHAR(24)     NULL,
    CostingId           INT             NULL,
    ItemTypeSX          VARCHAR(24)     NULL,
    BizDocId_SO         VARCHAR(24)     NULL,
    IsPackage           TINYINT         NOT NULL DEFAULT 0,
    BOMItemId           INT             NULL,
    NameB7              NVARCHAR(512)   NULL,
    BoardCore           NVARCHAR(512)   NULL,
    MainFaceCount       NVARCHAR(512)   NULL,
    MainFinish          NVARCHAR(512)   NULL,
    SecondaryFaceCount  NVARCHAR(512)   NULL,
    SecondaryFinish     NVARCHAR(512)   NULL,
    EdgeBanding         NVARCHAR(512)   NULL,
    Material            NVARCHAR(512)   NULL,
    WoodType            NVARCHAR(512)   NULL,
    NormRate            NUMERIC(18,4)   NULL,
    Name2               NVARCHAR(512)   NOT NULL DEFAULT ''
);
GO
CREATE INDEX IX_B20Item_Code      ON B20Item (Code);
CREATE INDEX IX_B20Item_IsActive  ON B20Item (IsActive);
GO

-- Dữ liệu mẫu
INSERT INTO B20Item (ParentId, IsGroup, BranchCode, Code, Name, Unit, ItemType, Isgen, IsPackage, CreatedBy, ModifiedBy, Name2)
VALUES
(0,0,'BHO','SP-001',N'Sofa góc L cao cấp',       N'Cái',0,0,0,1,1,N'Sofa goc L'),
(0,0,'BHO','VT-001',N'MDF HMR E2 17mm',           N'Tấm',0,0,0,1,1,N'MDF HMR E2'),
(0,0,'BHO','VT-002',N'Chỉ nhựa 0866CA nâu',       N'Md', 0,0,0,1,1,N'Chi nhua'),
(0,0,'BHO','VT-003',N'Ván gỗ ghép tràm 18mm',     N'Tấm',0,0,0,1,1,N'Van go ghep'),
(0,0,'BHO','VT-004',N'Keo PVA D3',                N'Lọ', 0,0,0,1,1,N'Keo PVA'),
(0,0,'BHO','VT-005',N'Ốc vít M6x30',             N'Cái',0,0,0,1,1,N'Oc vit M6'),
(0,0,'BHO','VT-006',N'Vải bọc Velvet Grey',       N'm',  0,0,0,1,1,N'Vai boc'),
(0,0,'BHO','VT-007',N'Xốp mút D45',              N'Tấm',0,0,0,1,1,N'Xop mut'),
(0,0,'BHO','VT-008',N'Lớp phủ Melamine trắng',   N'Tấm',0,0,0,1,1,N'Melamine'),
(0,0,'BHO','VT-009',N'Thanh gỗ tràm 18x90',      N'Cái',0,0,0,1,1,N'Thanh go tram');
GO

-- ============================================================
-- 2. B20BOM — header phiếu BOM
-- ============================================================
IF OBJECT_ID('B20BOM', 'U') IS NOT NULL DROP TABLE B20BOM;
GO
CREATE TABLE B20BOM (
    Id                  INT             NOT NULL IDENTITY(1,1) PRIMARY KEY,
    ParentId            INT             NOT NULL DEFAULT 0,
    IsGroup             BIT             NOT NULL DEFAULT 0,
    BranchCode          VARCHAR(3)      NOT NULL DEFAULT '',
    DocCode             VARCHAR(4)      NOT NULL DEFAULT 'BOM',
    DocDate             DATE            NOT NULL DEFAULT GETDATE(),
    RowId               VARCHAR(16)     NOT NULL DEFAULT '',
    DocStatus           TINYINT         NOT NULL DEFAULT 1,
    DocProcessId        VARCHAR(16)     NOT NULL DEFAULT '',
    RowId_DocProcess    VARCHAR(16)     NOT NULL DEFAULT '',
    ProductionProcessId INT             NULL,
    Code                VARCHAR(24)     NOT NULL DEFAULT '',
    ItemId0             INT             NULL,
    ProductId           INT             NULL,
    Version             NVARCHAR(32)    NOT NULL DEFAULT '1',
    DrawingNumber       NVARCHAR(64)    NOT NULL DEFAULT '',
    Description         NVARCHAR(254)   NOT NULL DEFAULT '',
    EffectiveDate       DATE            NULL,
    FinishedDate        DATE            NULL,
    BOMType             NVARCHAR(16)    NOT NULL DEFAULT 'SX',
    ClassCode1          VARCHAR(24)     NOT NULL DEFAULT '',
    ClassCode2          VARCHAR(24)     NOT NULL DEFAULT '',
    ClassCode3          VARCHAR(24)     NOT NULL DEFAULT '',
    IsActive            BIT             NOT NULL DEFAULT 1,
    CreatedBy           INT             NOT NULL DEFAULT 0,
    CreatedAt           DATETIME        NOT NULL DEFAULT GETDATE(),
    ModifiedBy          INT             NOT NULL DEFAULT 0,
    ModifiedAt          DATETIME        NOT NULL DEFAULT GETDATE(),
    [timestamp]         ROWVERSION      NOT NULL,
    EmployeeId          INT             NULL,
    WorkProcessId       INT             NULL,
    ApprovalStatus      VARCHAR(16)     NULL,
    ApprovedUserIdList  VARCHAR(128)    NULL,
    UserIdList          VARCHAR(128)    NULL,
    IsDraftData         TINYINT         NOT NULL DEFAULT 0,
    ESignLayoutName     NVARCHAR(64)    NOT NULL DEFAULT '',
    Material            NVARCHAR(MAX)   NULL,
    [Structure]         NVARCHAR(MAX)   NULL,
    ParentBizDocId      VARCHAR(24)     NULL,
    ProductId1          INT             NULL,
    DetailRowId_SO      NVARCHAR(512)   NULL,
    Quantity            NUMERIC(18,4)   NOT NULL DEFAULT 0,
    Size                NVARCHAR(MAX)   NULL,
    Finish              NVARCHAR(MAX)   NULL,
    TypeB               VARCHAR(24)     NOT NULL DEFAULT '',
    PeriodId            INT             NULL,
    EdgeCountAdj        NUMERIC(18,4)   NOT NULL DEFAULT 0,
    Version0            INT             NULL,
    SubcontractingCode  VARCHAR(24)     NULL,
    MaterialSP          NVARCHAR(MAX)   NULL
);
GO
CREATE INDEX IX_B20BOM_Code       ON B20BOM (Code);
CREATE INDEX IX_B20BOM_ItemId0    ON B20BOM (ItemId0);
CREATE INDEX IX_B20BOM_DocDate    ON B20BOM (DocDate);
GO

-- ============================================================
-- 3. B20BOMDetail — chi tiết BOM (BOMDetailType 2/3/4/5)
-- ============================================================
IF OBJECT_ID('B20BOMDetail', 'U') IS NOT NULL DROP TABLE B20BOMDetail;
GO
CREATE TABLE B20BOMDetail (
    Id                      INT             NOT NULL IDENTITY(1,1) PRIMARY KEY,
    ParentId                INT             NOT NULL DEFAULT 0,
    Code                    VARCHAR(24)     NULL,
    Name                    NVARCHAR(MAX)   NULL,
    BranchCode              CHAR(3)         NOT NULL DEFAULT '',
    BuiltinOrder            SMALLINT        NOT NULL DEFAULT 0,
    BOMId                   INT             NULL,
    EffectiveDate           DATE            NULL,
    FinishedDate            DATE            NULL,
    ProductId               INT             NULL,
    ItemId                  INT             NULL,
    ProductId0              INT             NULL,
    Id_BOM                  INT             NULL,
    [Type]                  NVARCHAR(16)    NOT NULL DEFAULT '',
    BOMDetailType           TINYINT         NOT NULL DEFAULT 2,
    Specification           NVARCHAR(64)    NOT NULL DEFAULT '',
    Unit                    NVARCHAR(8)     NOT NULL DEFAULT '',
    Quantity                NUMERIC(18,4)   NOT NULL DEFAULT 0,
    ConvertRate9            NUMERIC(18,6)   NOT NULL DEFAULT 1,
    Quantity9               NUMERIC(18,4)   NOT NULL DEFAULT 0,
    ScrapRate               NUMERIC(18,4)   NOT NULL DEFAULT 0,
    ProductQuantity         NUMERIC(18,4)   NOT NULL DEFAULT 0,
    ProductionProcessId     INT             NULL,
    Remark                  NVARCHAR(254)   NOT NULL DEFAULT '',
    IsActive                BIT             NOT NULL DEFAULT 1,
    CreatedBy               INT             NOT NULL DEFAULT 0,
    CreatedAt               DATETIME        NOT NULL DEFAULT GETDATE(),
    ModifiedBy              INT             NOT NULL DEFAULT 0,
    ModifiedAt              DATETIME        NOT NULL DEFAULT GETDATE(),
    [timestamp]             ROWVERSION      NOT NULL,
    Thickness               NUMERIC(18,4)   NULL,
    Width                   NUMERIC(18,4)   NULL,
    [Length]                NUMERIC(18,4)   NULL,
    WeightCalc              NUMERIC(18,4)   NULL,
    EdgeLength              NUMERIC(18,4)   NULL,
    FinishArea              NUMERIC(18,4)   NULL,
    PrimerArea              NUMERIC(18,4)   NULL,
    EdgeMaterialCode        NVARCHAR(100)   NULL,
    EdgeCount               VARCHAR(100)    NULL,
    FinishCode              NVARCHAR(100)   NULL,
    FinishSide              VARCHAR(100)    NULL,
    PrimerSide              VARCHAR(100)    NULL,
    CustomerId              INT             NULL,
    Material                NVARCHAR(MAX)   NULL,
    QuantityFactory         NUMERIC(18,4)   NULL,
    QuantityConstruction    NUMERIC(18,4)   NULL,
    QuantitySubcontractor   NUMERIC(18,4)   NULL,
    PackageName             NVARCHAR(100)   NULL,
    PackageThickness        NUMERIC(18,4)   NULL,
    PackageWidth            NUMERIC(18,4)   NULL,
    PackageLength           NUMERIC(18,4)   NULL,
    QuantityPackage         NUMERIC(18,4)   NULL,
    PackageStandard         NVARCHAR(100)   NULL,
    BOMTemplateId           INT             NULL,
    ParentBizDocId          VARCHAR(24)     NULL,
    DetailRowId_SO          NVARCHAR(512)   NULL,
    ItemType                VARCHAR(24)     NULL,
    PeriodId                INT             NULL,
    ItemId0                 INT             NULL,
    ItemName                NVARCHAR(MAX)   NULL,
    TypeBB                  VARCHAR(24)     NULL
);
GO
CREATE INDEX IX_B20BOMDetail_BOMId         ON B20BOMDetail (BOMId);
CREATE INDEX IX_B20BOMDetail_ItemId        ON B20BOMDetail (ItemId);
CREATE INDEX IX_B20BOMDetail_BOMDetailType ON B20BOMDetail (BOMDetailType);
GO

-- ============================================================
-- Verify
-- ============================================================
SELECT 'B20Item'      AS [Table], COUNT(*) AS Rows FROM B20Item
UNION ALL
SELECT 'B20BOM',       COUNT(*) FROM B20BOM
UNION ALL
SELECT 'B20BOMDetail', COUNT(*) FROM B20BOMDetail;
GO

PRINT N'Done — B10_Boho_Data san sang de test';
GO
