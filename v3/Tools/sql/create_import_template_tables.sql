/* ============================================================
   [BOMTool] — Bảng cấu hình Template import NVL
   Chạy 1 lần trên DB [BOMTool] (cùng server với Bravo).
   Idempotent: dùng IF OBJECT_ID IS NULL — chạy lại không lỗi.
   ============================================================ */

USE [BOMTool];
GO

-- ── ImportTemplate (header) ────────────────────────────────────
IF OBJECT_ID(N'dbo.ImportTemplate', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ImportTemplate (
        Id              INT IDENTITY(1,1) NOT NULL
                        CONSTRAINT PK_ImportTemplate PRIMARY KEY,
        TemplateCode    NVARCHAR(20)   NOT NULL,   -- 'NVL01', 'NVL02', ..., 'TP'
        TemplateName    NVARCHAR(100)  NOT NULL,   -- 'Vật tư tấm', 'Chỉ cạnh / Nẹp'...
        ItemGroupCode   NVARCHAR(20)   NOT NULL,   -- mã nhóm VT mặc định
        CodePattern     NVARCHAR(50)   NULL,        -- 'NVL01_XXXXXX' (gợi ý pattern mã)
        NameGenAuto     BIT            NOT NULL
                        CONSTRAINT DF_ImportTemplate_NameGenAuto DEFAULT (0),
        IsActive        BIT            NOT NULL
                        CONSTRAINT DF_ImportTemplate_IsActive DEFAULT (1),

        CONSTRAINT UQ_ImportTemplate_Code UNIQUE (TemplateCode)
    );
END
GO

-- ── ImportTemplateDetail (per-column config) ───────────────────
IF OBJECT_ID(N'dbo.ImportTemplateDetail', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ImportTemplateDetail (
        Id              INT IDENTITY(1,1) NOT NULL
                        CONSTRAINT PK_ImportTemplateDetail PRIMARY KEY,
        TemplateId      INT            NOT NULL,   -- FK → ImportTemplate.Id
        ColIndex        INT            NOT NULL,   -- thứ tự cột trong file Excel (1-based)
        ColGroup        NVARCHAR(60)   NOT NULL,   -- 'NHÓM THÔNG TIN SẢN PHẨM' | 'TIỀN TỐ'
        ColName         NVARCHAR(100)  NOT NULL,   -- tên cột hiển thị trong Excel
        ColKey          NVARCHAR(100)  NOT NULL,   -- tên cột đích trong B20Item
        DataType        NVARCHAR(20)   NOT NULL
                        CONSTRAINT DF_ImportTemplateDetail_DataType DEFAULT (N'NVARCHAR'),
        IsNamePart      BIT            NOT NULL
                        CONSTRAINT DF_ImportTemplateDetail_IsNamePart DEFAULT (0),
        NameConcatOrder INT            NULL,       -- thứ tự ghép tên (NULL nếu IsNamePart=0)

        CONSTRAINT FK_ImportTemplateDetail_Template
            FOREIGN KEY (TemplateId) REFERENCES dbo.ImportTemplate (Id)
            ON DELETE CASCADE,
        CONSTRAINT UQ_ImportTemplateDetail_TemplateCol
            UNIQUE (TemplateId, ColIndex)
    );

    CREATE NONCLUSTERED INDEX IX_ImportTemplateDetail_TemplateId
        ON dbo.ImportTemplateDetail (TemplateId, ColIndex);
END
GO

-- ── Cấp quyền cho login [bom] ──────────────────────────────────
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'bom')
BEGIN
    CREATE USER [bom] FOR LOGIN [bom];
END
GO

GRANT SELECT ON dbo.ImportTemplate       TO [bom];
GRANT SELECT ON dbo.ImportTemplateDetail TO [bom];
GO
