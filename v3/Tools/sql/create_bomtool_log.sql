/* ============================================================
   [BOMTool] — Database + bảng lưu lịch sử import của tool BOM2BRAVO
   Server   : cùng server với Bravo, nhưng TÁCH RIÊNG database
              (DB Bravo do Bravo quản lý, khách hàng không được
              tạo thêm bảng trong đó → tool dùng 1 DB riêng để log).
   Chạy 1 lần bằng tài khoản có quyền CREATE DATABASE + ALTER ANY LOGIN
   (hoặc nhờ DBA/IT của khách hàng chạy).
   Tool tự phát hiện DB + bảng: chưa có thì bỏ qua êm, có rồi thì
   tự ghi log — KHÔNG cần build lại tool.
   ============================================================ */

-- 1. Tạo database riêng cho log (nếu chưa có)
IF DB_ID(N'[BOMTool]') IS NULL
BEGIN
    CREATE DATABASE [BOMTool];
END
GO

USE [BOMTool];
GO

-- 2. Tạo bảng log
IF OBJECT_ID(N'dbo.BOMTool_ImportLog', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BOMTool_ImportLog (
        Id          INT IDENTITY(1,1) NOT NULL
                    CONSTRAINT PK_BOMTool_ImportLog PRIMARY KEY,
        LogTime     DATETIME       NOT NULL
                    CONSTRAINT DF_BOMTool_ImportLog_LogTime DEFAULT (GETDATE()),
        Computer    NVARCHAR(100)  NULL,        -- tên máy chạy tool
        LoginUser   NVARCHAR(100)  NULL,        -- user Windows
        Creator     NVARCHAR(100)  NULL,        -- mã Người lập chọn trong tool
        Action      NVARCHAR(20)   NOT NULL,    -- 'BOM' | 'THDM' | 'UNDO'
        FileName    NVARCHAR(255)  NULL,        -- file Excel nguồn
        RefId       NVARCHAR(50)   NULL,        -- BOM Id / Đơn hàng Id
        TotalRows   INT            NULL,        -- số dòng detail đã insert
        Status      NVARCHAR(10)   NOT NULL,    -- 'OK' | 'LOI'
        Detail      NVARCHAR(MAX)  NULL,        -- mô tả / nội dung lỗi
        AppVersion  NVARCHAR(20)   NULL         -- dự phòng: version tool
    );

    CREATE NONCLUSTERED INDEX IX_BOMTool_ImportLog_LogTime
        ON dbo.BOMTool_ImportLog (LogTime DESC);
END
GO

-- 3. Map login [bom] (login server dùng để connect Bravo) vào DB log này + cấp quyền.
--    Login [bom] đã tồn tại ở cấp server (đang dùng để connect DB Bravo),
--    ở đây chỉ tạo thêm user tương ứng bên trong [BOMTool] rồi cấp quyền.
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'bom')
BEGIN
    CREATE USER [bom] FOR LOGIN [bom];
END
GO

GRANT SELECT, INSERT ON dbo.BOMTool_ImportLog TO [bom];
GO
