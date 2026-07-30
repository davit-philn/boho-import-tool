/* ============================================================
   usp_BOMTool_DeleteBOM — Hoàn tác 1 lần import BOM của tool
   Database : B10_Boho_Data   (yêu cầu SQL Server 2016+ cho CREATE OR ALTER)
   Bắt buộc tạo trong đúng DB Bravo vì SP thao tác trực tiếp
   dbo.B20BOM / dbo.B20BOMDetail (bảng lõi Bravo, tên không kèm DB
   nên PHẢI resolve theo DB nơi SP được tạo).

   Xóa B20BOMDetail + B20BOM của đúng 1 BOM Id, trong 1 transaction.
   An toàn:
     - BOM Id không tồn tại  → báo lỗi, không xóa gì.
     - Nếu bảng log [BOMTool].dbo.BOMTool_ImportLog đã có: chỉ cho
       xóa BOM Id nào có trong log import của tool (chặn xóa nhầm BOM
       tạo tay trên BRAVO), và tự ghi thêm 1 dòng log Action='UNDO'.
       Bảng log nằm ở DB riêng [BOMTool] (xem create_[BOMTool].sql)
       — CHẠY create_[BOMTool].sql TRƯỚC script này, vì OBJECT_ID
       tham chiếu chéo DB sẽ lỗi nếu DB [BOMTool] chưa tồn tại.
       Login [bom] cần được cấp quyền SELECT/INSERT trong [BOMTool]
       (đã làm ở create_[BOMTool].sql).
   Chạy 1 lần bằng tài khoản có quyền CREATE PROCEDURE.
   ============================================================ */

CREATE OR ALTER PROCEDURE dbo.usp_BOMTool_DeleteBOM
    @BOMId INT
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.B20BOM WHERE Id = @BOMId)
    BEGIN
        RAISERROR(N'Không tìm thấy BOM Id %d trong B20BOM.', 16, 1, @BOMId);
        RETURN;
    END

    /* Chỉ cho xóa BOM do tool import (khi đã có bảng log ở DB [BOMTool]).
       Dùng dynamic SQL để SP tạo được cả khi bảng log chưa tồn tại. */
    IF OBJECT_ID(N'[BOMTool].dbo.BOMTool_ImportLog', N'U') IS NOT NULL
    BEGIN
        DECLARE @cnt INT;
        EXEC sp_executesql
            N'SELECT @c = COUNT(*) FROM [BOMTool].dbo.BOMTool_ImportLog
              WHERE Action = N''BOM'' AND Status = N''OK''
                AND RefId  = CAST(@id AS NVARCHAR(50))',
            N'@c INT OUTPUT, @id INT',
            @c = @cnt OUTPUT, @id = @BOMId;
        IF @cnt = 0
        BEGIN
            RAISERROR(N'BOM Id %d không có trong log import của tool — không cho phép xóa.',
                      16, 1, @BOMId);
            RETURN;
        END
    END

    DECLARE @nDet INT = 0, @nHdr INT = 0;

    BEGIN TRY
        BEGIN TRANSACTION;

        DELETE FROM dbo.B20BOMDetail WHERE BOMId = @BOMId;
        SET @nDet = @@ROWCOUNT;

        DELETE FROM dbo.B20BOM WHERE Id = @BOMId;
        SET @nHdr = @@ROWCOUNT;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH

    /* Ghi log UNDO (nếu có bảng log ở DB [BOMTool]) */
    IF OBJECT_ID(N'[BOMTool].dbo.BOMTool_ImportLog', N'U') IS NOT NULL
    BEGIN
        EXEC sp_executesql
            N'INSERT INTO [BOMTool].dbo.BOMTool_ImportLog
                (Computer, LoginUser, Action, RefId, TotalRows, Status, Detail)
              VALUES (HOST_NAME(), SUSER_SNAME(), N''UNDO'',
                      CAST(@id AS NVARCHAR(50)), @n, N''OK'',
                      N''Xóa qua usp_BOMTool_DeleteBOM'')',
            N'@id INT, @n INT', @id = @BOMId, @n = @nDet;
    END

    SELECT @nHdr AS HeaderDeleted, @nDet AS DetailDeleted;
END
GO

/* Cấp quyền chạy cho tài khoản tool (login: bom) */
GRANT EXECUTE ON dbo.usp_BOMTool_DeleteBOM TO [bom];
GO
