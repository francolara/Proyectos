-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene la cabecera de un almacen.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantAlmacenes_Obtener
    @IdEmpresa CHAR(2),
    @IdAlmacen CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT A.IdAlmacen, A.GlsAlmacen, A.IdSucursal
        FROM dbo.Almacenes AS A
        WHERE A.IdEmpresa = @IdEmpresa AND A.IdAlmacen = @IdAlmacen;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
