-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista ubicaciones de un almacen.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantAlmacenes_Ubicaciones_Listar
    @IdEmpresa CHAR(2),
    @IdAlmacen CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT ROW_NUMBER() OVER (ORDER BY AU.IdUbicacion) AS Item, AU.IdUbicacion, AU.GlsUbicacion, AU.IdAlmacen, AU.IdEmpresa
        FROM dbo.AlmacenesUbicacion AS AU
        WHERE AU.IdEmpresa = @IdEmpresa AND AU.IdAlmacen = @IdAlmacen
        ORDER BY AU.IdUbicacion;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
