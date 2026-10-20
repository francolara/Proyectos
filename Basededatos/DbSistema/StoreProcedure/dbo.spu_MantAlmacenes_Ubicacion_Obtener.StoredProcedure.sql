-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene una ubicacion de almacen.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantAlmacenes_Ubicacion_Obtener
    @IdEmpresa CHAR(2),
    @IdAlmacen CHAR(8),
    @IdUbicacion VARCHAR(45)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT AU.IdUbicacion, AU.GlsUbicacion, AU.IdAlmacen, AU.IdEmpresa
        FROM dbo.AlmacenesUbicacion AS AU
        WHERE AU.IdEmpresa = @IdEmpresa AND AU.IdAlmacen = @IdAlmacen AND AU.IdUbicacion = @IdUbicacion;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
