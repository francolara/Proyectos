-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un grupo de productos para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantGruposProductos_Obtener
    @IdEmpresa CHAR(2),
    @IdGrupo CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT GP.IdGrupo, GP.GlsGrupo
        FROM dbo.GruposProducto AS GP
        WHERE GP.IdEmpresa = @IdEmpresa AND GP.IdGrupo = @IdGrupo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
