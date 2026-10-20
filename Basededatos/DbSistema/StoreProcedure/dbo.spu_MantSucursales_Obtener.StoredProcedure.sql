-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene una sucursal para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantSucursales_Obtener
    @IdEmpresa CHAR(2),
    @IdSucursal CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT S.IdSucursal, S.GlsAbrev
        FROM dbo.Sucursales AS S
        WHERE S.IdEmpresa = @IdEmpresa AND S.IdSucursal = @IdSucursal;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
