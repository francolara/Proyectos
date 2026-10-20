-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene la cabecera de un proveedor para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedores_Obtener
    @IdEmpresa CHAR(2),
    @IdProveedor CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PR.IdProveedor,
            PR.GlsObservacion,
            PR.IndRetencion,
            PR.IndPercepcion,
            PR.IndBnContribuyente
        FROM dbo.Proveedores AS PR
        WHERE PR.IdEmpresa = @IdEmpresa
          AND PR.IdProveedor = @IdProveedor;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
