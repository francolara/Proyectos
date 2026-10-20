-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene el precio principal de un producto y unidad de medida.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProductos_PrecioPrincipal_Obtener
    @IdEmpresa CHAR(2),
    @IdLista CHAR(8),
    @IdProducto CHAR(8),
    @IdUM CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PV.VVUnit,
            PV.IGVUnit,
            PV.PVUnit,
            PV.MaxDcto
        FROM dbo.PreciosVenta AS PV
        WHERE PV.IdEmpresa = @IdEmpresa
          AND PV.IdLista = @IdLista
          AND PV.IdProducto = @IdProducto
          AND PV.IdUM = @IdUM;
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
