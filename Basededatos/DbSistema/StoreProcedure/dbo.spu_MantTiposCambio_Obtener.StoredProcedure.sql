-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un tipo de cambio para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantTiposCambio_Obtener
    @Fecha DATE
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            TC.Fecha,
            TC.TCFacturacion,
            TC.TCVenta,
            TC.TCCompra,
            TC.TCCompraC,
            TC.TCVentaSBS,
            TC.TCCompraSBS
        FROM dbo.TiposDeCambio AS TC
        WHERE TC.Fecha = @Fecha;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
