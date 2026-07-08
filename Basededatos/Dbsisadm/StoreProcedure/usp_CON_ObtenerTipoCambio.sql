-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Obtiene un tipo de cambio especifico por cuenta administradora.
-- =============================================
-- Firma: FRANCO LARA - 29/06/2026 | Devuelve el registro puntual del mantenimiento de tipos de cambio.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerTipoCambio
    @IdTipoCambio INT,
    @IdCuentaAdministradora INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            tc.IdTipoCambio,
            tc.IdCuentaAdministradora,
            tc.Fecha,
            tc.IdMoneda,
            tc.Compra,
            tc.Venta,
            tc.CompraSBS,
            tc.VentaSBS,
            tc.Fuente,
            tc.UsuarioRegistro,
            tc.Estado
        FROM dbo.CON_TipoCambio AS tc
        WHERE tc.IdTipoCambio = @IdTipoCambio
          AND tc.IdCuentaAdministradora = @IdCuentaAdministradora;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
