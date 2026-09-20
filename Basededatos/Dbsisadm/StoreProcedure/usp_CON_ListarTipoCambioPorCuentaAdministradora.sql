-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Lista tipos de cambio por cuenta administradora filtrando por periodo.
-- =============================================
-- Firma: FRANCO LARA - 29/06/2026 | Expone el mantenimiento de tipos de cambio por cuenta administradora para el periodo consultado.

-- Firma: FRANCO LARA - 19/09/2026 | Expone la trazabilidad de los tipos de cambio para su edicion.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarTipoCambioPorCuentaAdministradora
    @IdCuentaAdministradora INT,
    @Anio SMALLINT,
    @Mes TINYINT
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
            tc.FechaRegistro,
            tc.UsuarioActualizacion,
            tc.FechaActualizacion,
            tc.Estado
        FROM dbo.CON_TipoCambio AS tc
        WHERE tc.IdCuentaAdministradora = @IdCuentaAdministradora
          AND YEAR(tc.Fecha) = @Anio
          AND MONTH(tc.Fecha) = @Mes
        ORDER BY
            tc.Fecha DESC,
            tc.IdMoneda ASC;

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
