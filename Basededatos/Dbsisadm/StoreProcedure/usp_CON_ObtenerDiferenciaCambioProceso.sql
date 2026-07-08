-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/07/2026
-- Description:   Obtiene la cabecera y detalle del proceso de diferencia en cambio por empresa y periodo.
-- =============================================
-- Firma: FRANCO LARA - 01/07/2026 | Permite consultar desde el modulo Proceso si un periodo ya fue generado y que cuentas produjeron asientos de diferencia en cambio.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerDiferenciaCambioProceso
    @IdEmpresa INT,
    @Periodo CHAR(6)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.IdDiferenciaCambioProceso,
            p.IdEmpresa,
            p.Periodo,
            p.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            p.FechaAsiento,
            p.UsaTipoCambioSbs,
            p.TipoCambioCompra,
            p.TipoCambioVenta,
            p.TotalCuentas,
            p.TotalAsientos,
            p.TotalDebe,
            p.TotalHaber,
            p.FechaRegistro,
            p.UsuarioRegistro
        FROM dbo.CON_DiferenciaCambioProceso AS p
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = p.IdOrigen
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Periodo = @Periodo;

        SELECT
            d.IdDiferenciaCambioProcesoDetalle,
            d.IdDiferenciaCambioProceso,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.GeneraPorAnalisis,
            d.TipoCambioAplicado,
            d.IdAsiento,
            d.NumeroAsiento,
            d.TotalDebe,
            d.TotalHaber,
            d.Estado,
            d.Observacion,
            d.FechaRegistro,
            d.UsuarioRegistro
        FROM dbo.CON_DiferenciaCambioProcesoDetalle AS d
        INNER JOIN dbo.CON_DiferenciaCambioProceso AS p
            ON p.IdDiferenciaCambioProceso = d.IdDiferenciaCambioProceso
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Periodo = @Periodo
        ORDER BY
            pc.CodigoCuenta ASC;

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
