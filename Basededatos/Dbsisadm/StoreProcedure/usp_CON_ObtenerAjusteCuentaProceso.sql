-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Obtiene la cabecera y detalle del proceso de ajuste de cuentas por empresa y periodo.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Permite consultar desde el modulo Proceso si un periodo ya fue generado por AJU y que cuentas analiticas produjeron asiento, mostrando el tipo de cambio real del asiento generado cuando ya existe.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerAjusteCuentaProceso
    @IdEmpresa INT,
    @Periodo CHAR(6)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.IdAjusteCuentaProceso,
            p.IdEmpresa,
            p.Periodo,
            p.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            p.FechaAsiento,
            p.TotalCuentas,
            p.TotalAsientos,
            p.TotalDebe,
            p.TotalHaber,
            p.FechaRegistro,
            p.UsuarioRegistro
        FROM dbo.CON_AjusteCuentaProceso AS p
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = p.IdOrigen
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Periodo = @Periodo;

        SELECT
            d.IdAjusteCuentaProcesoDetalle,
            d.IdAjusteCuentaProceso,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.CodigoMoneda,
            ISNULL(a.TipoCambio, d.TipoCambioAplicado) AS TipoCambioAplicado,
            d.TotalAnalisis,
            d.IdAsiento,
            d.NumeroAsiento,
            d.TotalDebe,
            d.TotalHaber,
            d.Estado,
            d.Observacion,
            d.FechaRegistro,
            d.UsuarioRegistro
        FROM dbo.CON_AjusteCuentaProcesoDetalle AS d
        INNER JOIN dbo.CON_AjusteCuentaProceso AS p
            ON p.IdAjusteCuentaProceso = d.IdAjusteCuentaProceso
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        LEFT JOIN dbo.CON_Asiento AS a
            ON a.IdAsiento = d.IdAsiento
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
