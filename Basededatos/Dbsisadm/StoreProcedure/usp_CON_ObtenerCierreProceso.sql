-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Obtiene el proceso anual de asiento de cierre y su detalle por cuenta.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Permite consultar desde el modulo Proceso si un ejercicio ya fue generado por el asiento de cierre y que cuentas participaron en 14 y 15.
-- Firma: FRANCO LARA - 13/08/2026 | Expone el periodo de corte, periodo de generacion, asiento unico y las lineas contables con sus importes en soles y dolares.
-- Firma: FRANCO LARA - 22/08/2026 | Identifica el periodo 14 como cierre unico de Inventario.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerCierreProceso
    @IdEmpresa INT,
    @Anio SMALLINT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.IdCierreProceso,
            p.IdEmpresa,
            p.Anio,
            p.MesSaldoHasta,
            p.PeriodoSaldoHasta,
            p.MesGeneracion,
            p.PeriodoGeneracion,
            p.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            p.FechaAsiento,
            p.UsaTipoCambioSbs,
            p.TipoCambioCompra,
            p.TipoCambioVenta,
            p.ProcesaGananciasPerdidas,
            p.ProcesaInventarios,
            p.IdAsiento,
            p.NumeroAsiento,
            p.TotalLineas,
            p.TotalCuentas,
            p.TotalAsientos,
            p.TotalDebe,
            p.TotalHaber,
            p.FechaRegistro,
            p.UsuarioRegistro
        FROM dbo.CON_CierreProceso AS p
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = p.IdOrigen
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Anio = @Anio;

        SELECT
            d.IdCierreProcesoDetalle,
            d.IdCierreProceso,
            d.Item,
            d.TipoCierre,
            N'Cierre de Inventario' AS DescripcionCierre,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.CodigoMoneda,
            d.TipoCambioAplicado,
            d.IdAsiento,
            d.NumeroAsiento,
            d.DH,
            d.TotalDebe,
            d.TotalHaber,
            d.TotalImporteS,
            d.TotalImporteD,
            d.Estado,
            d.Observacion,
            d.FechaRegistro,
            d.UsuarioRegistro
        FROM dbo.CON_CierreProcesoDetalle AS d
        INNER JOIN dbo.CON_CierreProceso AS p
            ON p.IdCierreProceso = d.IdCierreProceso
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Anio = @Anio
        ORDER BY
            d.Item ASC;

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
