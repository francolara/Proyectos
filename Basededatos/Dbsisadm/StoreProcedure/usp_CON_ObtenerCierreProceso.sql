-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Obtiene el proceso anual de asiento de cierre y su detalle por cuenta.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Permite consultar desde el modulo Proceso si un ejercicio ya fue generado por el asiento de cierre y que cuentas participaron en 14 y 15.

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
            p.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            p.FechaAsiento,
            p.UsaTipoCambioSbs,
            p.TipoCambioCompra,
            p.TipoCambioVenta,
            p.ProcesaGananciasPerdidas,
            p.ProcesaInventarios,
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
            d.TipoCierre,
            CASE d.TipoCierre
                WHEN '14' THEN N'Cierre de Ganancias y Perdidas'
                ELSE N'Cierre de Inventarios'
            END AS DescripcionCierre,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.CodigoMoneda,
            d.TipoCambioAplicado,
            d.IdAsiento,
            d.NumeroAsiento,
            d.TotalDebe,
            d.TotalHaber,
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
            d.TipoCierre ASC,
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
