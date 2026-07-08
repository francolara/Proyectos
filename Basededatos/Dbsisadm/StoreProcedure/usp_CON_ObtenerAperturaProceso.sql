-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Obtiene el proceso de asiento de apertura y su detalle por ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Permite consultar desde Proceso si un anio de apertura ya fue generado y cuales fueron sus lineas resumen y analisis.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerAperturaProceso
    @IdEmpresa INT,
    @AnioApertura SMALLINT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.IdAperturaProceso,
            p.IdEmpresa,
            p.AnioApertura,
            p.AnioSaldo,
            p.MesSaldoHasta,
            p.PeriodoSaldoHasta,
            p.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            p.FechaAsiento,
            p.UsaTipoCambioSbs,
            p.TipoCambioCompra,
            p.TipoCambioVenta,
            p.IdAsiento,
            p.NumeroAsiento,
            p.TotalLineas,
            p.TotalDebe,
            p.TotalHaber,
            p.FechaRegistro,
            p.UsuarioRegistro
        FROM dbo.CON_AperturaProceso AS p
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = p.IdOrigen
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.AnioApertura = @AnioApertura;

        SELECT
            d.IdAperturaProcesoDetalle,
            d.IdAperturaProceso,
            d.Item,
            d.TipoDetalle,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.CodigoMoneda,
            d.TipoCambioAplicado,
            d.TipoDocumento,
            d.Serie,
            d.NumeroDocumento,
            d.Debe,
            d.Haber,
            d.TotalImporteS,
            d.TotalImporteD,
            d.Observacion,
            d.FechaRegistro,
            d.UsuarioRegistro
        FROM dbo.CON_AperturaProcesoDetalle AS d
        INNER JOIN dbo.CON_AperturaProceso AS p
            ON p.IdAperturaProceso = d.IdAperturaProceso
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.AnioApertura = @AnioApertura
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
