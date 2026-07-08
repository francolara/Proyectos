-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de un asiento contable.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Devuelve datos documentarios opcionales del detalle del asiento.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Expone TotalImporteS y TotalImporteD del detalle para conservar equivalencias por moneda al editar.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Incluye fecha de emision en la cabecera para asientos manuales y automaticos.
-- =============================================
-- Firma: FRANCO LARA - 03/07/2026 | Devuelve DH en el detalle para exponer el sentido contable explicito por linea.
-- Firma: FRANCO LARA - 06/07/2026 | Expone tambien TotalImporteS y TotalImporteD consolidados en la cabecera del asiento para reutilizar el resumen del listado y formulario.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerAsiento
    @IdAsiento INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            a.IdAsiento,
            a.IdEmpresa,
            a.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            o.PermiteRegistroManual,
            a.Ejercicio,
            a.Mes,
            a.Periodo,
            a.NumeroAsiento,
            a.FechaEmision,
            a.FechaAsiento,
            a.Glosa,
            a.IdMoneda,
            m.CodigoMoneda,
            m.NombreMoneda,
            m.SimboloMoneda,
            a.TipoCambio,
            a.TotalDebe,
            a.TotalHaber,
            CASE
                WHEN ISNULL(dt.TotalImporteDebeS, 0) >= ISNULL(dt.TotalImporteHaberS, 0) THEN ISNULL(dt.TotalImporteDebeS, 0)
                ELSE ISNULL(dt.TotalImporteHaberS, 0)
            END AS TotalImporteS,
            CASE
                WHEN ISNULL(dt.TotalImporteDebeD, 0) >= ISNULL(dt.TotalImporteHaberD, 0) THEN ISNULL(dt.TotalImporteDebeD, 0)
                ELSE ISNULL(dt.TotalImporteHaberD, 0)
            END AS TotalImporteD,
            a.Estado,
            a.ReferenciaExterna,
            a.Observacion
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = a.IdOrigen
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = a.IdMoneda
        OUTER APPLY
        (
            SELECT
                TotalImporteDebeS = SUM(CASE WHEN d.DH = 'D' THEN ISNULL(d.TotalImporteS, 0) ELSE 0 END),
                TotalImporteHaberS = SUM(CASE WHEN d.DH = 'H' THEN ISNULL(d.TotalImporteS, 0) ELSE 0 END),
                TotalImporteDebeD = SUM(CASE WHEN d.DH = 'D' THEN ISNULL(d.TotalImporteD, 0) ELSE 0 END),
                TotalImporteHaberD = SUM(CASE WHEN d.DH = 'H' THEN ISNULL(d.TotalImporteD, 0) ELSE 0 END)
            FROM dbo.CON_AsientoDetalle AS d
            WHERE d.IdAsiento = a.IdAsiento
        ) AS dt
        WHERE a.IdAsiento = @IdAsiento;

        SELECT
            d.IdAsientoDetalle,
            d.IdAsiento,
            d.Item,
            d.IdPlanCuenta,
            d.DH,
            p.CodigoCuenta,
            p.NombreCuenta,
            d.GlosaDetalle,
            d.CodigoCentroCosto,
            d.TipoDocumento,
            d.NumeroDocumento,
            d.Serie,
            d.TipoCambioLinea,
            d.Debe,
            d.Haber,
            d.TotalImporteS,
            d.TotalImporteD,
            d.ReferenciaLinea
        FROM dbo.CON_AsientoDetalle AS d
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        WHERE d.IdAsiento = @IdAsiento
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
