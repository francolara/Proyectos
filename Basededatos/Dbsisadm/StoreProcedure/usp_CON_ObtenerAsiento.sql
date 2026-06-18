-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de un asiento contable.
-- =============================================

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
            a.FechaAsiento,
            a.Glosa,
            a.IdMoneda,
            m.CodigoMoneda,
            m.NombreMoneda,
            m.SimboloMoneda,
            a.TipoCambio,
            a.TotalDebe,
            a.TotalHaber,
            a.Estado,
            a.ReferenciaExterna,
            a.Observacion
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = a.IdOrigen
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = a.IdMoneda
        WHERE a.IdAsiento = @IdAsiento;

        SELECT
            d.IdAsientoDetalle,
            d.IdAsiento,
            d.Item,
            d.IdPlanCuenta,
            p.CodigoCuenta,
            p.NombreCuenta,
            d.GlosaDetalle,
            d.Debe,
            d.Haber,
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
