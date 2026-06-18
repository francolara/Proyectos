-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de una configuracion contable automatica.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerConfiguracionContabilizacion
    @IdConfiguracionContabilizacion INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdConfiguracionContabilizacion,
            c.IdEmpresa,
            c.ModuloOperacion,
            c.EscenarioOperacion,
            c.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            c.Descripcion,
            c.GeneraAsientoAutomatico,
            c.UsaTipoCambio,
            c.Activo
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = c.IdOrigen
        WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;

        SELECT
            d.IdConfiguracionContabilizacionDetalle,
            d.IdConfiguracionContabilizacion,
            d.Orden,
            d.ComponenteContable,
            d.IdPlanCuenta,
            p.CodigoCuenta,
            p.NombreCuenta,
            d.NaturalezaMovimiento,
            d.Activo
        FROM dbo.CON_ConfiguracionContabilizacionDetalle AS d
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        WHERE d.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
        ORDER BY
            d.Orden ASC;

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
