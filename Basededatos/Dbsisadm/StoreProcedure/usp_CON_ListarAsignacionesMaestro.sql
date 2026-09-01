-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Lista parametros, impuestos y documentos cuyas cuentas se asignan desde el plan maestro.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   31/08/2026
-- Description:   Lista todos los parametros de tipo CONTABLE sin depender de un catalogo fijo de codigos.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarAsignacionesMaestro
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            parametro.IdParametroMaestro,
            parametro.TipoParametro,
            parametro.CodigoParametro,
            parametro.DescripcionParametro,
            NULLIF(LTRIM(RTRIM(parametro.ValorParametro)), N'') AS CodigoCuenta,
            cuenta.NombreCuenta,
            parametro.Activo
        FROM dbo.ADM_ParametroMaestro AS parametro
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
            ON cuenta.CodigoCuenta = parametro.ValorParametro
        WHERE parametro.TipoParametro = 'CONTABLE'
        ORDER BY parametro.Orden, parametro.CodigoParametro;

        SELECT
            impuesto.IdTipoImpuesto,
            impuesto.CodigoSunat,
            impuesto.NombreImpuesto,
            impuesto.CodigoCuenta,
            cuenta.NombreCuenta,
            impuesto.Estado
        FROM dbo.CON_TipoImpuesto AS impuesto
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
            ON cuenta.CodigoCuenta = impuesto.CodigoCuenta
        ORDER BY impuesto.CodigoSunat;

        SELECT
            tipo.IdTipoComprobante,
            tipo.CodigoTipoComprobante,
            tipo.Descripcion,
            tipo.UsoCompras,
            tipo.UsoVentas,
            tipo.CodigoCuentaVentaSoles,
            ventaSoles.NombreCuenta AS NombreCuentaVentaSoles,
            tipo.CodigoCuentaVentaDolares,
            ventaDolares.NombreCuenta AS NombreCuentaVentaDolares,
            tipo.CodigoCuentaCompraSoles,
            compraSoles.NombreCuenta AS NombreCuentaCompraSoles,
            tipo.CodigoCuentaCompraDolares,
            compraDolares.NombreCuenta AS NombreCuentaCompraDolares,
            tipo.Estado
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS ventaSoles ON ventaSoles.CodigoCuenta = tipo.CodigoCuentaVentaSoles
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS ventaDolares ON ventaDolares.CodigoCuenta = tipo.CodigoCuentaVentaDolares
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS compraSoles ON compraSoles.CodigoCuenta = tipo.CodigoCuentaCompraSoles
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS compraDolares ON compraDolares.CodigoCuenta = tipo.CodigoCuentaCompraDolares
        ORDER BY tipo.CodigoTipoComprobante;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
