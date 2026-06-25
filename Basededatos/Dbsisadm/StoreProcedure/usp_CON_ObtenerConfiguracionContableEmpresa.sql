-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Obtiene configuracion contable por empresa para tabs de provision, documentos e impuestos.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Devuelve cuentas de documento separadas y configuracion unica de impuestos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Incluye todas las provisiones configurables de compras, ventas, egresos, ingresos y aplicaciones NC.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerConfiguracionContableEmpresa
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdConfiguracionContabilizacion,
            c.ModuloOperacion,
            c.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            c.GeneraAsientoAutomatico,
            c.Activo
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = c.IdOrigen
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.EscenarioOperacion = 'PROVISION'
          AND c.ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC');

        SELECT
            t.IdTipoComprobante,
            t.CodigoTipoComprobante,
            t.Descripcion,
            t.UsoCompras,
            t.UsoVentas,
            cfg.IdDocumentoConfiguracionEmpresa,
            COALESCE(cfg.IdCuentaVentaSoles, t.IdCuentaVentaSoles) AS IdCuentaVentaSoles,
            pvs.CodigoCuenta AS CodigoCuentaVentaSoles,
            pvs.NombreCuenta AS NombreCuentaVentaSoles,
            COALESCE(cfg.IdCuentaVentaDolares, t.IdCuentaVentaDolares) AS IdCuentaVentaDolares,
            pvd.CodigoCuenta AS CodigoCuentaVentaDolares,
            pvd.NombreCuenta AS NombreCuentaVentaDolares,
            COALESCE(cfg.IdCuentaCompraSoles, t.IdCuentaCompraSoles) AS IdCuentaCompraSoles,
            pcs.CodigoCuenta AS CodigoCuentaCompraSoles,
            pcs.NombreCuenta AS NombreCuentaCompraSoles,
            COALESCE(cfg.IdCuentaCompraDolares, t.IdCuentaCompraDolares) AS IdCuentaCompraDolares,
            pcd.CodigoCuenta AS CodigoCuentaCompraDolares,
            pcd.NombreCuenta AS NombreCuentaCompraDolares,
            ISNULL(cfg.Activo, CONVERT(BIT, 1)) AS Activo
        FROM dbo.ADM_TipoComprobante AS t
        LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoComprobante = t.IdTipoComprobante
           AND cfg.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.CON_PlanCuenta AS pvs
            ON pvs.IdPlanCuenta = COALESCE(cfg.IdCuentaVentaSoles, t.IdCuentaVentaSoles)
        LEFT JOIN dbo.CON_PlanCuenta AS pvd
            ON pvd.IdPlanCuenta = COALESCE(cfg.IdCuentaVentaDolares, t.IdCuentaVentaDolares)
        LEFT JOIN dbo.CON_PlanCuenta AS pcs
            ON pcs.IdPlanCuenta = COALESCE(cfg.IdCuentaCompraSoles, t.IdCuentaCompraSoles)
        LEFT JOIN dbo.CON_PlanCuenta AS pcd
            ON pcd.IdPlanCuenta = COALESCE(cfg.IdCuentaCompraDolares, t.IdCuentaCompraDolares)
        ORDER BY
            t.CodigoTipoComprobante ASC;

        SELECT
            i.IdTipoImpuesto,
            i.CodigoSunat,
            i.NombreImpuesto,
            cfg.IdTipoImpuestoConfiguracionEmpresa AS IdConfiguracion,
            COALESCE(cfg.IdPlanCuenta, i.IdPlanCuenta) AS IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            ISNULL(cfg.Activo, CONVERT(BIT, 1)) AS Activo
        FROM dbo.CON_TipoImpuesto AS i
        LEFT JOIN dbo.CON_TipoImpuestoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoImpuesto = i.IdTipoImpuesto
           AND cfg.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = COALESCE(cfg.IdPlanCuenta, i.IdPlanCuenta)
        WHERE i.Estado = 1
        ORDER BY
            i.IdTipoImpuesto ASC;

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
