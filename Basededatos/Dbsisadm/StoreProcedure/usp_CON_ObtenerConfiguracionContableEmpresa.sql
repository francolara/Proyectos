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
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Incluye el origen de provision para detracciones y la cuenta SPOT dentro de impuestos.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Expone tambien las configuraciones DIF, AJU, APR y CIE para seleccionar los origenes de diferencia en cambio, ajuste de cuentas, apertura y cierre desde configuracion contable.
-- Firma: FRANCO LARA - 25/08/2026 | Expone exclusivamente las cuentas configuradas por empresa; el maestro solo se utiliza durante la carga inicial.

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
          AND c.ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC', 'DET', 'DIF', 'AJU', 'APR', 'CIE');

        SELECT
            t.IdTipoComprobante,
            t.CodigoTipoComprobante,
            t.Descripcion,
            t.UsoCompras,
            t.UsoVentas,
            cfg.IdDocumentoConfiguracionEmpresa,
            cfg.IdCuentaVentaSoles,
            pvs.CodigoCuenta AS CodigoCuentaVentaSoles,
            pvs.NombreCuenta AS NombreCuentaVentaSoles,
            cfg.IdCuentaVentaDolares,
            pvd.CodigoCuenta AS CodigoCuentaVentaDolares,
            pvd.NombreCuenta AS NombreCuentaVentaDolares,
            cfg.IdCuentaCompraSoles,
            pcs.CodigoCuenta AS CodigoCuentaCompraSoles,
            pcs.NombreCuenta AS NombreCuentaCompraSoles,
            cfg.IdCuentaCompraDolares,
            pcd.CodigoCuenta AS CodigoCuentaCompraDolares,
            pcd.NombreCuenta AS NombreCuentaCompraDolares,
            ISNULL(cfg.Activo, CONVERT(BIT, 1)) AS Activo
        FROM dbo.ADM_TipoComprobante AS t
        LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoComprobante = t.IdTipoComprobante
           AND cfg.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.CON_PlanCuenta AS pvs
            ON pvs.IdPlanCuenta = cfg.IdCuentaVentaSoles
        LEFT JOIN dbo.CON_PlanCuenta AS pvd
            ON pvd.IdPlanCuenta = cfg.IdCuentaVentaDolares
        LEFT JOIN dbo.CON_PlanCuenta AS pcs
            ON pcs.IdPlanCuenta = cfg.IdCuentaCompraSoles
        LEFT JOIN dbo.CON_PlanCuenta AS pcd
            ON pcd.IdPlanCuenta = cfg.IdCuentaCompraDolares
        ORDER BY
            t.CodigoTipoComprobante ASC;

        SELECT
            i.IdTipoImpuesto,
            i.CodigoSunat,
            i.NombreImpuesto,
            cfg.IdTipoImpuestoConfiguracionEmpresa AS IdConfiguracion,
            cfg.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            ISNULL(cfg.Activo, CONVERT(BIT, 1)) AS Activo
        FROM dbo.CON_TipoImpuesto AS i
        LEFT JOIN dbo.CON_TipoImpuestoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoImpuesto = i.IdTipoImpuesto
           AND cfg.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = cfg.IdPlanCuenta
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
