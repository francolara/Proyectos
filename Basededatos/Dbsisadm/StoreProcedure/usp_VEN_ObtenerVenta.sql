-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de una venta.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Incorpora subtotal, total exonerado, total inafecto e ICBPER interno, con cuenta contable y afectacion IGV por detalle.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Incluye saldo, descripcion del comprobante y numero de documento de la persona para la edicion y ayudas del modulo.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_VEN_ObtenerVenta
    @IdVenta INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            v.IdVenta,
            v.IdEmpresa,
            v.IdCliente,
            c.CodigoCliente,
            pe.TipoDocumento,
            pe.NumeroDocumento,
            pe.NombreCompleto AS NombreCliente,
            v.IdConfiguracionContabilizacion,
            cfg.ModuloOperacion,
            cfg.EscenarioOperacion,
            cfg.Descripcion AS DescripcionConfiguracion,
            v.IdAsiento,
            v.FechaEmision,
            v.FechaContabilizacion,
            v.TipoComprobante,
            tc.Descripcion AS DescripcionTipoComprobante,
            v.Serie,
            v.Numero,
            pe.NumeroDocumento AS NumeroDocumentoPersona,
            v.IdMoneda,
            m.CodigoMoneda,
            v.TipoCambio,
            v.BaseImponible,
            v.TotalExonerado,
            v.TotalInafecto,
            v.Icbper,
            v.Igv,
            v.Isc,
            v.OtrosTributos,
            v.Redondeo,
            v.ImporteTotal,
            v.Saldo,
            v.Observacion,
            v.Estado
        FROM dbo.VEN_Venta AS v
        INNER JOIN dbo.ADM_Cliente AS c
            ON c.IdCliente = v.IdCliente
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = c.IdPersona
        INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
            ON cfg.IdConfiguracionContabilizacion = v.IdConfiguracionContabilizacion
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = v.IdMoneda
        INNER JOIN dbo.ADM_TipoComprobante AS tc
            ON tc.CodigoTipoComprobante = v.TipoComprobante
        WHERE v.IdVenta = @IdVenta;

        SELECT
            d.IdVentaDetalle,
            d.IdVenta,
            d.Item,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.IdTipoAfectacionIGV,
            ta.CodigoSunat AS CodigoAfectacionIGV,
            ta.NombreAfectacion AS NombreAfectacionIGV,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto
        FROM dbo.VEN_VentaDetalle AS d
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        INNER JOIN dbo.CON_TipoAfectacionIGV AS ta
            ON ta.IdTipoAfectacionIGV = d.IdTipoAfectacionIGV
        WHERE d.IdVenta = @IdVenta
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
