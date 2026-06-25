-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de una provision de compra.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Devuelve subtotal, totales exonerado/inafecto, cuenta contable y afectacion IGV del detalle de compra.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Incluye saldo, descripcion del comprobante y numero de documento de la persona para la edicion y ayudas del modulo.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_COM_ObtenerCompra
    @IdCompra INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdCompra,
            c.IdEmpresa,
            c.IdProveedor,
            p.CodigoProveedor,
            pe.TipoDocumento,
            pe.NumeroDocumento,
            pe.NombreCompleto AS NombreProveedor,
            c.IdConfiguracionContabilizacion,
            cfg.ModuloOperacion,
            cfg.EscenarioOperacion,
            cfg.Descripcion AS DescripcionConfiguracion,
            c.IdAsiento,
            c.FechaEmision,
            c.FechaContabilizacion,
            c.TipoComprobante,
            tc.Descripcion AS DescripcionTipoComprobante,
            c.Serie,
            c.Numero,
            pe.NumeroDocumento AS NumeroDocumentoPersona,
            c.IdMoneda,
            m.CodigoMoneda,
            c.TipoCambio,
            c.BaseImponible,
            c.TotalExonerado,
            c.TotalInafecto,
            c.Icbper,
            c.Igv,
            c.Isc,
            c.OtrosTributos,
            c.Redondeo,
            c.ImporteTotal,
            c.Saldo,
            c.Observacion,
            c.Estado
        FROM dbo.COM_Compra AS c
        INNER JOIN dbo.ADM_Proveedor AS p
            ON p.IdProveedor = c.IdProveedor
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = p.IdPersona
        INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
            ON cfg.IdConfiguracionContabilizacion = c.IdConfiguracionContabilizacion
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = c.IdMoneda
        INNER JOIN dbo.ADM_TipoComprobante AS tc
            ON tc.CodigoTipoComprobante = c.TipoComprobante
        WHERE c.IdCompra = @IdCompra;

        SELECT
            d.IdCompraDetalle,
            d.IdCompra,
            d.Item,
            d.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            d.IdTipoAfectacionIGV,
            a.CodigoSunat AS CodigoAfectacionIGV,
            a.NombreAfectacion AS NombreAfectacionIGV,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto
        FROM dbo.COM_CompraDetalle AS d
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        INNER JOIN dbo.CON_TipoAfectacionIGV AS a
            ON a.IdTipoAfectacionIGV = d.IdTipoAfectacionIGV
        WHERE d.IdCompra = @IdCompra
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
