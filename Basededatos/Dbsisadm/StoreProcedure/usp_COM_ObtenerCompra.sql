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
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Devuelve la detraccion configurada en la compra y su documento SPOT independiente.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Devuelve la ultima fecha, estado y mensaje de validacion CPE de la compra, ademas de la percepcion configurada y su documento pendiente.
-- =============================================
-- Firma: FRANCO LARA - 30/06/2026 | Devuelve la retencion de renta de 4ta en cabecera de compras, expone el IdCompraRetencion vinculado y permite recuperar detalles importados sin cuenta contable asignada.

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
            c.ExoneracionRenta4ta,
            c.PorcentajeRetencion,
            c.Retencion,
            cr.IdCompraRetencion,
            cd.IdCompraDetraccion,
            cd.IdAsiento AS IdAsientoDetraccion,
            cp.IdCompraPercepcion,
            cp.IdAsiento AS IdAsientoPercepcion,
            c.TieneDetraccion,
            c.IdDetraccionSunat,
            ISNULL(d.CodigoSunat, '') AS CodigoDetraccionSunat,
            ISNULL(d.Descripcion, '') AS DescripcionDetraccionSunat,
            c.PorcentajeDetraccion,
            c.ImporteDetraccion,
            c.TienePercepcion,
            c.IdTipoPercepcion,
            ISNULL(tp.Codigo, '') AS CodigoPercepcion,
            ISNULL(tp.Descripcion, '') AS DescripcionPercepcion,
            c.PorcentajePercepcion,
            c.BasePercepcion,
            c.ImportePercepcion,
            c.Observacion,
            c.FechaValidacionCpe,
            c.EstadoValidacionCpe,
            c.MensajeValidacionCpe,
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
        LEFT JOIN dbo.ADM_DetraccionSunat AS d
            ON d.IdDetraccionSunat = c.IdDetraccionSunat
        LEFT JOIN dbo.COM_CompraRetencion AS cr
            ON cr.IdCompra = c.IdCompra
        LEFT JOIN dbo.COM_CompraDetraccion AS cd
            ON cd.IdCompra = c.IdCompra
        LEFT JOIN dbo.ADM_TipoPercepcion AS tp
            ON tp.IdTipoPercepcion = c.IdTipoPercepcion
        LEFT JOIN dbo.COM_CompraPercepcion AS cp
            ON cp.IdCompra = c.IdCompra
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
        LEFT JOIN dbo.CON_PlanCuenta AS pc
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
