-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de una provision de compra.
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
            c.Serie,
            c.Numero,
            c.IdMoneda,
            m.CodigoMoneda,
            c.TipoCambio,
            c.BaseImponible,
            c.Igv,
            c.Isc,
            c.OtrosTributos,
            c.Redondeo,
            c.ImporteTotal,
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
        WHERE c.IdCompra = @IdCompra;

        SELECT
            d.IdCompraDetalle,
            d.IdCompra,
            d.Item,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto
        FROM dbo.COM_CompraDetalle AS d
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
