-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Obtiene la cabecera y detalle de una venta.
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
            v.Serie,
            v.Numero,
            v.IdMoneda,
            m.CodigoMoneda,
            v.TipoCambio,
            v.BaseImponible,
            v.Igv,
            v.Isc,
            v.OtrosTributos,
            v.Redondeo,
            v.ImporteTotal,
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
        WHERE v.IdVenta = @IdVenta;

        SELECT
            d.IdVentaDetalle,
            d.IdVenta,
            d.Item,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto
        FROM dbo.VEN_VentaDetalle AS d
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
