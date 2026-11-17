-- Firma: FRANCO LARA | Create date: 17/11/2026 | Determina si una venta tiene efecto de facturación, independiente de factoring y detracción.
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_ComprobanteHYT_FacturaVigente(@ComprobanteId BIGINT)
RETURNS TABLE
AS
RETURN
(
    SELECT documento.ComprobanteId,documento.OrdenCompraId,documento.Tipo,documento.Serie,documento.Numero,
        documento.Fecha,documento.FechaRegistro,documento.UsuarioRegistro
    FROM dbo.ComprobanteHYT AS documento
    CROSS APPLY dbo.fn_ComprobanteHYT_SaldoVenta(documento.ComprobanteId) AS saldo
    WHERE documento.ComprobanteId=@ComprobanteId AND documento.EsVenta=1 AND documento.Tipo IN ('01','03')
        AND documento.Serie IS NOT NULL AND documento.Numero IS NOT NULL
        AND documento.EstadoSunat NOT IN (N'RECHAZADO',N'ANULADO',N'BAJA')
        AND (saldo.ImporteAplicado=0 OR saldo.ImporteAplicado<documento.Importe)
);
GO
