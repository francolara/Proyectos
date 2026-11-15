-- Firma: FRANCO LARA | Create date: 14/11/2026 | Calcula el saldo de ventas aplicando NC registradas y cierres de factoring.
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_ComprobanteHYT_SaldoVenta(@ComprobanteId BIGINT)
RETURNS TABLE
AS
RETURN
(
    SELECT lote.LoteId, lote.FechaDesembolso, lote.Estado AS EstadoFactoring,
        credito.ImporteAplicado,
        CONVERT(DECIMAL(18,2), CASE
            WHEN documento.Tipo='07' OR lote.Estado=N'CERRADO' OR saldo.Importe<=0 THEN 0
            ELSE saldo.Importe
        END) AS SaldoPendiente
    FROM dbo.ComprobanteHYT AS documento
    LEFT JOIN dbo.FactoringFacturaHYT AS vinculo ON vinculo.ComprobanteId=documento.ComprobanteId
    LEFT JOIN dbo.FactoringLoteHYT AS lote ON lote.LoteId=vinculo.LoteId
    OUTER APPLY
    (
        SELECT CONVERT(DECIMAL(33,2), COALESCE(SUM(nota.Importe),0)) AS ImporteAplicado
        FROM dbo.ComprobanteHYT AS nota
        WHERE documento.Tipo IN ('01','03') AND documento.Serie IS NOT NULL AND documento.Numero IS NOT NULL
            AND nota.EsVenta=1 AND nota.Tipo='07' AND nota.Serie IS NOT NULL AND nota.Numero IS NOT NULL
            AND nota.EstadoSunat NOT IN (N'RECHAZADO',N'ANULADO',N'BAJA')
            AND nota.RucHYT=documento.RucHYT AND nota.ClienteId=documento.ClienteId AND nota.Moneda=documento.Moneda
            AND LTRIM(RTRIM(nota.ReferenciaNota))=CONCAT(documento.Serie,N'-',documento.Numero)
    ) AS credito
    CROSS APPLY (VALUES (documento.Importe-documento.DetraccionDocumento-credito.ImporteAplicado)) AS saldo(Importe)
    WHERE documento.ComprobanteId=@ComprobanteId AND documento.EsVenta=1
);
GO
