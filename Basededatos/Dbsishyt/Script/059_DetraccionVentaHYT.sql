-- Firma: FRANCO LARA | Create date: 11/11/2026 | Conserva el porcentaje aplicado y recupera la detracción de facturas anteriores.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.ComprobanteHYT',N'PorcentajeDetraccion') IS NULL
    ALTER TABLE dbo.ComprobanteHYT ADD PorcentajeDetraccion DECIMAL(7,4) NULL
        CONSTRAINT CK_ComprobanteHYT_PorcentajeDetraccion CHECK (PorcentajeDetraccion BETWEEN 0 AND 100);
GO

;WITH XMLNAMESPACES (
    'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2' AS cac,
    'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' AS cbc
)
UPDATE comprobante
SET CodigoDetraccion=COALESCE(material.Codigo,termino.Codigo,comprobante.CodigoDetraccion),
    PorcentajeDetraccion=tasa.Porcentaje,
    DetraccionDocumento=importe.Equivalente,
    ImporteDetraccion=CASE WHEN deposito.OtraMoneda=1 THEN termino.Importe ELSE importe.Equivalente END,
    MonedaDetraccion=CASE WHEN deposito.OtraMoneda=1 THEN termino.Moneda ELSE comprobante.Moneda END
FROM dbo.ComprobanteHYT AS comprobante
OUTER APPLY dbo.fn_ComprobanteHYT_DetraccionMaterial(CASE WHEN comprobante.XmlOriginal IS NULL THEN comprobante.Detalle END) AS material
OUTER APPLY (SELECT TRY_CONVERT(XML,comprobante.XmlOriginal) AS Documento) AS origen
OUTER APPLY (
    SELECT TOP (1) NULLIF(TRIM(nodo.Termino.value('(cbc:PaymentMeansID/text())[1]','nvarchar(30)')),N'') AS Codigo,
        datos.Porcentaje,datos.Importe,datos.Moneda
    FROM origen.Documento.nodes('/*/cac:PaymentTerms') AS nodo(Termino)
    CROSS APPLY (VALUES (
        TRY_CONVERT(DECIMAL(7,4),nodo.Termino.value('(cbc:PaymentPercent/text())[1]','nvarchar(40)')),
        TRY_CONVERT(DECIMAL(18,2),nodo.Termino.value('(cbc:Amount/text())[1]','nvarchar(40)')),
        nodo.Termino.value('(cbc:Amount/@currencyID)[1]','nvarchar(3)')
    )) AS datos(Porcentaje,Importe,Moneda)
    WHERE UPPER(nodo.Termino.value('(cbc:ID/text())[1]','nvarchar(40)'))=N'DETRACCION'
    ORDER BY datos.Porcentaje DESC,datos.Importe DESC
) AS termino
LEFT JOIN dbo.DetraccionCatalogoHYT AS catalogo
    ON catalogo.Codigo=COALESCE(termino.Codigo,comprobante.CodigoDetraccion)
CROSS APPLY (VALUES (COALESCE(
    CASE WHEN termino.Porcentaje BETWEEN 0 AND 100 THEN termino.Porcentaje END,
    material.Porcentaje,catalogo.Porcentaje,
    CASE WHEN termino.Moneda=comprobante.Moneda AND termino.Importe BETWEEN 0 AND comprobante.Importe
        THEN TRY_CONVERT(DECIMAL(7,4),100*termino.Importe/NULLIF(comprobante.Importe,0)) END,
    0
))) AS tasa(Porcentaje)
CROSS APPLY (VALUES (CASE WHEN termino.Moneda IN (N'PEN',N'USD') AND termino.Moneda<>comprobante.Moneda
    AND termino.Importe>0 AND tasa.Porcentaje>0 THEN 1 ELSE 0 END)) AS deposito(OtraMoneda)
CROSS APPLY (VALUES (COALESCE(
    CASE WHEN COALESCE(termino.Moneda,comprobante.Moneda)=comprobante.Moneda AND termino.Importe>0
        AND termino.Importe<=comprobante.Importe THEN termino.Importe END,
    NULLIF(comprobante.DetraccionDocumento,0),
    ROUND(comprobante.Importe*tasa.Porcentaje/100,2)
))) AS importe(Equivalente)
WHERE comprobante.EsVenta=1 AND comprobante.Tipo='01' AND comprobante.PorcentajeDetraccion IS NULL;
GO
