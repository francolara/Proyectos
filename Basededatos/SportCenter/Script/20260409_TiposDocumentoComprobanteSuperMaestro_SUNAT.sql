/*
Firma: Codex - 09/04/2026
Descripcion: Carga completa de tipos de documento/comprobante SUNAT (Catalogo 01) en supermaestro.
Fuente: https://www.sunat.gob.pe/legislacion/superin/2017/anexoE-245-2017.pdf (Catalogo No. 01)
Regla: Todos los codigos SUNAT se cargan con Tributario = 1; solo Factura (01) y Boleta (03) habilitados.
Adicional: Recibo Interno (RI) como no tributario y habilitado.
*/
USE [DbSportCenter]
GO

SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TiposDocumentoComprobanteSuperMaestro', N'U') IS NULL
BEGIN
    RAISERROR('No existe la tabla dbo.TiposDocumentoComprobanteSuperMaestro. Ejecuta primero script de estructura.', 16, 1);
    RETURN;
END;

MERGE dbo.TiposDocumentoComprobanteSuperMaestro AS tgt
USING
(
    VALUES
        (N'01', N'FACTURA', CAST(1 AS bit), CAST(1 AS bit), CAST(1 AS tinyint), CAST(1 AS bit)),
        (N'03', N'BOLETA DE VENTA', CAST(1 AS bit), CAST(1 AS bit), CAST(2 AS tinyint), CAST(1 AS bit)),
        (N'07', N'NOTA DE CREDITO', CAST(1 AS bit), CAST(0 AS bit), CAST(3 AS tinyint), CAST(1 AS bit)),
        (N'08', N'NOTA DE DEBITO', CAST(1 AS bit), CAST(0 AS bit), CAST(4 AS tinyint), CAST(1 AS bit)),
        (N'09', N'GUIA DE REMISION REMITENTE', CAST(1 AS bit), CAST(0 AS bit), CAST(5 AS tinyint), CAST(1 AS bit)),
        (N'12', N'TICKET DE MAQUINA REGISTRADORA', CAST(1 AS bit), CAST(0 AS bit), CAST(6 AS tinyint), CAST(1 AS bit)),
        (N'13', N'DOCUMENTO EMITIDO POR BANCOS, INSTITUCIONES FINANCIERAS, CREDITICIAS Y DE SEGUROS BAJO CONTROL SBS', CAST(1 AS bit), CAST(0 AS bit), CAST(7 AS tinyint), CAST(1 AS bit)),
        (N'14', N'RECIBO SERVICIOS PUBLICOS', CAST(1 AS bit), CAST(0 AS bit), CAST(8 AS tinyint), CAST(1 AS bit)),
        (N'16', N'BOLETO DE VIAJE EMITIDO POR EMPRESAS DE TRANSPORTE PUBLICO INTERPROVINCIAL DE PASAJEROS', CAST(1 AS bit), CAST(0 AS bit), CAST(9 AS tinyint), CAST(1 AS bit)),
        (N'18', N'DOCUMENTOS EMITIDOS POR LAS AFP', CAST(1 AS bit), CAST(0 AS bit), CAST(10 AS tinyint), CAST(1 AS bit)),
        (N'20', N'COMPROBANTE DE RETENCION', CAST(1 AS bit), CAST(0 AS bit), CAST(11 AS tinyint), CAST(1 AS bit)),
        (N'31', N'GUIA DE REMISION TRANSPORTISTA', CAST(1 AS bit), CAST(0 AS bit), CAST(12 AS tinyint), CAST(1 AS bit)),
        (N'40', N'COMPROBANTE DE PERCEPCION', CAST(1 AS bit), CAST(0 AS bit), CAST(13 AS tinyint), CAST(1 AS bit)),
        (N'41', N'COMPROBANTE DE PERCEPCION - VENTA INTERNA (FISICO - FORMATO IMPRESO)', CAST(1 AS bit), CAST(0 AS bit), CAST(14 AS tinyint), CAST(1 AS bit)),
        (N'56', N'COMPROBANTE DE PAGO SEAE', CAST(1 AS bit), CAST(0 AS bit), CAST(15 AS tinyint), CAST(1 AS bit)),
        (N'71', N'GUIA DE REMISION REMITENTE COMPLEMENTARIA', CAST(1 AS bit), CAST(0 AS bit), CAST(16 AS tinyint), CAST(1 AS bit)),
        (N'72', N'GUIA DE REMISION TRANSPORTISTA COMPLEMENTARIA', CAST(1 AS bit), CAST(0 AS bit), CAST(17 AS tinyint), CAST(1 AS bit)),
        (N'RI', N'RECIBO INTERNO', CAST(0 AS bit), CAST(1 AS bit), CAST(18 AS tinyint), CAST(1 AS bit))
) AS src (CodigoSunat, Nombre, Tributario, Habilitado, Orden, Activo)
ON tgt.CodigoSunat = src.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        tgt.Nombre = src.Nombre,
        tgt.Tributario = src.Tributario,
        tgt.Habilitado = src.Habilitado,
        tgt.Orden = src.Orden,
        tgt.Activo = src.Activo
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoSunat, Nombre, Tributario, Habilitado, Orden, Activo)
    VALUES (src.CodigoSunat, src.Nombre, src.Tributario, src.Habilitado, src.Orden, src.Activo);
