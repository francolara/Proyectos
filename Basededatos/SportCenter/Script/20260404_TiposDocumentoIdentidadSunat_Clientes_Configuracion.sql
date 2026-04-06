USE [DbSportCenter]
GO
-- Firma: Codex - 04/04/2026 | Seed SUNAT de tipos de documento, ajuste de longitudes y control de indices/FK en Clientes/Negocios para uso en comprobante electronico.

SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TiposDocumentoIdentidadSunat', N'U') IS NULL
BEGIN
    RAISERROR('Primero ejecuta dbo.TiposDocumentoIdentidadSunat.Table.sql.', 16, 1);
    RETURN;
END;

MERGE dbo.TiposDocumentoIdentidadSunat AS tgt
USING
(
    VALUES
        (N'0', N'OTRO', N'Doc. trib. no dom. sin RUC', CAST(1 AS bit), CAST(1 AS tinyint)),
        (N'1', N'DNI', N'DNI', CAST(1 AS bit), CAST(2 AS tinyint)),
        (N'4', N'CE', N'Carnet de extranjeria', CAST(1 AS bit), CAST(3 AS tinyint)),
        (N'6', N'RUC', N'RUC', CAST(1 AS bit), CAST(4 AS tinyint)),
        (N'7', N'PASAPORTE', N'Pasaporte', CAST(1 AS bit), CAST(5 AS tinyint)),
        (N'A', N'CED_DIPLOMATICA', N'Cedula diplomatica de identidad', CAST(1 AS bit), CAST(6 AS tinyint))
) AS src (CodigoSunat, CodigoInterno, Nombre, Activo, Orden)
ON tgt.CodigoSunat = src.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        tgt.CodigoInterno = src.CodigoInterno,
        tgt.Nombre = src.Nombre,
        tgt.Activo = src.Activo,
        tgt.Orden = src.Orden
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoSunat, CodigoInterno, Nombre, Activo, Orden)
    VALUES (src.CodigoSunat, src.CodigoInterno, src.Nombre, src.Activo, src.Orden);

UPDATE c
SET c.TipoDocumento = m.CodigoSunat
FROM dbo.Clientes c
INNER JOIN
(
    SELECT N'DNI' AS ValorAnterior, N'1' AS CodigoSunat UNION ALL
    SELECT N'RUC', N'6' UNION ALL
    SELECT N'CE', N'4' UNION ALL
    SELECT N'PASAPORTE', N'7' UNION ALL
    SELECT N'OTRO', N'0'
) m ON UPPER(LTRIM(RTRIM(c.TipoDocumento))) = m.ValorAnterior;

UPDATE n
SET n.TipoDocumentoFiscal = m.CodigoSunat
FROM dbo.Negocios n
INNER JOIN
(
    SELECT N'DNI' AS ValorAnterior, N'1' AS CodigoSunat UNION ALL
    SELECT N'RUC', N'6' UNION ALL
    SELECT N'CE', N'4' UNION ALL
    SELECT N'PASAPORTE', N'7' UNION ALL
    SELECT N'OTRO', N'0'
) m ON UPPER(LTRIM(RTRIM(COALESCE(n.TipoDocumentoFiscal, N'')))) = m.ValorAnterior;

UPDATE c
SET c.TipoDocumento = UPPER(LTRIM(RTRIM(c.TipoDocumento)))
FROM dbo.Clientes c
WHERE c.TipoDocumento IS NOT NULL;

UPDATE n
SET n.TipoDocumentoFiscal = UPPER(LTRIM(RTRIM(n.TipoDocumentoFiscal)))
FROM dbo.Negocios n
WHERE n.TipoDocumentoFiscal IS NOT NULL;

IF EXISTS
(
    SELECT 1
    FROM dbo.Clientes c
    WHERE NULLIF(LTRIM(RTRIM(c.TipoDocumento)), N'') IS NOT NULL
      AND NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = c.TipoDocumento)
)
BEGIN
    RAISERROR('Existen clientes con TipoDocumento sin equivalencia SUNAT.', 16, 1);
    RETURN;
END;

IF EXISTS
(
    SELECT 1
    FROM dbo.Negocios n
    WHERE NULLIF(LTRIM(RTRIM(COALESCE(n.TipoDocumentoFiscal, N''))), N'') IS NOT NULL
      AND NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = n.TipoDocumentoFiscal)
)
BEGIN
    RAISERROR('Existen negocios con TipoDocumentoFiscal sin equivalencia SUNAT.', 16, 1);
    RETURN;
END;

IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Clientes_TiposDocumentoIdentidadSunat_TipoDocumento')
    ALTER TABLE dbo.Clientes DROP CONSTRAINT [FK_Clientes_TiposDocumentoIdentidadSunat_TipoDocumento];

IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Negocios_TiposDocumentoIdentidadSunat_TipoDocumentoFiscal')
    ALTER TABLE dbo.Negocios DROP CONSTRAINT [FK_Negocios_TiposDocumentoIdentidadSunat_TipoDocumentoFiscal];

IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Clientes_TipoDocumento_NumeroDocumento' AND object_id = OBJECT_ID(N'dbo.Clientes'))
    DROP INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON dbo.Clientes;

ALTER TABLE dbo.Clientes ALTER COLUMN TipoDocumento NVARCHAR(2) NOT NULL;
ALTER TABLE dbo.Negocios ALTER COLUMN TipoDocumentoFiscal NVARCHAR(2) NULL;

ALTER TABLE dbo.Clientes WITH CHECK
ADD CONSTRAINT [FK_Clientes_TiposDocumentoIdentidadSunat_TipoDocumento]
FOREIGN KEY ([TipoDocumento]) REFERENCES dbo.TiposDocumentoIdentidadSunat ([CodigoSunat]);

ALTER TABLE dbo.Clientes CHECK CONSTRAINT [FK_Clientes_TiposDocumentoIdentidadSunat_TipoDocumento];

ALTER TABLE dbo.Negocios WITH CHECK
ADD CONSTRAINT [FK_Negocios_TiposDocumentoIdentidadSunat_TipoDocumentoFiscal]
FOREIGN KEY ([TipoDocumentoFiscal]) REFERENCES dbo.TiposDocumentoIdentidadSunat ([CodigoSunat]);

ALTER TABLE dbo.Negocios CHECK CONSTRAINT [FK_Negocios_TiposDocumentoIdentidadSunat_TipoDocumentoFiscal];

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Clientes_TipoDocumento_NumeroDocumento' AND object_id = OBJECT_ID(N'dbo.Clientes'))
    CREATE INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON dbo.Clientes (TipoDocumento, NumeroDocumento);
