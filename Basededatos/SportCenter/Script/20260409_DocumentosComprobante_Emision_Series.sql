USE [DbSportCenter]
GO
-- Firma: Codex - 09/04/2026 | Crea supermaestro de documentos de comprobante, documentos por negocio, series por negocio/sede y campos de emision+IGV en Negocios.

SET NOCOUNT ON;

IF COL_LENGTH('dbo.Negocios', 'PorcentajeIgv') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios ADD PorcentajeIgv INT NOT NULL CONSTRAINT DF_Negocios_PorcentajeIgv DEFAULT (18);
END;

IF COL_LENGTH('dbo.Negocios', 'EmisionComprobantesElectronicos') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios ADD EmisionComprobantesElectronicos BIT NOT NULL CONSTRAINT DF_Negocios_EmisionComprobantesElectronicos DEFAULT (0);
END;

IF COL_LENGTH('dbo.Negocios', 'EmisionReciboInterno') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios ADD EmisionReciboInterno BIT NOT NULL CONSTRAINT DF_Negocios_EmisionReciboInterno DEFAULT (0);
END;

IF OBJECT_ID(N'dbo.TiposDocumentoComprobanteSuperMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TiposDocumentoComprobanteSuperMaestro
    (
        CodigoSunat NVARCHAR(4) NOT NULL CONSTRAINT PK_TiposDocumentoComprobanteSuperMaestro PRIMARY KEY,
        Nombre NVARCHAR(150) NOT NULL,
        Tributario BIT NOT NULL CONSTRAINT DF_TiposDocumentoComprobanteSuperMaestro_Tributario DEFAULT (1),
        Habilitado BIT NOT NULL CONSTRAINT DF_TiposDocumentoComprobanteSuperMaestro_Habilitado DEFAULT (0),
        Orden TINYINT NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_TiposDocumentoComprobanteSuperMaestro_Activo DEFAULT (1)
    );
END;

IF OBJECT_ID(N'dbo.NegociosTiposDocumentoComprobante', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NegociosTiposDocumentoComprobante
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_NegociosTiposDocumentoComprobante PRIMARY KEY,
        NegocioId INT NOT NULL,
        CodigoSunat NVARCHAR(4) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_NegociosTiposDocumentoComprobante_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_NegociosTiposDocumentoComprobante_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioCreacion NVARCHAR(200) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );

    ALTER TABLE dbo.NegociosTiposDocumentoComprobante
        ADD CONSTRAINT UX_NegociosTiposDocumentoComprobante_Negocio_Documento UNIQUE (NegocioId, CodigoSunat);

    ALTER TABLE dbo.NegociosTiposDocumentoComprobante
        ADD CONSTRAINT FK_NegociosTiposDocumentoComprobante_Negocios
        FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);

    ALTER TABLE dbo.NegociosTiposDocumentoComprobante
        ADD CONSTRAINT FK_NegociosTiposDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro
        FOREIGN KEY (CodigoSunat) REFERENCES dbo.TiposDocumentoComprobanteSuperMaestro(CodigoSunat);
END;

IF OBJECT_ID(N'dbo.NegociosSeriesDocumentoComprobante', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NegociosSeriesDocumentoComprobante
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_NegociosSeriesDocumentoComprobante PRIMARY KEY,
        NegocioId INT NOT NULL,
        CodigoSunat NVARCHAR(4) NOT NULL,
        Serie NVARCHAR(4) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_NegociosSeriesDocumentoComprobante_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_NegociosSeriesDocumentoComprobante_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioCreacion NVARCHAR(200) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );

    ALTER TABLE dbo.NegociosSeriesDocumentoComprobante
        ADD CONSTRAINT UX_NegociosSeriesDocumentoComprobante_Negocio_Documento_Serie UNIQUE (NegocioId, CodigoSunat, Serie);

    ALTER TABLE dbo.NegociosSeriesDocumentoComprobante
        ADD CONSTRAINT FK_NegociosSeriesDocumentoComprobante_Negocios
        FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);

    ALTER TABLE dbo.NegociosSeriesDocumentoComprobante
        ADD CONSTRAINT FK_NegociosSeriesDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro
        FOREIGN KEY (CodigoSunat) REFERENCES dbo.TiposDocumentoComprobanteSuperMaestro(CodigoSunat);
END;

IF OBJECT_ID(N'dbo.SedesSeriesDocumentoComprobante', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SedesSeriesDocumentoComprobante
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SedesSeriesDocumentoComprobante PRIMARY KEY,
        SedeId INT NOT NULL,
        CodigoSunat NVARCHAR(4) NOT NULL,
        NegocioSerieId INT NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_SedesSeriesDocumentoComprobante_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_SedesSeriesDocumentoComprobante_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioCreacion NVARCHAR(200) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );

    ALTER TABLE dbo.SedesSeriesDocumentoComprobante
        ADD CONSTRAINT UX_SedesSeriesDocumentoComprobante_Sede_Documento UNIQUE (SedeId, CodigoSunat);

    ALTER TABLE dbo.SedesSeriesDocumentoComprobante
        ADD CONSTRAINT FK_SedesSeriesDocumentoComprobante_Sedes
        FOREIGN KEY (SedeId) REFERENCES dbo.Sedes(Id);

    ALTER TABLE dbo.SedesSeriesDocumentoComprobante
        ADD CONSTRAINT FK_SedesSeriesDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro
        FOREIGN KEY (CodigoSunat) REFERENCES dbo.TiposDocumentoComprobanteSuperMaestro(CodigoSunat);

    ALTER TABLE dbo.SedesSeriesDocumentoComprobante
        ADD CONSTRAINT FK_SedesSeriesDocumentoComprobante_NegociosSeriesDocumentoComprobante
        FOREIGN KEY (NegocioSerieId) REFERENCES dbo.NegociosSeriesDocumentoComprobante(Id);
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

;WITH DocsNegocioBase AS
(
    SELECT n.Id AS NegocioId, t.CodigoSunat
    FROM dbo.Negocios n
    CROSS JOIN dbo.TiposDocumentoComprobanteSuperMaestro t
    WHERE n.Activo = 1
      AND t.Activo = 1
      AND t.Habilitado = 1
)
MERGE dbo.NegociosTiposDocumentoComprobante AS tgt
USING DocsNegocioBase AS src
ON tgt.NegocioId = src.NegocioId AND tgt.CodigoSunat = src.CodigoSunat
WHEN NOT MATCHED BY TARGET THEN
    INSERT (NegocioId, CodigoSunat, Activo, UsuarioCreacion)
    VALUES (src.NegocioId, src.CodigoSunat, 1, N'sistema');

