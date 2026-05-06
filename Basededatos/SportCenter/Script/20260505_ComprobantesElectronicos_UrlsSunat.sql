USE [DbSportCenter]
GO

-- Firma: Codex - 05/05/2026 | Agrega columnas URL PDF/XML/CDR para documentos SUNAT en comprobantes electronicos.
IF COL_LENGTH(''dbo.ComprobantesElectronicos'', ''UrlPdfSunat'') IS NULL
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
    ADD UrlPdfSunat NVARCHAR(500) NULL;
END
GO

IF COL_LENGTH(''dbo.ComprobantesElectronicos'', ''UrlXmlSunat'') IS NULL
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
    ADD UrlXmlSunat NVARCHAR(500) NULL;
END
GO

IF COL_LENGTH(''dbo.ComprobantesElectronicos'', ''UrlCdrSunat'') IS NULL
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
    ADD UrlCdrSunat NVARCHAR(500) NULL;
END
GO
