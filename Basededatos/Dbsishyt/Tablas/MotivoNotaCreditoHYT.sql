-- Firma: FRANCO LARA | Create date: 11/11/2026 | Catálogo de motivos de nota de crédito con códigos SUNAT.
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.MotivoNotaCreditoHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MotivoNotaCreditoHYT
    (
        Codigo CHAR(2) NOT NULL CONSTRAINT PK_MotivoNotaCreditoHYT PRIMARY KEY,
        Descripcion NVARCHAR(300) NOT NULL
    );
END;
GO
