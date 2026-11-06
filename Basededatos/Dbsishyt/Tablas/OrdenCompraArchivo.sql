-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define OrdenCompraArchivo para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.OrdenCompraArchivo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OrdenCompraArchivo
    (
        OrdenCompraId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_OrdenCompraArchivo PRIMARY KEY (OrdenCompraId, ArchivoId),
        CONSTRAINT FK_OrdenCompraArchivo_Owner FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        CONSTRAINT FK_OrdenCompraArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO
