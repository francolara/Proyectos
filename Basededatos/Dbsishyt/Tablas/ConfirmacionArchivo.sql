-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ConfirmacionArchivo para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ConfirmacionArchivo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConfirmacionArchivo
    (
        ConfirmacionId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_ConfirmacionArchivo PRIMARY KEY (ConfirmacionId, ArchivoId),
        CONSTRAINT FK_ConfirmacionArchivo_Owner FOREIGN KEY (ConfirmacionId) REFERENCES dbo.ConfirmacionFacturacion(ConfirmacionId),
        CONSTRAINT FK_ConfirmacionArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO
