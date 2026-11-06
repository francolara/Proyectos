-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define GuiaArchivo para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.GuiaArchivo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GuiaArchivo
    (
        GuiaId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_GuiaArchivo PRIMARY KEY (GuiaId, ArchivoId),
        CONSTRAINT FK_GuiaArchivo_Owner FOREIGN KEY (GuiaId) REFERENCES dbo.GuiaConformidad(GuiaId),
        CONSTRAINT FK_GuiaArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO
