-- Firma: FRANCO LARA | Create date: 14/11/2026 | Vincula la póliza SCTR con su archivo privado.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.PolizaSctrHYT',N'ArchivoId') IS NULL
    ALTER TABLE dbo.PolizaSctrHYT ADD ArchivoId BIGINT NULL
        CONSTRAINT FK_PolizaSctrHYT_Archivo REFERENCES dbo.ArchivoAdjunto(ArchivoId);
GO
