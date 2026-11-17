-- Firma: FRANCO LARA | Create date: 17/11/2026 | Agrega control de concurrencia para la edición de comprobantes.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.ComprobanteHYT',N'VersionFila') IS NULL
    ALTER TABLE dbo.ComprobanteHYT ADD VersionFila ROWVERSION NOT NULL;
GO
