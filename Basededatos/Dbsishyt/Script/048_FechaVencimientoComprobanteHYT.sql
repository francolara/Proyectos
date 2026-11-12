-- Firma: FRANCO LARA | Create date: 10/11/2026 | Agrega el vencimiento opcional a compras y ventas.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ComprobanteHYT',N'FechaVencimiento') IS NULL
    ALTER TABLE dbo.ComprobanteHYT ADD FechaVencimiento DATE NULL;
GO
