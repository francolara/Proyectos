-- Firma: FRANCO LARA | Create date: 11/11/2026 | Separa nombres y apellidos en la ficha del trabajador.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.TrabajadorHYT',N'Nombres') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD Nombres NVARCHAR(179) NULL;
IF COL_LENGTH(N'dbo.TrabajadorHYT',N'Apellidos') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD Apellidos NVARCHAR(179) NULL;
GO
