-- Firma: FRANCO LARA | Create date: 06/11/2026 | Conserva nombres completos del tipo de servicio en el detalle del presupuesto.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'TipoServicio') IS NOT NULL ALTER TABLE dbo.PresupuestoPartida ALTER COLUMN TipoServicio NVARCHAR(150) NULL;
GO
