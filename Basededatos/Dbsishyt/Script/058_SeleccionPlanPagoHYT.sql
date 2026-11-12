-- Firma: FRANCO LARA | Create date: 11/11/2026 | Persiste la inclusión de cada trabajador y conserva incluidos los detalles existentes.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.PlanPagoDetalleHYT', N'Incluido') IS NULL
    ALTER TABLE dbo.PlanPagoDetalleHYT ADD Incluido BIT NOT NULL
        CONSTRAINT DF_PlanPagoDetalleHYT_Incluido DEFAULT (1);
GO
