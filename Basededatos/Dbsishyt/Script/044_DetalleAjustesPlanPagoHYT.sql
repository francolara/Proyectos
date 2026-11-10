-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Conserva ajustes detallados de extras y descuentos del periodo.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.PlanPagoDetalleHYT', N'Ajustes') IS NULL
    ALTER TABLE dbo.PlanPagoDetalleHYT ADD Ajustes NVARCHAR(MAX) NULL
        CONSTRAINT CK_PlanPagoDetalleHYT_Ajustes CHECK (ISJSON(Ajustes, ARRAY) = 1);
GO
