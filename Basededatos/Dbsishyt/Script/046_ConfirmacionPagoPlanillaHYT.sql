-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Agrega fecha y referencia del pago confirmado de una planificacion.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.PlanPagoHYT', N'FechaPago') IS NULL
    ALTER TABLE dbo.PlanPagoHYT ADD FechaPago DATE NULL, ReferenciaPago NVARCHAR(100) NULL;
GO
IF OBJECT_ID(N'dbo.CK_PlanPagoHYT_Pago', N'C') IS NULL
    ALTER TABLE dbo.PlanPagoHYT ADD CONSTRAINT CK_PlanPagoHYT_Pago CHECK (
        (FechaPago IS NULL AND ReferenciaPago IS NULL)
        OR (FechaPago IS NOT NULL AND FechaPago >= Desde AND ReferenciaPago IS NOT NULL));
GO
