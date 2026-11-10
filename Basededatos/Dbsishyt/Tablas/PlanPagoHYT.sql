-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   PlanPagoHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PlanPagoHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PlanPagoHYT
    (
        PlanId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PlanPagoHYT PRIMARY KEY,
        Frecuencia VARCHAR(10) NOT NULL,
        Desde DATE NOT NULL,
        Hasta DATE NOT NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT UQ_PlanPagoHYT_Periodo UNIQUE (Frecuencia, Desde),
        CONSTRAINT CK_PlanPagoHYT_Rango CHECK (Desde <= Hasta AND Frecuencia IN ('SEMANAL', 'QUINCENAL', 'MENSUAL'))
    );
END;
GO
