-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.VacacionesHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.VacacionesHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.VacacionesHYT
    (
        VacacionesId BIGINT IDENTITY NOT NULL CONSTRAINT PK_VacacionesHYT PRIMARY KEY,
        PeriodoVacacionesId BIGINT NOT NULL CONSTRAINT FK_Vacaciones_Periodo REFERENCES dbo.PeriodoVacacionesHYT(PeriodoVacacionesId),
        FechaInicio DATE NOT NULL,
        FechaFin DATE NOT NULL,
        Dias SMALLINT NOT NULL,
        Estado VARCHAR(10) NOT NULL CONSTRAINT CK_Vacaciones_Estado CHECK (Estado IN ('PENDIENTE', 'APROBADA', 'ANULADA')),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_Vacaciones_FechasDias CHECK (FechaFin >= FechaInicio AND Dias BETWEEN 1 AND DATEDIFF(DAY, FechaInicio, FechaFin) + 1)
    );
END;
GO
