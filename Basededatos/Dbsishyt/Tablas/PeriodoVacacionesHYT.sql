-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.PeriodoVacacionesHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PeriodoVacacionesHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PeriodoVacacionesHYT
    (
        PeriodoVacacionesId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PeriodoVacacionesHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_PeriodoVacaciones_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        FechaDesde DATE NOT NULL,
        FechaHasta DATE NOT NULL,
        DiasReconocidos SMALLINT NOT NULL CONSTRAINT CK_PeriodoVacaciones_Dias CHECK (DiasReconocidos > 0),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_PeriodoVacaciones_Fechas CHECK (FechaHasta >= FechaDesde),
        CONSTRAINT UQ_PeriodoVacaciones_TrabajadorDesde UNIQUE (TrabajadorId, FechaDesde)
    );
END;
GO
