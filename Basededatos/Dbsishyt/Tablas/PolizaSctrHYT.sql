-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.PolizaSctrHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PolizaSctrHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PolizaSctrHYT
    (
        PolizaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PolizaSctrHYT PRIMARY KEY,
        Aseguradora NVARCHAR(120) NOT NULL,
        NumeroPoliza NVARCHAR(80) NOT NULL,
        Cobertura VARCHAR(10) NOT NULL CONSTRAINT CK_PolizaSctrHYT_Cobertura CHECK (Cobertura IN ('SALUD', 'PENSION')),
        VigenciaDesde DATE NOT NULL,
        VigenciaHasta DATE NOT NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_PolizaSctrHYT_Fechas CHECK (VigenciaDesde <= VigenciaHasta),
        CONSTRAINT UQ_PolizaSctrHYT_Vigencia UNIQUE (Aseguradora, NumeroPoliza, Cobertura, VigenciaDesde)
    );
END;
GO
