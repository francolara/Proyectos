-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Firma:         Columnas para contrato de suscripcion (tipo de cobro y gracia) en NegociosSuscripcion.
-- =============================================
IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.NegociosSuscripcion', 'TipoCobro') IS NULL
        ALTER TABLE dbo.NegociosSuscripcion ADD TipoCobro NVARCHAR(20) NULL;

    IF COL_LENGTH('dbo.NegociosSuscripcion', 'DiasGracia') IS NULL
        ALTER TABLE dbo.NegociosSuscripcion ADD DiasGracia INT NOT NULL CONSTRAINT DF_NegociosSuscripcion_DiasGracia DEFAULT (5);

    IF COL_LENGTH('dbo.NegociosSuscripcion', 'FechaFinGracia') IS NULL
        ALTER TABLE dbo.NegociosSuscripcion ADD FechaFinGracia DATE NULL;
END
