USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 16/04/2026 | ALTER TABLE incremental para columnas nuevas en Negocios (limites/flags) y Sedes (redes sociales).

/* ==============================
   NEGOCIOS
   ============================== */
IF COL_LENGTH('dbo.Negocios', 'PermitirModificarPrecioReserva') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD PermitirModificarPrecioReserva BIT NOT NULL
        CONSTRAINT DF_Negocios_PermitirModificarPrecioReserva DEFAULT ((0));
END
GO

IF COL_LENGTH('dbo.Negocios', 'CancelacionAutomaticaNoConfirmada') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD CancelacionAutomaticaNoConfirmada BIT NOT NULL
        CONSTRAINT DF_Negocios_CancelacionAutomaticaNoConfirmada DEFAULT ((0));
END
GO

IF COL_LENGTH('dbo.Negocios', 'MinutosCancelacionNoConfirmada') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD MinutosCancelacionNoConfirmada INT NULL;
END
GO

IF COL_LENGTH('dbo.Negocios', 'SedesPermitidas') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD SedesPermitidas INT NOT NULL
        CONSTRAINT DF_Negocios_SedesPermitidas DEFAULT ((2));
END
GO

IF COL_LENGTH('dbo.Negocios', 'EspaciosPermitidos') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD EspaciosPermitidos INT NOT NULL
        CONSTRAINT DF_Negocios_EspaciosPermitidos DEFAULT ((6));
END
GO

/* ==============================
   SEDES
   ============================== */
IF COL_LENGTH('dbo.Sedes', 'FacebookUrl') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD FacebookUrl NVARCHAR(500) NULL;
END
GO

IF COL_LENGTH('dbo.Sedes', 'InstagramUrl') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD InstagramUrl NVARCHAR(500) NULL;
END
GO

IF COL_LENGTH('dbo.Sedes', 'TwitterUrl') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD TwitterUrl NVARCHAR(500) NULL;
END
GO
