/*
Firma: Codex - 09/04/2026
Descripcion: ALTER TABLE de Negocios para emision de comprobantes y configuracion de IGV.
*/
USE [DbSportCenter]
GO

IF COL_LENGTH('dbo.Negocios', 'EmisionComprobantesElectronicos') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD EmisionComprobantesElectronicos BIT NOT NULL
        CONSTRAINT DF_Negocios_EmisionComprobantesElectronicos DEFAULT (0);
END
GO

IF COL_LENGTH('dbo.Negocios', 'EmisionReciboInterno') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD EmisionReciboInterno BIT NOT NULL
        CONSTRAINT DF_Negocios_EmisionReciboInterno DEFAULT (0);
END
GO

IF COL_LENGTH('dbo.Negocios', 'PorcentajeIgv') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD PorcentajeIgv INT NOT NULL
        CONSTRAINT DF_Negocios_PorcentajeIgv DEFAULT (18);
END
GO
