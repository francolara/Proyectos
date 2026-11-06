-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define TarifarioPartidaZona para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioPartidaZona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioPartidaZona
    (
        PartidaId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_TarifarioPartidaZona PRIMARY KEY (PartidaId, ZonaId),
        CONSTRAINT FK_TarifarioPartidaZona_Partida FOREIGN KEY (PartidaId) REFERENCES dbo.TarifarioPartida(PartidaId),
        CONSTRAINT FK_TarifarioPartidaZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
    );
END;
GO
