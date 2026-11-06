-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define TarifarioVersionZona para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioVersionZona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioVersionZona
    (
        VersionId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_TarifarioVersionZona PRIMARY KEY (VersionId, ZonaId),
        CONSTRAINT FK_TarifarioVersionZona_Version FOREIGN KEY (VersionId) REFERENCES dbo.TarifarioVersion(VersionId),
        CONSTRAINT FK_TarifarioVersionZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
    );
END;
GO
