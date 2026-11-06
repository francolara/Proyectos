-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define MaterialZona para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.MaterialZona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MaterialZona
    (
        MaterialId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_MaterialZona PRIMARY KEY (MaterialId, ZonaId),
        CONSTRAINT FK_MaterialZona_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_MaterialZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
    );
END;
GO
