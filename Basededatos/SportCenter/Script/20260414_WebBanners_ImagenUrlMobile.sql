USE [DbSportCenter]
GO

-- Firma: Codex - 14/04/2026 | Agrega columna ImagenUrlMobile para banner responsive en mobile.
IF COL_LENGTH(''dbo.WebBanners'', ''ImagenUrlMobile'') IS NULL
BEGIN
    ALTER TABLE dbo.WebBanners
    ADD ImagenUrlMobile NVARCHAR(500) NULL;
END
GO