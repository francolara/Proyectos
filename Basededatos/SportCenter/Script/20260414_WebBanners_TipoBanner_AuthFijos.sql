USE [DbSportCenter]
GO

-- Firma: Codex - 14/04/2026 | Agrega TipoBanner para Home/Login/Registro y normaliza valor inicial.
IF COL_LENGTH(''dbo.WebBanners'', ''TipoBanner'') IS NULL
BEGIN
    ALTER TABLE dbo.WebBanners
    ADD TipoBanner TINYINT NOT NULL
        CONSTRAINT DF_WebBanners_TipoBanner DEFAULT ((1));
END
GO

UPDATE dbo.WebBanners
SET TipoBanner = 1
WHERE TipoBanner IS NULL OR TipoBanner NOT IN (1,2,3);
GO