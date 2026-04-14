USE [DbSportCenter];

-- Firma: Codex - 13/04/2026 | Agrega columna LogoUrl en Negocios para almacenar logo del club en bucket.
IF COL_LENGTH('dbo.Negocios', 'LogoUrl') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD LogoUrl NVARCHAR(500) NULL;
END;
