USE [DbSportCenter];

IF COL_LENGTH('dbo.EspaciosDeportivos', 'AdministracionPrivada') IS NULL
BEGIN
    ALTER TABLE dbo.EspaciosDeportivos
    ADD AdministracionPrivada BIT NOT NULL
        CONSTRAINT DF_EspaciosDeportivos_AdministracionPrivada DEFAULT (0);
END;

-- Firma: Codex - 18/04/2026 | Se agrega columna AdministracionPrivada para ocultar espacios del portal publico cuando el negocio lo requiera.
