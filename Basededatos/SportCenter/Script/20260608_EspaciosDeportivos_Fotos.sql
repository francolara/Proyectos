-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/06/2026
-- Description:   Agrega columnas de fotos para espacios deportivos.
-- =============================================

IF COL_LENGTH('dbo.EspaciosDeportivos', 'FotoPrincipalUrl') IS NULL
BEGIN
    ALTER TABLE dbo.EspaciosDeportivos
    ADD FotoPrincipalUrl NVARCHAR(500) NULL;
END;

IF COL_LENGTH('dbo.EspaciosDeportivos', 'FotosUrlsCsv') IS NULL
BEGIN
    ALTER TABLE dbo.EspaciosDeportivos
    ADD FotosUrlsCsv NVARCHAR(MAX) NULL;
END;
