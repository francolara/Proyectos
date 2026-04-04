-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/04/2026
-- Description:   Agrega campos de ubicacion georreferenciada y fotos en tabla Sedes.
-- Firma:         Codex - 02/04/2026 | Campos de mapa/fotos en dbo.Sedes sin tabla adicional.
-- =============================================

IF COL_LENGTH('dbo.Sedes', 'Latitud') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD Latitud DECIMAL(10,7) NULL;
END;
GO

IF COL_LENGTH('dbo.Sedes', 'Longitud') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD Longitud DECIMAL(10,7) NULL;
END;
GO

IF COL_LENGTH('dbo.Sedes', 'GooglePlaceId') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD GooglePlaceId NVARCHAR(200) NULL;
END;
GO

IF COL_LENGTH('dbo.Sedes', 'GoogleMapsUrl') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD GoogleMapsUrl NVARCHAR(500) NULL;
END;
GO

IF COL_LENGTH('dbo.Sedes', 'FotoPrincipalUrl') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD FotoPrincipalUrl NVARCHAR(500) NULL;
END;
GO

IF COL_LENGTH('dbo.Sedes', 'FotosUrlsCsv') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD FotosUrlsCsv NVARCHAR(MAX) NULL;
END;
GO
