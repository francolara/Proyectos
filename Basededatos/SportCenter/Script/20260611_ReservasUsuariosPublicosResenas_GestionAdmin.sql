GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/06/2026
-- Firma Codex:   Agrega campos de gestion administrativa para moderar y responder reseñas publicas por espacio deportivo.
-- =============================================

IF COL_LENGTH('dbo.ReservasUsuariosPublicosResenas', 'Activo') IS NULL
BEGIN
    ALTER TABLE dbo.ReservasUsuariosPublicosResenas
    ADD Activo BIT NOT NULL
        CONSTRAINT DF_ReservasUsuariosPublicosResenas_Activo DEFAULT ((1));
END;
GO

IF COL_LENGTH('dbo.ReservasUsuariosPublicosResenas', 'Respuesta') IS NULL
BEGIN
    ALTER TABLE dbo.ReservasUsuariosPublicosResenas
    ADD Respuesta NVARCHAR(800) NULL;
END;
GO

UPDATE dbo.ReservasUsuariosPublicosResenas
SET Activo = 1
WHERE Activo IS NULL;
GO
