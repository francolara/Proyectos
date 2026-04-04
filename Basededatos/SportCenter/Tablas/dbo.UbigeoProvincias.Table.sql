USE [DbSportCenter]
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   04/04/2026
-- Firma Codex:   Creacion de tabla maestra de provincias relacionada a departamentos.
-- =============================================

IF OBJECT_ID(N'dbo.UbigeoProvincias', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoProvincias
    (
        CodigoProvincia CHAR(4) NOT NULL,
        CodigoDepartamento CHAR(2) NOT NULL,
        Nombre NVARCHAR(100) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoProvincias_Activo DEFAULT (1),
        CONSTRAINT PK_UbigeoProvincias PRIMARY KEY CLUSTERED (CodigoProvincia)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UbigeoProvincias_UbigeoDepartamentos')
BEGIN
    ALTER TABLE dbo.UbigeoProvincias
    WITH CHECK ADD CONSTRAINT FK_UbigeoProvincias_UbigeoDepartamentos
    FOREIGN KEY (CodigoDepartamento)
    REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_UbigeoProvincias_CodigoDepartamento' AND object_id = OBJECT_ID(N'dbo.UbigeoProvincias'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_UbigeoProvincias_CodigoDepartamento
    ON dbo.UbigeoProvincias (CodigoDepartamento);
END;
GO
