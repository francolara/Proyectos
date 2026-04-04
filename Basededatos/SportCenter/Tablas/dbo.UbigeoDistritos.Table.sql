USE [DbSportCenter]
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   04/04/2026
-- Firma Codex:   Creacion de tabla maestra de distritos relacionada a provincias y departamentos.
-- =============================================

IF OBJECT_ID(N'dbo.UbigeoDistritos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDistritos
    (
        CodigoUbigeo CHAR(6) NOT NULL,
        CodigoDepartamento CHAR(2) NOT NULL,
        CodigoProvincia CHAR(4) NOT NULL,
        Nombre NVARCHAR(120) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoDistritos_Activo DEFAULT (1),
        CONSTRAINT PK_UbigeoDistritos PRIMARY KEY CLUSTERED (CodigoUbigeo)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UbigeoDistritos_UbigeoDepartamentos')
BEGIN
    ALTER TABLE dbo.UbigeoDistritos
    WITH CHECK ADD CONSTRAINT FK_UbigeoDistritos_UbigeoDepartamentos
    FOREIGN KEY (CodigoDepartamento)
    REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UbigeoDistritos_UbigeoProvincias')
BEGIN
    ALTER TABLE dbo.UbigeoDistritos
    WITH CHECK ADD CONSTRAINT FK_UbigeoDistritos_UbigeoProvincias
    FOREIGN KEY (CodigoProvincia)
    REFERENCES dbo.UbigeoProvincias (CodigoProvincia);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_UbigeoDistritos_CodigoProvincia' AND object_id = OBJECT_ID(N'dbo.UbigeoDistritos'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_UbigeoDistritos_CodigoProvincia
    ON dbo.UbigeoDistritos (CodigoProvincia);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_UbigeoDistritos_CodigoDepartamento' AND object_id = OBJECT_ID(N'dbo.UbigeoDistritos'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_UbigeoDistritos_CodigoDepartamento
    ON dbo.UbigeoDistritos (CodigoDepartamento);
END;
GO
