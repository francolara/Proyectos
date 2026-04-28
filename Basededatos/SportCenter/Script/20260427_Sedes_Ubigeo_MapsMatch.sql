USE [DbSportCenter]
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Agrega CodigoUbigeo en sedes, crea tabla UbigeoMapsMatch y realiza precarga inicial desde maestro SUNAT.
-- Firma:         Codex - 27/04/2026 | Script incremental para soportar filtros Home por ubigeo de sede y match Google->SUNAT.
-- =============================================

IF COL_LENGTH('dbo.Sedes', 'CodigoUbigeo') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
    ADD CodigoUbigeo CHAR(6) NULL;
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = 'FK_Sedes_UbigeoDistritos_CodigoUbigeo'
      AND parent_object_id = OBJECT_ID('dbo.Sedes')
)
BEGIN
    ALTER TABLE dbo.Sedes WITH CHECK
    ADD CONSTRAINT FK_Sedes_UbigeoDistritos_CodigoUbigeo
    FOREIGN KEY (CodigoUbigeo) REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_Sedes_CodigoUbigeo'
      AND object_id = OBJECT_ID('dbo.Sedes')
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_Sedes_CodigoUbigeo
    ON dbo.Sedes(CodigoUbigeo);
END
GO

IF OBJECT_ID('dbo.UbigeoMapsMatch', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoMapsMatch
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_UbigeoMapsMatch PRIMARY KEY,
        CountryCode CHAR(2) NOT NULL CONSTRAINT DF_UbigeoMapsMatch_CountryCode DEFAULT ('PE'),
        GooglePlaceId NVARCHAR(200) NULL,
        GoogleDepartamento NVARCHAR(120) NULL,
        GoogleProvincia NVARCHAR(120) NULL,
        GoogleDistrito NVARCHAR(120) NULL,
        CodigoUbigeo CHAR(6) NOT NULL,
        EsManual BIT NOT NULL CONSTRAINT DF_UbigeoMapsMatch_EsManual DEFAULT (0),
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoMapsMatch_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_UbigeoMapsMatch_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        GoogleDepartamentoNorm AS UPPER(LTRIM(RTRIM(GoogleDepartamento))) PERSISTED,
        GoogleProvinciaNorm AS UPPER(LTRIM(RTRIM(GoogleProvincia))) PERSISTED,
        GoogleDistritoNorm AS UPPER(LTRIM(RTRIM(GoogleDistrito))) PERSISTED
    );
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = 'FK_UbigeoMapsMatch_UbigeoDistritos_CodigoUbigeo'
      AND parent_object_id = OBJECT_ID('dbo.UbigeoMapsMatch')
)
BEGIN
    ALTER TABLE dbo.UbigeoMapsMatch WITH CHECK
    ADD CONSTRAINT FK_UbigeoMapsMatch_UbigeoDistritos_CodigoUbigeo
    FOREIGN KEY (CodigoUbigeo) REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
END
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE name = 'UX_UbigeoMapsMatch_GooglePlaceId'
      AND object_id = OBJECT_ID('dbo.UbigeoMapsMatch')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_UbigeoMapsMatch_GooglePlaceId
    ON dbo.UbigeoMapsMatch(GooglePlaceId)
    WHERE GooglePlaceId IS NOT NULL;
END
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE name = 'UX_UbigeoMapsMatch_Texto'
      AND object_id = OBJECT_ID('dbo.UbigeoMapsMatch')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_UbigeoMapsMatch_Texto
    ON dbo.UbigeoMapsMatch(CountryCode, GoogleDepartamentoNorm, GoogleProvinciaNorm, GoogleDistritoNorm)
    WHERE GoogleDepartamentoNorm IS NOT NULL
      AND GoogleProvinciaNorm IS NOT NULL
      AND GoogleDistritoNorm IS NOT NULL;
END
GO

UPDATE s
SET s.CodigoUbigeo = n.CodigoUbigeo
FROM dbo.Sedes s
INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
WHERE s.CodigoUbigeo IS NULL
  AND n.CodigoUbigeo IS NOT NULL;
GO

INSERT INTO dbo.UbigeoMapsMatch
(
    CountryCode,
    GooglePlaceId,
    GoogleDepartamento,
    GoogleProvincia,
    GoogleDistrito,
    CodigoUbigeo,
    EsManual,
    Activo,
    UsuarioCreacion
)
SELECT
    'PE',
    NULL,
    dep.Nombre,
    prov.Nombre,
    dist.Nombre,
    dist.CodigoUbigeo,
    1,
    1,
    'seed'
FROM dbo.UbigeoDistritos dist
INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = prov.CodigoDepartamento
WHERE NOT EXISTS
(
    SELECT 1
    FROM dbo.UbigeoMapsMatch m
    WHERE m.CountryCode = 'PE'
      AND m.GoogleDepartamentoNorm = UPPER(LTRIM(RTRIM(dep.Nombre)))
      AND m.GoogleProvinciaNorm = UPPER(LTRIM(RTRIM(prov.Nombre)))
      AND m.GoogleDistritoNorm = UPPER(LTRIM(RTRIM(dist.Nombre)))
);
GO
