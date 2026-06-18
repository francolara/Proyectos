-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Reestructura personas por empresa e incorpora maestros SUNAT de ubigeo y tipo de documento.
-- =============================================

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID(N'dbo.UbigeoDepartamentos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDepartamentos
    (
        CodigoDepartamento CHAR(2) NOT NULL CONSTRAINT PK_UbigeoDepartamentos PRIMARY KEY,
        Nombre NVARCHAR(100) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoDepartamentos_Activo DEFAULT (1)
    );
END;
GO

IF OBJECT_ID(N'dbo.UbigeoProvincias', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoProvincias
    (
        CodigoProvincia CHAR(4) NOT NULL CONSTRAINT PK_UbigeoProvincias PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        Nombre NVARCHAR(100) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoProvincias_Activo DEFAULT (1)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UbigeoProvincias_UbigeoDepartamentos')
BEGIN
    ALTER TABLE dbo.UbigeoProvincias
        ADD CONSTRAINT FK_UbigeoProvincias_UbigeoDepartamentos
        FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento);
END;
GO

IF OBJECT_ID(N'dbo.UbigeoDistritos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDistritos
    (
        CodigoUbigeo CHAR(6) NOT NULL CONSTRAINT PK_UbigeoDistritos PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        CodigoProvincia CHAR(4) NOT NULL,
        Nombre NVARCHAR(120) NOT NULL,
        Zona NVARCHAR(20) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoDistritos_Activo DEFAULT (1)
    );
END;
GO

IF COL_LENGTH(N'dbo.UbigeoDistritos', N'Zona') IS NULL
BEGIN
    ALTER TABLE dbo.UbigeoDistritos
        ADD Zona NVARCHAR(20) NULL;
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UbigeoDistritos_UbigeoDepartamentos')
BEGIN
    ALTER TABLE dbo.UbigeoDistritos
        ADD CONSTRAINT FK_UbigeoDistritos_UbigeoDepartamentos
        FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UbigeoDistritos_UbigeoProvincias')
BEGIN
    ALTER TABLE dbo.UbigeoDistritos
        ADD CONSTRAINT FK_UbigeoDistritos_UbigeoProvincias
        FOREIGN KEY (CodigoProvincia) REFERENCES dbo.UbigeoProvincias (CodigoProvincia);
END;
GO

IF OBJECT_ID(N'dbo.TiposDocumentoIdentidadSunat', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TiposDocumentoIdentidadSunat
    (
        CodigoSunat NVARCHAR(2) NOT NULL CONSTRAINT PK_TiposDocumentoIdentidadSunat PRIMARY KEY,
        CodigoInterno NVARCHAR(20) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_TiposDocumentoIdentidadSunat_Activo DEFAULT (1),
        Orden TINYINT NOT NULL,
        CONSTRAINT UQ_TiposDocumentoIdentidadSunat_CodigoInterno UNIQUE (CodigoInterno)
    );
END;
GO

MERGE dbo.TiposDocumentoIdentidadSunat AS destino
USING
(
    VALUES
        (N'0', N'OTRO', N'Otro documento de identidad', CAST(1 AS BIT), CAST(1 AS TINYINT)),
        (N'1', N'DNI', N'Documento nacional de identidad', CAST(1 AS BIT), CAST(2 AS TINYINT)),
        (N'4', N'CE', N'Carnet de extranjeria', CAST(1 AS BIT), CAST(3 AS TINYINT)),
        (N'6', N'RUC', N'Registro unico de contribuyentes', CAST(1 AS BIT), CAST(4 AS TINYINT)),
        (N'7', N'PASAPORTE', N'Pasaporte', CAST(1 AS BIT), CAST(5 AS TINYINT)),
        (N'A', N'CED_DIPLOMATICA', N'Cedula diplomatica de identidad', CAST(1 AS BIT), CAST(6 AS TINYINT))
) AS origen (CodigoSunat, CodigoInterno, Nombre, Activo, Orden)
ON destino.CodigoSunat = origen.CodigoSunat
WHEN MATCHED THEN
    UPDATE
    SET destino.CodigoInterno = origen.CodigoInterno,
        destino.Nombre = origen.Nombre,
        destino.Activo = origen.Activo,
        destino.Orden = origen.Orden
WHEN NOT MATCHED THEN
    INSERT (CodigoSunat, CodigoInterno, Nombre, Activo, Orden)
    VALUES (origen.CodigoSunat, origen.CodigoInterno, origen.Nombre, origen.Activo, origen.Orden);
GO

IF COL_LENGTH(N'dbo.ADM_Persona', N'IdEmpresa') IS NULL
BEGIN
    ALTER TABLE dbo.ADM_Persona
        ADD IdEmpresa INT NULL;
END;
GO

IF COL_LENGTH(N'dbo.ADM_Persona', N'CodigoUbigeo') IS NULL
BEGIN
    ALTER TABLE dbo.ADM_Persona
        ADD CodigoUbigeo CHAR(6) NULL;
END;
GO

;WITH EmpresasRelacionadas AS
(
    SELECT c.IdPersona, c.IdEmpresa
    FROM dbo.ADM_Cliente AS c
    UNION
    SELECT p.IdPersona, p.IdEmpresa
    FROM dbo.ADM_Proveedor AS p
),
EmpresasOrdenadas AS
(
    SELECT
        er.IdPersona,
        er.IdEmpresa,
        ROW_NUMBER() OVER (PARTITION BY er.IdPersona ORDER BY er.IdEmpresa) AS NumeroEmpresa
    FROM EmpresasRelacionadas AS er
)
SELECT
    eo.IdPersona,
    eo.IdEmpresa
INTO #PersonasEmpresaDuplicada
FROM EmpresasOrdenadas AS eo
WHERE eo.NumeroEmpresa > 1;
GO

IF OBJECT_ID(N'tempdb..#PersonasEmpresaDuplicada', N'U') IS NOT NULL
BEGIN
    DECLARE @MapeoPersona TABLE
    (
        IdPersonaOriginal INT NOT NULL,
        IdEmpresa INT NOT NULL,
        IdPersonaNueva INT NOT NULL
    );

    IF EXISTS (SELECT 1 FROM #PersonasEmpresaDuplicada)
    BEGIN
        INSERT INTO dbo.ADM_Persona
        (
            IdEmpresa,
            TipoPersona,
            TipoDocumento,
            NumeroDocumento,
            ApellidoPaterno,
            ApellidoMaterno,
            Nombres,
            RazonSocial,
            CorreoElectronico,
            Telefono,
            Direccion,
            CodigoUbigeo,
            Estado,
            FechaRegistro,
            UsuarioRegistro
        )
        OUTPUT duplicadas.IdPersona, duplicadas.IdEmpresa, inserted.IdPersona
        INTO @MapeoPersona (IdPersonaOriginal, IdEmpresa, IdPersonaNueva)
        SELECT
            duplicadas.IdEmpresa,
            persona.TipoPersona,
            persona.TipoDocumento,
            persona.NumeroDocumento,
            persona.ApellidoPaterno,
            persona.ApellidoMaterno,
            persona.Nombres,
            persona.RazonSocial,
            persona.CorreoElectronico,
            persona.Telefono,
            persona.Direccion,
            persona.CodigoUbigeo,
            persona.Estado,
            persona.FechaRegistro,
            persona.UsuarioRegistro
        FROM #PersonasEmpresaDuplicada AS duplicadas
        INNER JOIN dbo.ADM_Persona AS persona
            ON persona.IdPersona = duplicadas.IdPersona;

        UPDATE c
        SET c.IdPersona = m.IdPersonaNueva
        FROM dbo.ADM_Cliente AS c
        INNER JOIN @MapeoPersona AS m
            ON m.IdPersonaOriginal = c.IdPersona
           AND m.IdEmpresa = c.IdEmpresa;

        UPDATE p
        SET p.IdPersona = m.IdPersonaNueva
        FROM dbo.ADM_Proveedor AS p
        INNER JOIN @MapeoPersona AS m
            ON m.IdPersonaOriginal = p.IdPersona
           AND m.IdEmpresa = p.IdEmpresa;
    END;
END;
GO

DROP TABLE IF EXISTS #PersonasEmpresaDuplicada;
GO

;WITH EmpresasBase AS
(
    SELECT c.IdPersona, MIN(c.IdEmpresa) AS IdEmpresa
    FROM
    (
        SELECT IdPersona, IdEmpresa
        FROM dbo.ADM_Cliente
        UNION
        SELECT IdPersona, IdEmpresa
        FROM dbo.ADM_Proveedor
    ) AS c
    GROUP BY
        c.IdPersona
)
UPDATE p
SET p.IdEmpresa = e.IdEmpresa
FROM dbo.ADM_Persona AS p
INNER JOIN EmpresasBase AS e
    ON e.IdPersona = p.IdPersona
WHERE p.IdEmpresa IS NULL;
GO

DECLARE @IdEmpresaUnica INT = NULL;

IF (SELECT COUNT(1) FROM dbo.SEG_Empresa) = 1
BEGIN
    SELECT @IdEmpresaUnica = MIN(e.IdEmpresa)
    FROM dbo.SEG_Empresa AS e;

    UPDATE dbo.ADM_Persona
    SET IdEmpresa = @IdEmpresaUnica
    WHERE IdEmpresa IS NULL;
END;
GO

IF EXISTS
(
    SELECT 1
    FROM dbo.ADM_Persona AS p
    WHERE p.IdEmpresa IS NULL
)
BEGIN
    RAISERROR(N'Existen personas sin empresa asignada. Complete la relacion manualmente antes de continuar.', 16, 1);
END;
GO

IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = N'UQ_ADM_Persona_Documento')
BEGIN
    ALTER TABLE dbo.ADM_Persona DROP CONSTRAINT UQ_ADM_Persona_Documento;
END;
GO

IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = N'UQ_ADM_Persona_EmpresaDocumento')
BEGIN
    ALTER TABLE dbo.ADM_Persona DROP CONSTRAINT UQ_ADM_Persona_EmpresaDocumento;
END;
GO

ALTER TABLE dbo.ADM_Persona
    ALTER COLUMN IdEmpresa INT NOT NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ADM_Persona_SEG_Empresa')
BEGIN
    ALTER TABLE dbo.ADM_Persona
        ADD CONSTRAINT FK_ADM_Persona_SEG_Empresa
        FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = N'UQ_ADM_Persona_EmpresaDocumento')
BEGIN
    ALTER TABLE dbo.ADM_Persona
        ADD CONSTRAINT UQ_ADM_Persona_EmpresaDocumento UNIQUE (IdEmpresa, TipoDocumento, NumeroDocumento);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ADM_Persona_TiposDocumentoIdentidadSunat')
BEGIN
    ALTER TABLE dbo.ADM_Persona
        ADD CONSTRAINT FK_ADM_Persona_TiposDocumentoIdentidadSunat
        FOREIGN KEY (TipoDocumento) REFERENCES dbo.TiposDocumentoIdentidadSunat (CodigoSunat);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ADM_Persona_UbigeoDistritos')
BEGIN
    ALTER TABLE dbo.ADM_Persona
        ADD CONSTRAINT FK_ADM_Persona_UbigeoDistritos
        FOREIGN KEY (CodigoUbigeo) REFERENCES dbo.UbigeoDistritos (CodigoUbigeo);
END;
GO
