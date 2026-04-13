/*
Firma: Codex - 10/04/2026
Descripcion: Crea/ajusta tabla ParametrosGlobales con NombreParametro e inserta parametro VALIDA_MONTO_BSINDOC.
*/
USE [DbSportCenter]
GO

IF OBJECT_ID(N'dbo.ParametrosGlobales', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ParametrosGlobales
    (
        ParametroId INT IDENTITY(1,1) NOT NULL,
        NombreParametro NVARCHAR(100) NOT NULL,
        Descripcion NVARCHAR(500) NOT NULL,
        ValorParametro NVARCHAR(100) NOT NULL,
        CONSTRAINT PK_ParametrosGlobales PRIMARY KEY CLUSTERED (ParametroId ASC),
        CONSTRAINT UQ_ParametrosGlobales_Descripcion UNIQUE (Descripcion),
        CONSTRAINT UQ_ParametrosGlobales_NombreParametro UNIQUE (NombreParametro)
    );
END
GO

IF COL_LENGTH('dbo.ParametrosGlobales', 'NombreParametro') IS NULL
BEGIN
    ALTER TABLE dbo.ParametrosGlobales
    ADD NombreParametro NVARCHAR(100) NULL;

    UPDATE dbo.ParametrosGlobales
    SET NombreParametro = CONCAT(N'PARAM_', ParametroId)
    WHERE NombreParametro IS NULL;

    ALTER TABLE dbo.ParametrosGlobales
    ALTER COLUMN NombreParametro NVARCHAR(100) NOT NULL;
END
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes i
    WHERE i.object_id = OBJECT_ID(N'dbo.ParametrosGlobales')
      AND i.name = N'UQ_ParametrosGlobales_NombreParametro'
)
BEGIN
    ALTER TABLE dbo.ParametrosGlobales
    ADD CONSTRAINT UQ_ParametrosGlobales_NombreParametro UNIQUE (NombreParametro);
END
GO

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ParametrosGlobales p
    WHERE p.NombreParametro = N'VALIDA_MONTO_BSINDOC'
)
BEGIN
    INSERT INTO dbo.ParametrosGlobales (NombreParametro, Descripcion, ValorParametro)
    VALUES (N'VALIDA_MONTO_BSINDOC', N'Monto Maximo para atencion de boletas sin DOC', N'700');
END
ELSE
BEGIN
    UPDATE p
    SET p.Descripcion = N'Monto Maximo para atencion de boletas sin DOC',
        p.ValorParametro = N'700'
    FROM dbo.ParametrosGlobales p
    WHERE p.NombreParametro = N'VALIDA_MONTO_BSINDOC';
END
GO
