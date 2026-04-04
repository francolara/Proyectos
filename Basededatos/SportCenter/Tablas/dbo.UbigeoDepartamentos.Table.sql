USE [DbSportCenter]
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   04/04/2026
-- Firma Codex:   Creacion de tabla maestra de departamentos para ubigeo SUNAT.
-- =============================================

IF OBJECT_ID(N'dbo.UbigeoDepartamentos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDepartamentos
    (
        CodigoDepartamento CHAR(2) NOT NULL,
        Nombre NVARCHAR(100) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoDepartamentos_Activo DEFAULT (1),
        CONSTRAINT PK_UbigeoDepartamentos PRIMARY KEY CLUSTERED (CodigoDepartamento)
    );
END;
GO
