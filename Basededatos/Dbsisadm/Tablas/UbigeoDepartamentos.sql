-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Tabla maestra de departamentos para ubigeo SUNAT.
-- =============================================

IF OBJECT_ID(N'dbo.UbigeoDepartamentos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDepartamentos
    (
        CodigoDepartamento CHAR(2) NOT NULL CONSTRAINT PK_UbigeoDepartamentos PRIMARY KEY,
        Nombre NVARCHAR(100) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoDepartamentos_Activo DEFAULT (1)
    );
END;
