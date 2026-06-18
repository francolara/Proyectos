-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Tabla maestra centralizada de tipos de documento de identidad SUNAT.
-- =============================================

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
