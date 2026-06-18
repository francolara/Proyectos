-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Tabla maestra de provincias relacionadas al ubigeo SUNAT.
-- =============================================

IF OBJECT_ID(N'dbo.UbigeoProvincias', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoProvincias
    (
        CodigoProvincia CHAR(4) NOT NULL CONSTRAINT PK_UbigeoProvincias PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        Nombre NVARCHAR(100) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_UbigeoProvincias_Activo DEFAULT (1)
    );

    ALTER TABLE dbo.UbigeoProvincias
        ADD CONSTRAINT FK_UbigeoProvincias_UbigeoDepartamentos
        FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento);
END;
