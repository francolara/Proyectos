-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Tabla maestra de distritos y codigo ubigeo SUNAT.
-- =============================================

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

    ALTER TABLE dbo.UbigeoDistritos
        ADD CONSTRAINT FK_UbigeoDistritos_UbigeoDepartamentos
        FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento);

    ALTER TABLE dbo.UbigeoDistritos
        ADD CONSTRAINT FK_UbigeoDistritos_UbigeoProvincias
        FOREIGN KEY (CodigoProvincia) REFERENCES dbo.UbigeoProvincias (CodigoProvincia);
END;
