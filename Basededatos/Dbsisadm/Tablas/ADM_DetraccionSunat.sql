-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Maestro general SUNAT de codigos de detraccion SPOT con porcentaje base para compras y futuras operaciones.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_DetraccionSunat', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_DetraccionSunat
    (
        IdDetraccionSunat INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_DetraccionSunat PRIMARY KEY,
        CodigoSunat VARCHAR(3) NOT NULL,
        Descripcion NVARCHAR(250) NOT NULL,
        Porcentaje DECIMAL(7,4) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_ADM_DetraccionSunat_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.ADM_DetraccionSunat
        ADD CONSTRAINT UQ_ADM_DetraccionSunat_CodigoSunat UNIQUE (CodigoSunat);

    ALTER TABLE dbo.ADM_DetraccionSunat
        ADD CONSTRAINT CK_ADM_DetraccionSunat_Porcentaje
            CHECK (Porcentaje >= 0 AND Porcentaje <= 100);
END;
