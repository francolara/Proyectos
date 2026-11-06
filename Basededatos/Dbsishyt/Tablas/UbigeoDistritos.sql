-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo de FralseCont y direcciones de clientes/tiendas.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UbigeoDistritos',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDistritos
    (
        CodigoUbigeo CHAR(6) NOT NULL CONSTRAINT PK_UbigeoDistritos PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        CodigoProvincia CHAR(4) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT(1),
        CONSTRAINT FK_UbigeoDistritos_CodigoDepartamento FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento),
        CONSTRAINT FK_UbigeoDistritos_CodigoProvincia FOREIGN KEY (CodigoProvincia) REFERENCES dbo.UbigeoProvincias (CodigoProvincia)
    );
END;
GO
