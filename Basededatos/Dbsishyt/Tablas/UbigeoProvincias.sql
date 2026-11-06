-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo de FralseCont y direcciones de clientes/tiendas.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UbigeoProvincias',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoProvincias
    (
        CodigoProvincia CHAR(4) NOT NULL CONSTRAINT PK_UbigeoProvincias PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT(1),
        CONSTRAINT FK_UbigeoProvincias_CodigoDepartamento FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento)
    );
END;
GO
