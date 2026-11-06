-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo de FralseCont y direcciones de clientes/tiendas.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UbigeoDepartamentos',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDepartamentos
    (
        CodigoDepartamento CHAR(2) NOT NULL CONSTRAINT PK_UbigeoDepartamentos PRIMARY KEY,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT(1)
    );
END;
GO
