-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.ProveedorHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ProveedorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ProveedorHYT
    (
        ProveedorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ProveedorHYT PRIMARY KEY,
        Ruc VARCHAR(11) NOT NULL CONSTRAINT UQ_ProveedorHYT_Ruc UNIQUE,
        RazonSocial NVARCHAR(180) NOT NULL,
        Banco NVARCHAR(80) NULL,
        NumeroCuenta NVARCHAR(34) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_ProveedorHYT_Activo DEFAULT (1),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_ProveedorHYT_Ruc CHECK (LEN(Ruc) = 11 AND Ruc NOT LIKE '%[^0-9]%'),
        CONSTRAINT CK_ProveedorHYT_Cuenta CHECK ((Banco IS NULL AND NumeroCuenta IS NULL)
            OR (Banco IS NOT NULL AND NumeroCuenta IS NOT NULL AND LEN(Banco) > 0 AND LEN(NumeroCuenta) > 0))
    );
END;
GO
