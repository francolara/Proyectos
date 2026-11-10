-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Catalogo de cargos del personal.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.CargoTrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CargoTrabajadorHYT
    (
        CargoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_CargoTrabajadorHYT PRIMARY KEY,
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT UQ_CargoTrabajadorHYT_Nombre UNIQUE,
        Activo BIT NOT NULL,
        VersionFila ROWVERSION NOT NULL
    );
END;
GO
