-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ConfiguracionHYT para la base inicial de HYT.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Código de ubigeo de la dirección de HYT.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ConfiguracionHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConfiguracionHYT
    (
        Id TINYINT NOT NULL CONSTRAINT PK_ConfiguracionHYT PRIMARY KEY,
        Ruc NVARCHAR(20) NOT NULL,
        RazonSocial NVARCHAR(200) NOT NULL,
        Direccion NVARCHAR(300) NULL,
        CodigoUbigeo CHAR(6) NULL,
        Telefono NVARCHAR(50) NULL,
        Correo NVARCHAR(256) NULL,
        NombreFirmante NVARCHAR(150) NULL,
        CargoFirmante NVARCHAR(150) NULL,
        CONSTRAINT CK_ConfiguracionHYT_Unica CHECK (Id = 1),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ConfiguracionHYT_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ConfiguracionHYT_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ConfiguracionHYT_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
