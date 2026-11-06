-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Tienda para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Tienda', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tienda
    (
        TiendaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Tienda PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        Codigo NVARCHAR(40) NOT NULL,
        Nombre NVARCHAR(200) NOT NULL,
        Direccion NVARCHAR(300) NULL,
        CentroSAPFICO NVARCHAR(50) NULL,
        ZonaId INT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT UQ_Tienda_ClienteCodigo UNIQUE (ClienteId, Codigo),
        CONSTRAINT UQ_Tienda_ClienteId UNIQUE (ClienteId, TiendaId),
        CONSTRAINT FK_Tienda_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT FK_Tienda_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Tienda_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Tienda_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Tienda_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
