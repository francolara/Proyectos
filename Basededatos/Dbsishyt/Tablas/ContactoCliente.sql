-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ContactoCliente para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ContactoCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ContactoCliente
    (
        ContactoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ContactoCliente PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        TiendaId BIGINT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Telefono NVARCHAR(50) NULL,
        Cargo NVARCHAR(100) NULL,
        Correo NVARCHAR(256) NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_Contacto_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT FK_Contacto_Tienda FOREIGN KEY (ClienteId, TiendaId) REFERENCES dbo.Tienda(ClienteId, TiendaId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ContactoCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ContactoCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ContactoCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
