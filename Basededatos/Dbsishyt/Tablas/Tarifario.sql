-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Tarifario para la base inicial de HYT.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Identifica tarifario interno de materiales, separado de publicaciones convencionales.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Tarifario', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tarifario
    (
        TarifarioId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Tarifario PRIMARY KEY,
        Codigo NVARCHAR(40) NOT NULL CONSTRAINT UQ_Tarifario_Codigo UNIQUE,
        EsTarifarioMaterial BIT NOT NULL DEFAULT(0),
        Nombre NVARCHAR(200) NOT NULL,
        ClienteId BIGINT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_Tarifario_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Tarifario_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Tarifario_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Tarifario_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
