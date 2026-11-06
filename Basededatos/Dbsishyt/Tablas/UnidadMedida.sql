-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define UnidadMedida para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UnidadMedida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UnidadMedida
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_UnidadMedida PRIMARY KEY,
        Codigo NVARCHAR(30) NOT NULL CONSTRAINT UQ_UnidadMedida_Codigo UNIQUE,
        Descripcion NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_UnidadMedida_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_UnidadMedida_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_UnidadMedida_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
