-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Especialidad para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Especialidad', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Especialidad
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_Especialidad PRIMARY KEY,
        Codigo NVARCHAR(30) NOT NULL CONSTRAINT UQ_Especialidad_Codigo UNIQUE,
        Descripcion NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Especialidad_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Especialidad_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Especialidad_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
