-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ArchivoAdjunto para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ArchivoAdjunto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ArchivoAdjunto
    (
        ArchivoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ArchivoAdjunto PRIMARY KEY,
        NombreOriginal NVARCHAR(255) NOT NULL,
        ClaveAlmacenamiento NVARCHAR(500) NOT NULL CONSTRAINT UQ_Archivo_Clave UNIQUE,
        TipoContenido NVARCHAR(150) NOT NULL,
        TamanoBytes BIGINT NOT NULL,
        SHA256 CHAR(64) NOT NULL,
        Retirado BIT NOT NULL DEFAULT 0,
        CONSTRAINT CK_Archivo_Tamano CHECK (TamanoBytes > 0),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ArchivoAdjunto_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ArchivoAdjunto_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ArchivoAdjunto_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
